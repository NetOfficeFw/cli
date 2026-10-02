#include "pch.h"
#include "GrpcServer.h"

#include <algorithm>
#include <chrono>
#include <climits>
#include <condition_variable>
#include <cstdio>
#include <cwchar>
#include <exception>
#include <new>
#include <utility>

namespace
{
	constexpr UINT GrpcDispatchMessage = WM_APP + 1;
	constexpr int ServerPort = 50051;
	constexpr auto CancellationPollInterval = std::chrono::milliseconds(20);

	grpc::Status ErrorStatus(grpc::StatusCode code, HRESULT hr,
		const std::string &operation, const std::string &details = std::string())
	{
		char hresult[32];
		sprintf_s(hresult, " (HRESULT 0x%08lX)", static_cast<unsigned long>(hr));
		std::string message = operation + " failed" + hresult;
		if (!details.empty())
			message += ": " + details;
		return grpc::Status(code, message);
	}

	grpc::Status StoppedStatus()
	{
		return ErrorStatus(grpc::StatusCode::UNAVAILABLE,
			HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED), "PowerPoint addin connection");
	}

	grpc::Status CancellationStatus(std::chrono::system_clock::time_point deadline)
	{
		if (std::chrono::system_clock::now() >= deadline)
			return ErrorStatus(grpc::StatusCode::DEADLINE_EXCEEDED,
				HRESULT_FROM_WIN32(ERROR_TIMEOUT), "PowerPoint request deadline");
		return ErrorStatus(grpc::StatusCode::CANCELLED,
			HRESULT_FROM_WIN32(ERROR_CANCELLED), "PowerPoint request");
	}

	HRESULT Utf8ToVariant(const std::string &text, ATL::CComVariant &result)
	{
		if (text.size() > INT_MAX)
			return E_INVALIDARG;

		int length = 0;
		if (!text.empty())
		{
			length = MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
				text.data(), static_cast<int>(text.size()), nullptr, 0);
			if (length == 0)
				return HRESULT_FROM_WIN32(GetLastError());
		}

		result.bstrVal = SysAllocStringLen(nullptr, length);
		if (result.bstrVal == nullptr)
			return E_OUTOFMEMORY;
		result.vt = VT_BSTR;
		if (length != 0 && MultiByteToWideChar(CP_UTF8, MB_ERR_INVALID_CHARS,
			text.data(), static_cast<int>(text.size()), result.bstrVal, length) != length)
			return HRESULT_FROM_WIN32(GetLastError());
		return S_OK;
	}

	HRESULT BstrToUtf8(BSTR text, std::string &result)
	{
		UINT length = SysStringLen(text);
		if (length > INT_MAX)
			return E_INVALIDARG;
		if (length == 0)
		{
			result.clear();
			return S_OK;
		}

		int bytes = WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS,
			text, static_cast<int>(length), nullptr, 0, nullptr, nullptr);
		if (bytes == 0)
			return HRESULT_FROM_WIN32(GetLastError());
		result.resize(bytes);
		if (WideCharToMultiByte(CP_UTF8, WC_ERR_INVALID_CHARS, text,
			static_cast<int>(length), result.data(), bytes, nullptr, nullptr) != bytes)
			return HRESULT_FROM_WIN32(GetLastError());
		return S_OK;
	}

	std::string MemberName(const wchar_t *pName)
	{
		// Automation member names in this bridge are ASCII, not user text.
		std::string name;
		while (*pName != L'\0')
			name.push_back(static_cast<char>(*pName++));
		return "PowerPoint." + name;
	}
}

struct GrpcServer::PendingCall
{
	enum class Operation { NewPresentation, SetSlideTitle };

	explicit PendingCall(Operation operation) : operation(operation) {}

	Operation operation;
	std::string text;
	int slideIndex = 0;
	netoffice::PresentationReply presentationReply;

	std::mutex mutex;
	std::condition_variable completed;
	// This pointer is only inspected under mutex and is cleared before the RPC returns.
	grpc::ServerContext *pContext = nullptr;
	std::chrono::system_clock::time_point deadline;
	bool finished = false;
	// Preallocate the shutdown result so disconnect itself need not allocate.
	grpc::Status status = StoppedStatus();

	void FinishForShutdown()
	{
		{
			std::lock_guard<std::mutex> lock(mutex);
			finished = true;
		}
		completed.notify_all();
	}
};

GrpcServer::~GrpcServer()
{
	Stop();
}

HRESULT GrpcServer::Start(IDispatch *pApplication)
{
	if (pApplication == nullptr)
		return E_POINTER;

	APTTYPE apartment;
	APTTYPEQUALIFIER qualifier;
	HRESULT hr = CoGetApartmentType(&apartment, &qualifier);
	if (FAILED(hr))
		return hr;
	if (apartment != APTTYPE_STA && apartment != APTTYPE_MAINSTA)
		return RPC_E_WRONG_THREAD;

	m_ownerThreadId = GetCurrentThreadId();
	m_pApplication = pApplication;
	m_hInstance = ATL::_AtlBaseModule.GetModuleInstance();

	wchar_t className[80];
	swprintf_s(className, L"NetOffice.PowerPoint.Grpc.%p", static_cast<void *>(this));
	WNDCLASSEXW windowClass = {};
	windowClass.cbSize = sizeof(windowClass);
	windowClass.lpfnWndProc = WindowProc;
	windowClass.hInstance = m_hInstance;
	windowClass.lpszClassName = className;
	m_windowClass = RegisterClassExW(&windowClass);
	if (m_windowClass == 0)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		Stop();
		return hr;
	}

	m_hWnd = CreateWindowExW(0, MAKEINTATOM(m_windowClass), L"", 0,
		0, 0, 0, 0, HWND_MESSAGE, nullptr, m_hInstance, this);
	if (m_hWnd == nullptr)
	{
		hr = HRESULT_FROM_WIN32(GetLastError());
		Stop();
		return hr;
	}

	try
	{
		{
			std::lock_guard<std::mutex> lock(m_mutex);
			m_stopping = false;
		}
		grpc::ServerBuilder builder;
		// A second addin connection must fail, not share the existing listener.
		builder.AddChannelArgument("grpc.so_reuseport", 0);
		int selectedPort = 0;
		builder.AddListeningPort("127.0.0.1:50051", grpc::InsecureServerCredentials(), &selectedPort);
		builder.RegisterService(this);
		m_server = builder.BuildAndStart();
		if (!m_server || selectedPort != ServerPort)
		{
			Stop();
			return HRESULT_FROM_WIN32(ERROR_ADDRESS_ALREADY_ASSOCIATED);
		}
	}
	catch (const std::bad_alloc &)
	{
		Stop();
		return E_OUTOFMEMORY;
	}
	catch (const std::exception &)
	{
		Stop();
		return E_FAIL;
	}
	return S_OK;
}

void GrpcServer::Stop()
{
	ATLASSERT(m_ownerThreadId == 0 || m_ownerThreadId == GetCurrentThreadId());

	HWND hWnd;
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_stopping = true;
		// Wake handlers before Shutdown, including an active call whose COM
		// Invoke may have reentered an Office disconnect callback.
		for (const auto &call : m_pendingCalls)
			call->FinishForShutdown();
		m_pendingCalls.clear();
		if (m_activeCall)
			m_activeCall->FinishForShutdown();
		hWnd = m_hWnd;
		m_hWnd = nullptr;
	}

	if (hWnd != nullptr)
	{
		SetWindowLongPtrW(hWnd, GWLP_USERDATA, 0);
		DestroyWindow(hWnd);
	}
	if (m_server)
	{
		m_server->Shutdown(std::chrono::system_clock::now());
		m_server->Wait();
		m_server.reset();
	}
	if (m_windowClass != 0)
	{
		UnregisterClassW(MAKEINTATOM(m_windowClass), m_hInstance);
		m_windowClass = 0;
	}
	m_pApplication.Release();
}

bool GrpcServer::IsStopping()
{
	std::lock_guard<std::mutex> lock(m_mutex);
	return m_stopping;
}

grpc::Status GrpcServer::GetStatus(grpc::ServerContext *pContext,
	const netoffice::Empty *, netoffice::StatusReply *pReply)
{
	if (pContext->IsCancelled())
		return CancellationStatus(pContext->deadline());
	std::lock_guard<std::mutex> lock(m_mutex);
	if (m_stopping)
		return StoppedStatus();
	pReply->set_process_id(GetCurrentProcessId());
	return grpc::Status::OK;
}

grpc::Status GrpcServer::NewPresentation(grpc::ServerContext *pContext,
	const netoffice::NewPresentationRequest *pRequest, netoffice::PresentationReply *pReply)
{
	try
	{
		auto call = std::make_shared<PendingCall>(PendingCall::Operation::NewPresentation);
		call->text = pRequest->title();
		grpc::Status status = Dispatch(call, pContext);
		if (status.ok())
			*pReply = std::move(call->presentationReply);
		return status;
	}
	catch (const std::bad_alloc &)
	{
		return ErrorStatus(grpc::StatusCode::RESOURCE_EXHAUSTED, E_OUTOFMEMORY, "NewPresentation");
	}
	catch (const std::exception &error)
	{
		return ErrorStatus(grpc::StatusCode::INTERNAL, E_FAIL, "NewPresentation", error.what());
	}
}

grpc::Status GrpcServer::SetSlideTitle(grpc::ServerContext *pContext,
	const netoffice::SetSlideTitleRequest *pRequest, netoffice::Empty *)
{
	if (pRequest->slide_index() < 1)
		return ErrorStatus(grpc::StatusCode::INVALID_ARGUMENT, E_INVALIDARG,
			"SetSlideTitle", "slide_index is one-based and must be at least 1");
	try
	{
		auto call = std::make_shared<PendingCall>(PendingCall::Operation::SetSlideTitle);
		call->slideIndex = pRequest->slide_index();
		call->text = pRequest->text();
		return Dispatch(call, pContext);
	}
	catch (const std::bad_alloc &)
	{
		return ErrorStatus(grpc::StatusCode::RESOURCE_EXHAUSTED, E_OUTOFMEMORY, "SetSlideTitle");
	}
	catch (const std::exception &error)
	{
		return ErrorStatus(grpc::StatusCode::INTERNAL, E_FAIL, "SetSlideTitle", error.what());
	}
}

grpc::Status GrpcServer::Dispatch(const std::shared_ptr<PendingCall> &call,
	grpc::ServerContext *pContext)
{
	call->pContext = pContext;
	call->deadline = pContext->deadline();
	struct ContextBinding
	{
		PendingCall &call;
		~ContextBinding()
		{
			// Also detach on exceptional exits: queued requests must never
			// outlive the RPC while retaining its ServerContext pointer.
			std::lock_guard<std::mutex> lock(call.mutex);
			call.pContext = nullptr;
			call.finished = true;
		}
	} binding{ *call };
	if (pContext->IsCancelled() || std::chrono::system_clock::now() >= call->deadline)
		return CancellationStatus(call->deadline);

	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (m_stopping)
			return StoppedStatus();
		m_pendingCalls.push_back(call);
		if (!PostMessageW(m_hWnd, GrpcDispatchMessage, 0, 0))
		{
			HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
			m_pendingCalls.pop_back();
			return ErrorStatus(grpc::StatusCode::UNAVAILABLE, hr, "Posting PowerPoint STA request");
		}
	}

	std::unique_lock<std::mutex> lock(call->mutex);
	bool cancelled = false;
	while (!call->finished)
	{
		if (pContext->IsCancelled() || std::chrono::system_clock::now() >= call->deadline)
		{
			call->status = CancellationStatus(call->deadline);
			call->finished = true;
			cancelled = true;
			break;
		}
		call->completed.wait_until(lock,
			(std::min)(call->deadline, std::chrono::system_clock::now() + CancellationPollInterval));
	}
	// A canceled RPC may return while its shared request remains queued or is
	// finishing on the STA. Neither path can subsequently touch ServerContext.
	call->pContext = nullptr;
	grpc::Status status = call->status;
	lock.unlock();
	if (cancelled)
	{
		std::lock_guard<std::mutex> queueLock(m_mutex);
		auto position = std::find(m_pendingCalls.begin(), m_pendingCalls.end(), call);
		if (position != m_pendingCalls.end())
			m_pendingCalls.erase(position);
	}
	return status;
}

LRESULT CALLBACK GrpcServer::WindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_NCCREATE)
	{
		auto pCreate = reinterpret_cast<CREATESTRUCTW *>(lParam);
		SetWindowLongPtrW(hWnd, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(pCreate->lpCreateParams));
	}
	else if (message == GrpcDispatchMessage)
	{
		auto pServer = reinterpret_cast<GrpcServer *>(GetWindowLongPtrW(hWnd, GWLP_USERDATA));
		if (pServer != nullptr)
		{
			// COM may reenter OnDisconnection during Invoke and release CConnect's
			// ownership. Keep the bridge alive until this window callback returns.
			auto server = pServer->shared_from_this();
			server->DispatchOnSta();
		}
		return 0;
	}
	else if (message == WM_NCDESTROY)
		SetWindowLongPtrW(hWnd, GWLP_USERDATA, 0);
	return DefWindowProcW(hWnd, message, wParam, lParam);
}

void GrpcServer::DispatchOnSta()
{
	ATLASSERT(m_ownerThreadId == GetCurrentThreadId());
	if (m_dispatching)
		return;
	m_dispatching = true;

	for (;;)
	{
		std::shared_ptr<PendingCall> call;
		{
			std::lock_guard<std::mutex> lock(m_mutex);
			if (m_stopping || m_pendingCalls.empty())
				break;
			call = m_pendingCalls.front();
			m_pendingCalls.pop_front();
			m_activeCall = call;
		}

		bool execute = false;
		{
			std::lock_guard<std::mutex> lock(call->mutex);
			if (!call->finished)
			{
				// Check the live context as well as the deadline immediately before
				// mutation, rather than trusting the worker's periodic polling.
				if (std::chrono::system_clock::now() >= call->deadline ||
					call->pContext == nullptr || call->pContext->IsCancelled())
				{
					call->status = CancellationStatus(call->deadline);
					call->finished = true;
				}
				else
					execute = true;
			}
		}

		if (execute)
		{
			grpc::Status status;
			netoffice::PresentationReply reply;
			try
			{
				if (call->operation == PendingCall::Operation::NewPresentation)
					status = CreatePresentation(call->text, &reply);
				else
					status = UpdateSlideTitle(call->slideIndex, call->text);
			}
			catch (const std::bad_alloc &)
			{
				status = ErrorStatus(grpc::StatusCode::RESOURCE_EXHAUSTED, E_OUTOFMEMORY,
					"PowerPoint STA request");
			}
			catch (const std::exception &error)
			{
				status = ErrorStatus(grpc::StatusCode::INTERNAL, E_FAIL,
					"PowerPoint STA request", error.what());
			}
			std::lock_guard<std::mutex> lock(call->mutex);
			if (!call->finished)
			{
				call->presentationReply = std::move(reply);
				call->status = std::move(status);
				call->finished = true;
			}
		}
		call->completed.notify_all();
		{
			std::lock_guard<std::mutex> lock(m_mutex);
			m_activeCall.reset();
		}
	}
	m_dispatching = false;
}

grpc::Status GrpcServer::Invoke(IDispatch *pObject, const wchar_t *pName, WORD flags,
	ATL::CComVariant *pArguments, UINT argumentCount, ATL::CComVariant *pResult,
	grpc::StatusCode failureCode)
{
	ATLASSERT(m_ownerThreadId == GetCurrentThreadId());
	if (IsStopping())
		return StoppedStatus();
	if (pObject == nullptr)
		return ErrorStatus(failureCode, E_POINTER, MemberName(pName));

	DISPID dispid;
	LPOLESTR name = const_cast<LPOLESTR>(pName);
	HRESULT hr = pObject->GetIDsOfNames(IID_NULL, &name, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr))
		return ErrorStatus(failureCode, hr, MemberName(pName));
	if (IsStopping())
		return StoppedStatus();

	DISPID propertyPut = DISPID_PROPERTYPUT;
	DISPPARAMS parameters = {};
	parameters.rgvarg = pArguments;
	parameters.cArgs = argumentCount;
	if ((flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0)
	{
		parameters.rgdispidNamedArgs = &propertyPut;
		parameters.cNamedArgs = 1;
	}
	EXCEPINFO exception = {};
	UINT argumentError = 0;
	hr = pObject->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, flags,
		&parameters, pResult, &exception, &argumentError);

	std::string details;
	if (FAILED(hr) && exception.pfnDeferredFillIn != nullptr)
	{
		HRESULT detailHr = exception.pfnDeferredFillIn(&exception);
		if (FAILED(detailHr))
			details = "Office could not supply deferred exception details";
	}
	ATL::CComBSTR source;
	ATL::CComBSTR description;
	ATL::CComBSTR helpFile;
	source.Attach(exception.bstrSource);
	description.Attach(exception.bstrDescription);
	helpFile.Attach(exception.bstrHelpFile);
	if (hr == DISP_E_EXCEPTION && FAILED(exception.scode))
		hr = exception.scode;
	if (FAILED(hr) && exception.bstrDescription != nullptr)
	{
		HRESULT detailHr = BstrToUtf8(exception.bstrDescription, details);
		if (FAILED(detailHr))
			details = "Office supplied an invalid Unicode exception description";
	}

	if (FAILED(hr))
		return ErrorStatus(failureCode, hr, MemberName(pName), details);
	return grpc::Status::OK;
}

grpc::Status GrpcServer::ReadObject(ATL::CComVariant &value, const wchar_t *pName,
	ATL::CComPtr<IDispatch> &result, grpc::StatusCode failureCode)
{
	HRESULT hr = E_NOINTERFACE;
	if (value.vt == VT_DISPATCH && value.pdispVal != nullptr)
	{
		result = value.pdispVal;
		return grpc::Status::OK;
	}
	if (value.vt == VT_UNKNOWN && value.punkVal != nullptr)
		hr = value.punkVal->QueryInterface(IID_IDispatch, reinterpret_cast<void **>(&result));
	if (FAILED(hr))
		return ErrorStatus(failureCode, hr, MemberName(pName), "No Automation object was returned");
	return grpc::Status::OK;
}

grpc::Status GrpcServer::GetObject(IDispatch *pObject, const wchar_t *pName,
	ATL::CComPtr<IDispatch> &result, grpc::StatusCode failureCode)
{
	ATL::CComVariant value;
	grpc::Status status = Invoke(pObject, pName, DISPATCH_PROPERTYGET, nullptr, 0, &value, failureCode);
	if (!status.ok())
		return status;
	return ReadObject(value, pName, result, failureCode);
}

grpc::Status GrpcServer::GetInteger(IDispatch *pObject, const wchar_t *pName, long &result)
{
	ATL::CComVariant value;
	grpc::Status status = Invoke(pObject, pName, DISPATCH_PROPERTYGET, nullptr, 0, &value);
	if (!status.ok())
		return status;
	HRESULT hr = value.ChangeType(VT_I4);
	if (FAILED(hr))
		return ErrorStatus(grpc::StatusCode::INTERNAL, hr, MemberName(pName));
	result = value.lVal;
	return grpc::Status::OK;
}

grpc::Status GrpcServer::PutSlideTitle(IDispatch *pSlide, ATL::CComVariant &text)
{
	ATL::CComPtr<IDispatch> shapes;
	grpc::Status status = GetObject(pSlide, L"Shapes", shapes);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> title;
	status = GetObject(shapes, L"Title", title, grpc::StatusCode::FAILED_PRECONDITION);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> textFrame;
	status = GetObject(title, L"TextFrame", textFrame);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> textRange;
	status = GetObject(textFrame, L"TextRange", textRange);
	if (!status.ok())
		return status;
	return Invoke(textRange, L"Text", DISPATCH_PROPERTYPUT, &text, 1, nullptr);
}

grpc::Status GrpcServer::CreatePresentation(const std::string &title,
	netoffice::PresentationReply *pReply)
{
	ATL::CComVariant titleText;
	HRESULT hr = Utf8ToVariant(title, titleText);
	if (FAILED(hr))
		return ErrorStatus(grpc::StatusCode::INVALID_ARGUMENT, hr, "Decoding presentation title as UTF-8");

	// All Automation references, including these local references held across
	// potentially reentrant Office calls, are acquired and released on the STA.
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> presentations;
	grpc::Status status = GetObject(application, L"Presentations", presentations);
	if (!status.ok())
		return status;

	ATL::CComVariant withWindow(-1L); // msoTrue
	ATL::CComVariant value;
	status = Invoke(presentations, L"Add", DISPATCH_METHOD, &withWindow, 1, &value);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> presentation;
	status = ReadObject(value, L"Presentations.Add", presentation, grpc::StatusCode::INTERNAL);
	if (!status.ok())
		return status;

	ATL::CComPtr<IDispatch> slides;
	status = GetObject(presentation, L"Slides", slides);
	if (!status.ok())
		return status;
	// IDispatch arguments are right-to-left: Layout, Index.
	ATL::CComVariant arguments[2] = { ATL::CComVariant(1L), ATL::CComVariant(1L) }; // ppLayoutTitle, first slide
	ATL::CComVariant slideValue;
	status = Invoke(slides, L"Add", DISPATCH_METHOD, arguments, 2, &slideValue);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slide;
	status = ReadObject(slideValue, L"Slides.Add", slide, grpc::StatusCode::INTERNAL);
	if (!status.ok())
		return status;
	status = PutSlideTitle(slide, titleText);
	if (!status.ok())
		return status;

	ATL::CComVariant name;
	status = Invoke(presentation, L"Name", DISPATCH_PROPERTYGET, nullptr, 0, &name);
	if (!status.ok())
		return status;
	if (name.vt != VT_BSTR)
		return ErrorStatus(grpc::StatusCode::INTERNAL, DISP_E_TYPEMISMATCH, "PowerPoint.Presentation.Name");
	std::string documentName;
	hr = BstrToUtf8(name.bstrVal, documentName);
	if (FAILED(hr))
		return ErrorStatus(grpc::StatusCode::INTERNAL, hr, "Encoding presentation name as UTF-8");
	long slideCount = 0;
	status = GetInteger(slides, L"Count", slideCount);
	if (!status.ok())
		return status;

	ATL::CComPtr<IDispatch> windows;
	status = GetObject(presentation, L"Windows", windows);
	if (!status.ok())
		return status;
	ATL::CComVariant windowIndex(1L);
	ATL::CComVariant windowValue;
	status = Invoke(windows, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &windowIndex, 1, &windowValue);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> window;
	status = ReadObject(windowValue, L"Windows.Item", window, grpc::StatusCode::INTERNAL);
	if (!status.ok())
		return status;
	status = Invoke(window, L"Activate", DISPATCH_METHOD, nullptr, 0, nullptr);
	if (!status.ok())
		return status;

	pReply->set_name(std::move(documentName));
	pReply->set_slide_count(slideCount);
	return grpc::Status::OK;
}

grpc::Status GrpcServer::UpdateSlideTitle(int slideIndex, const std::string &text)
{
	ATL::CComVariant titleText;
	HRESULT hr = Utf8ToVariant(text, titleText);
	if (FAILED(hr))
		return ErrorStatus(grpc::StatusCode::INVALID_ARGUMENT, hr, "Decoding slide title as UTF-8");

	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> presentations;
	grpc::Status status = GetObject(application, L"Presentations", presentations);
	if (!status.ok())
		return status;
	long presentationCount = 0;
	status = GetInteger(presentations, L"Count", presentationCount);
	if (!status.ok())
		return status;
	if (presentationCount == 0)
		return ErrorStatus(grpc::StatusCode::FAILED_PRECONDITION,
			HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "SetSlideTitle", "There is no active presentation");

	ATL::CComPtr<IDispatch> presentation;
	status = GetObject(application, L"ActivePresentation", presentation, grpc::StatusCode::FAILED_PRECONDITION);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slides;
	status = GetObject(presentation, L"Slides", slides);
	if (!status.ok())
		return status;
	long slideCount = 0;
	status = GetInteger(slides, L"Count", slideCount);
	if (!status.ok())
		return status;
	if (slideIndex > slideCount)
		return ErrorStatus(grpc::StatusCode::INVALID_ARGUMENT, E_INVALIDARG,
			"SetSlideTitle", "slide_index exceeds the active presentation's slide count");

	ATL::CComVariant index(static_cast<long>(slideIndex));
	ATL::CComVariant value;
	status = Invoke(slides, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &index, 1, &value);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slide;
	status = ReadObject(value, L"Slides.Item", slide, grpc::StatusCode::INTERNAL);
	if (!status.ok())
		return status;
	return PutSlideTitle(slide, titleText);
}
