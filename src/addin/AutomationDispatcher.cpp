#include "pch.h"
#include "AutomationDispatcher.h"

#include <algorithm>
#include <climits>
#include <condition_variable>
#include <cstdio>
#include <cwchar>
#include <exception>
#include <new>
#include <utility>

namespace
{
	constexpr UINT AutomationDispatchMessage = WM_APP + 1;
	constexpr auto CancellationPollInterval = std::chrono::milliseconds(20);
	constexpr size_t MaximumQueuedCalls = 128;
	constexpr uint64_t MaximumSafeId = 9007199254740991ULL;

	bool IntegerInRange(const nlohmann::json &value, uint64_t minimum, uint64_t maximum)
	{
		if (value.is_number_unsigned())
		{
			auto number = value.get<uint64_t>();
			return number >= minimum && number <= maximum;
		}
		if (!value.is_number_integer())
			return false;
		auto number = value.get<int64_t>();
		return number >= 0 && static_cast<uint64_t>(number) >= minimum &&
			static_cast<uint64_t>(number) <= maximum;
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

	std::string MemberName(const wchar_t *name)
	{
		// These Automation member names are ASCII; user text uses checked UTF-8.
		std::string result = "PowerPoint.";
		while (*name != L'\0')
			result.push_back(static_cast<char>(*name++));
		return result;
	}
}

AutomationDispatcher::Status AutomationDispatcher::ErrorStatus(int code, HRESULT hr,
	const std::string &operation, const std::string &details)
{
	char hresult[16];
	sprintf_s(hresult, "0x%08lX", static_cast<unsigned long>(hr));
	Status status;
	status.code = code;
	status.message = operation + " failed (HRESULT " + hresult + ")";
	status.data = { { "hresult", hresult }, { "operation", operation } };
	if (!details.empty())
	{
		status.message += ": " + details;
		status.data["details"] = details;
	}
	return status;
}

AutomationDispatcher::Status AutomationDispatcher::StoppedStatus()
{
	return ErrorStatus(-32003, HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED),
		"PowerPoint addin connection");
}

AutomationDispatcher::Status AutomationDispatcher::CancellationStatus(
	std::chrono::steady_clock::time_point deadline)
{
	if (std::chrono::steady_clock::now() >= deadline)
		return ErrorStatus(-32002, HRESULT_FROM_WIN32(ERROR_TIMEOUT), "PowerPoint request deadline");
	return ErrorStatus(-32003, HRESULT_FROM_WIN32(ERROR_CANCELLED), "PowerPoint request connection");
}

nlohmann::json AutomationDispatcher::ErrorReply(const nlohmann::json &id, const Status &status)
{
	nlohmann::json error = { { "code", status.code }, { "message", status.message } };
	if (!status.data.is_null())
		error["data"] = status.data;
	return { { "id", id }, { "error", std::move(error) } };
}

struct AutomationDispatcher::PendingCall
{
	enum class Operation { GetStatus, NewPresentation, SetSlideTitle };
	explicit PendingCall(Operation operation) : operation(operation) {}

	Operation operation;
	std::string text;
	int slideIndex = 0;
	nlohmann::json result;
	std::shared_ptr<std::atomic_bool> cancelled;
	std::chrono::steady_clock::time_point deadline;
	std::mutex mutex;
	std::condition_variable completed;
	bool finished = false;
	// Preallocated: reentrant disconnect only marks completion and wakes workers.
	Status status = StoppedStatus();

	bool IsCancelled() const
	{
		return (cancelled && cancelled->load(std::memory_order_relaxed)) ||
			std::chrono::steady_clock::now() >= deadline;
	}

	void FinishForShutdown()
	{
		{
			std::lock_guard<std::mutex> lock(mutex);
			finished = true;
		}
		completed.notify_all();
	}
};

AutomationDispatcher::~AutomationDispatcher()
{
	// The transport stops us on the STA before releasing any worker ownership.
	ATLASSERT(m_hWnd == nullptr && m_pApplication.p == nullptr);
}

HRESULT AutomationDispatcher::Start(IDispatch *app)
{
	if (app == nullptr)
		return E_POINTER;
	APTTYPE apartment;
	APTTYPEQUALIFIER qualifier;
	HRESULT hr = CoGetApartmentType(&apartment, &qualifier);
	if (FAILED(hr))
		return hr;
	if (apartment != APTTYPE_STA && apartment != APTTYPE_MAINSTA)
		return RPC_E_WRONG_THREAD;
	if (weak_from_this().expired())
		return E_UNEXPECTED;
	if (m_hWnd != nullptr || m_dispatching)
		return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);

	m_ownerThreadId = GetCurrentThreadId();
	m_pApplication = app;
	m_hInstance = ATL::_AtlBaseModule.GetModuleInstance();
	wchar_t className[80];
	swprintf_s(className, L"NetOffice.PowerPoint.Automation.%p", static_cast<void *>(this));
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
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_stopping = false;
	}
	return S_OK;
}

void AutomationDispatcher::Stop()
{
	ATLASSERT(m_ownerThreadId == 0 || m_ownerThreadId == GetCurrentThreadId());
	HWND window;
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_stopping = true;
		for (const auto &call : m_pendingCalls)
			call->FinishForShutdown();
		m_pendingCalls.clear();
		if (m_activeCall)
			m_activeCall->FinishForShutdown();
		window = m_hWnd;
		m_hWnd = nullptr;
	}
	if (window != nullptr)
	{
		SetWindowLongPtrW(window, GWLP_USERDATA, 0);
		DestroyWindow(window);
	}
	if (m_windowClass != 0)
	{
		UnregisterClassW(MAKEINTATOM(m_windowClass), m_hInstance);
		m_windowClass = 0;
	}
	m_pApplication.Release();
}

bool AutomationDispatcher::IsStopping()
{
	std::lock_guard<std::mutex> lock(m_mutex);
	return m_stopping;
}

nlohmann::json AutomationDispatcher::HandleRequest(const nlohmann::json &request,
	const std::shared_ptr<std::atomic_bool> &cancelled)
{
	const auto received = std::chrono::steady_clock::now();
	nlohmann::json id = nullptr;
	try
	{
		if (request.is_object() && request.contains("id") &&
			IntegerInRange(request["id"], 0, MaximumSafeId))
			id = request["id"];
		auto invalid = [&id](int code, const char *message)
		{
			return nlohmann::json{ { "id", id }, { "error", { { "code", code }, { "message", message } } } };
		};
		if (!request.is_object() || id.is_null() || !request.contains("method") ||
			!request["method"].is_string())
			return invalid(-32600, "Request requires a nonnegative JavaScript-safe integer id and a string method");
		if (request.contains("params") && !request["params"].is_object())
			return invalid(-32600, "params must be an object when present");
		int timeoutMs = 10000;
		if (request.contains("timeoutMs"))
		{
			if (!IntegerInRange(request["timeoutMs"], 1, INT_MAX))
				return invalid(-32600, "timeoutMs must be a positive integer no greater than 2147483647");
			timeoutMs = request["timeoutMs"].get<int>();
		}
		const auto &method = request["method"].get_ref<const std::string &>();
		PendingCall::Operation operation;
		if (method == "PowerPoint.getStatus")
			operation = PendingCall::Operation::GetStatus;
		else if (method == "PowerPoint.newPresentation")
			operation = PendingCall::Operation::NewPresentation;
		else if (method == "PowerPoint.setSlideTitle")
			operation = PendingCall::Operation::SetSlideTitle;
		else
			return invalid(-32601, "Unknown PowerPoint method");

		static const nlohmann::json emptyParams = nlohmann::json::object();
		const auto &params = request.contains("params") ? request["params"] : emptyParams;
		if (operation == PendingCall::Operation::NewPresentation && params.contains("title") &&
			!params["title"].is_string())
			return invalid(-32602, "title must be a string");
		if (operation == PendingCall::Operation::SetSlideTitle &&
			(!params.contains("slideIndex") || !IntegerInRange(params["slideIndex"], 1, INT_MAX) ||
			 !params.contains("text") || !params["text"].is_string()))
			return invalid(-32602, "setSlideTitle requires a positive int32 slideIndex and string text");

		auto call = std::make_shared<PendingCall>(operation);
		call->cancelled = cancelled;
		call->deadline = received + std::chrono::milliseconds(timeoutMs);
		if (operation == PendingCall::Operation::NewPresentation && params.contains("title"))
			call->text = params["title"].get<std::string>();
		else if (operation == PendingCall::Operation::SetSlideTitle)
		{
			call->slideIndex = params["slideIndex"].get<int>();
			call->text = params["text"].get<std::string>();
		}
		Status status = Dispatch(call);
		if (!status.ok())
			return ErrorReply(id, status);
		return { { "id", id }, { "result", std::move(call->result) } };
	}
	catch (const std::bad_alloc &)
	{
		return ErrorReply(id, ErrorStatus(-32000, E_OUTOFMEMORY, "PowerPoint request"));
	}
	catch (const std::exception &error)
	{
		return ErrorReply(id, ErrorStatus(-32000, E_FAIL, "PowerPoint request", error.what()));
	}
}

AutomationDispatcher::Status AutomationDispatcher::Dispatch(const std::shared_ptr<PendingCall> &call)
{
	// Exceptional worker exits must also disarm any request still waiting on the STA.
	struct CompletionGuard
	{
		PendingCall &call;
		~CompletionGuard()
		{
			std::lock_guard<std::mutex> lock(call.mutex);
			call.finished = true;
		}
	} guard{ *call };
	if (call->IsCancelled())
		return CancellationStatus(call->deadline);
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (m_stopping)
			return StoppedStatus();
		if (m_pendingCalls.size() >= MaximumQueuedCalls)
			return ErrorStatus(-32000, HRESULT_FROM_WIN32(ERROR_BUSY), "PowerPoint request queue");
		m_pendingCalls.push_back(call);
		if (!PostMessageW(m_hWnd, AutomationDispatchMessage, 0, 0))
		{
			HRESULT hr = HRESULT_FROM_WIN32(GetLastError());
			m_pendingCalls.pop_back();
			return ErrorStatus(-32000, hr, "Posting PowerPoint STA request");
		}
	}

	std::unique_lock<std::mutex> lock(call->mutex);
	bool cancelled = false;
	while (!call->finished)
	{
		if (call->IsCancelled())
		{
			call->status = CancellationStatus(call->deadline);
			call->finished = true;
			cancelled = true;
			break;
		}
		call->completed.wait_until(lock,
			(std::min)(call->deadline, std::chrono::steady_clock::now() + CancellationPollInterval));
	}
	Status status = std::move(call->status);
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

LRESULT CALLBACK AutomationDispatcher::WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam)
{
	if (message == WM_NCCREATE)
	{
		auto create = reinterpret_cast<CREATESTRUCTW *>(lParam);
		SetWindowLongPtrW(window, GWLP_USERDATA, reinterpret_cast<LONG_PTR>(create->lpCreateParams));
	}
	else if (message == AutomationDispatchMessage)
	{
		auto dispatcher = reinterpret_cast<AutomationDispatcher *>(GetWindowLongPtrW(window, GWLP_USERDATA));
		if (dispatcher != nullptr)
		{
			// Invoke can reenter disconnect and release the transport's ownership.
			auto lifetime = dispatcher->shared_from_this();
			lifetime->DispatchOnSta();
		}
		return 0;
	}
	else if (message == WM_NCDESTROY)
		SetWindowLongPtrW(window, GWLP_USERDATA, 0);
	return DefWindowProcW(window, message, wParam, lParam);
}

void AutomationDispatcher::DispatchOnSta()
{
	ATLASSERT(m_ownerThreadId == GetCurrentThreadId());
	if (m_dispatching)
		return;
	m_dispatching = true;
	struct DispatchGuard
	{
		bool &dispatching;
		~DispatchGuard() { dispatching = false; }
	} guard{ m_dispatching };

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
				// Recheck immediately before starting, independently of worker polling.
				if (call->IsCancelled())
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
			Status status;
			nlohmann::json result;
			try
			{
				result = nlohmann::json::object();
				if (call->operation == PendingCall::Operation::GetStatus)
					result["processId"] = GetCurrentProcessId();
				else if (call->operation == PendingCall::Operation::NewPresentation)
					status = CreatePresentation(call->text, result);
				else
					status = UpdateSlideTitle(call->slideIndex, call->text);
				// A running COM call may complete after cancellation, but its late
				// completion must not become a successful deadline/connection reply.
				if (call->IsCancelled())
					status = CancellationStatus(call->deadline);
			}
			catch (const std::bad_alloc &)
			{
				status = ErrorStatus(-32000, E_OUTOFMEMORY, "PowerPoint STA request");
			}
			catch (const std::exception &error)
			{
				status = ErrorStatus(-32000, E_FAIL, "PowerPoint STA request", error.what());
			}
			std::lock_guard<std::mutex> lock(call->mutex);
			if (!call->finished)
			{
				call->result = std::move(result);
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
}

AutomationDispatcher::Status AutomationDispatcher::Invoke(IDispatch *object, const wchar_t *name,
	WORD flags, ATL::CComVariant *arguments, UINT argumentCount, ATL::CComVariant *result, int failureCode)
{
	ATLASSERT(m_ownerThreadId == GetCurrentThreadId());
	if (IsStopping())
		return StoppedStatus();
	if (object == nullptr)
		return ErrorStatus(failureCode, E_POINTER, MemberName(name));
	DISPID dispid;
	LPOLESTR member = const_cast<LPOLESTR>(name);
	HRESULT hr = object->GetIDsOfNames(IID_NULL, &member, 1, LOCALE_USER_DEFAULT, &dispid);
	if (FAILED(hr))
		return ErrorStatus(failureCode, hr, MemberName(name));
	if (IsStopping())
		return StoppedStatus();
	DISPID propertyPut = DISPID_PROPERTYPUT;
	DISPPARAMS parameters = {};
	parameters.rgvarg = arguments;
	parameters.cArgs = argumentCount;
	if ((flags & (DISPATCH_PROPERTYPUT | DISPATCH_PROPERTYPUTREF)) != 0)
	{
		parameters.rgdispidNamedArgs = &propertyPut;
		parameters.cNamedArgs = 1;
	}
	EXCEPINFO exception = {};
	UINT argumentError = 0;
	hr = object->Invoke(dispid, IID_NULL, LOCALE_USER_DEFAULT, flags,
		&parameters, result, &exception, &argumentError);
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
	if (IsStopping())
		return StoppedStatus();
	if (FAILED(hr))
		return ErrorStatus(failureCode, hr, MemberName(name), details);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadObject(ATL::CComVariant &value,
	const wchar_t *name, ATL::CComPtr<IDispatch> &result, int failureCode)
{
	HRESULT hr = E_NOINTERFACE;
	if (value.vt == VT_DISPATCH && value.pdispVal != nullptr)
	{
		result = value.pdispVal;
		return {};
	}
	if (value.vt == VT_UNKNOWN && value.punkVal != nullptr)
		hr = value.punkVal->QueryInterface(IID_IDispatch, reinterpret_cast<void **>(&result));
	if (FAILED(hr))
		return ErrorStatus(failureCode, hr, MemberName(name), "No Automation object was returned");
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetObject(IDispatch *object, const wchar_t *name,
	ATL::CComPtr<IDispatch> &result, int failureCode)
{
	ATL::CComVariant value;
	Status status = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &value, failureCode);
	if (!status.ok())
		return status;
	return ReadObject(value, name, result, failureCode);
}

AutomationDispatcher::Status AutomationDispatcher::GetInteger(IDispatch *object,
	const wchar_t *name, long &result)
{
	ATL::CComVariant value;
	Status status = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &value);
	if (!status.ok())
		return status;
	HRESULT hr = value.ChangeType(VT_I4);
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, MemberName(name));
	result = value.lVal;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::PutSlideTitle(IDispatch *slide, ATL::CComVariant &text)
{
	ATL::CComPtr<IDispatch> shapes;
	Status status = GetObject(slide, L"Shapes", shapes);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> title;
	status = GetObject(shapes, L"Title", title, -32001);
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

AutomationDispatcher::Status AutomationDispatcher::CreatePresentation(const std::string &title,
	nlohmann::json &result)
{
	ATL::CComVariant titleText;
	HRESULT hr = Utf8ToVariant(title, titleText);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Decoding presentation title as UTF-8");
	// Local Automation references survive reentrant disconnect and release on STA.
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> presentations;
	Status status = GetObject(application, L"Presentations", presentations);
	if (!status.ok())
		return status;
	ATL::CComVariant withWindow(-1L); // msoTrue
	ATL::CComVariant value;
	status = Invoke(presentations, L"Add", DISPATCH_METHOD, &withWindow, 1, &value);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> presentation;
	status = ReadObject(value, L"Presentations.Add", presentation, -32000);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slides;
	status = GetObject(presentation, L"Slides", slides);
	if (!status.ok())
		return status;
	// IDispatch arguments are right-to-left: ppLayoutTitle, one-based index.
	ATL::CComVariant arguments[2] = { ATL::CComVariant(1L), ATL::CComVariant(1L) };
	ATL::CComVariant slideValue;
	status = Invoke(slides, L"Add", DISPATCH_METHOD, arguments, 2, &slideValue);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slide;
	status = ReadObject(slideValue, L"Slides.Add", slide, -32000);
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
		return ErrorStatus(-32000, DISP_E_TYPEMISMATCH, "PowerPoint.Presentation.Name");
	std::string documentName;
	hr = BstrToUtf8(name.bstrVal, documentName);
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, "Encoding presentation name as UTF-8");
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
	status = ReadObject(windowValue, L"Windows.Item", window, -32000);
	if (!status.ok())
		return status;
	status = Invoke(window, L"Activate", DISPATCH_METHOD, nullptr, 0, nullptr);
	if (!status.ok())
		return status;
	result = { { "name", std::move(documentName) }, { "slideCount", slideCount } };
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::UpdateSlideTitle(int slideIndex, const std::string &text)
{
	ATL::CComVariant titleText;
	HRESULT hr = Utf8ToVariant(text, titleText);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Decoding slide title as UTF-8");
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> presentations;
	Status status = GetObject(application, L"Presentations", presentations);
	if (!status.ok())
		return status;
	long presentationCount = 0;
	status = GetInteger(presentations, L"Count", presentationCount);
	if (!status.ok())
		return status;
	if (presentationCount == 0)
		return ErrorStatus(-32001, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "SetSlideTitle",
			"There is no active presentation");
	ATL::CComPtr<IDispatch> presentation;
	status = GetObject(application, L"ActivePresentation", presentation, -32001);
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
		return ErrorStatus(-32602, E_INVALIDARG, "SetSlideTitle",
			"slideIndex exceeds the active presentation's slide count");
	ATL::CComVariant index(static_cast<long>(slideIndex));
	ATL::CComVariant value;
	status = Invoke(slides, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &index, 1, &value);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slide;
	status = ReadObject(value, L"Slides.Item", slide, -32000);
	if (!status.ok())
		return status;
	return PutSlideTitle(slide, titleText);
}
