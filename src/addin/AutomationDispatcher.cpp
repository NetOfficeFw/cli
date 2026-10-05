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
#include <vector>
#include <winver.h>

namespace
{
	constexpr UINT AutomationDispatchMessage = WM_APP + 1;
	constexpr auto CancellationPollInterval = std::chrono::milliseconds(20);
	constexpr size_t MaximumQueuedCalls = 128;
	constexpr uint64_t MaximumSafeId = 9007199254740991ULL;
	constexpr size_t MaximumHttpArgumentBytes = 1024 * 1024;

	bool IsTargetId(const std::string &id)
	{
		if (id.size() != 36)
			return false;
		for (size_t index = 0; index < id.size(); ++index)
		{
			if (index == 8 || index == 13 || index == 18 || index == 23)
			{
				if (id[index] != '-')
					return false;
			}
			else if (!((id[index] >= '0' && id[index] <= '9') ||
				(id[index] >= 'a' && id[index] <= 'f') ||
				(id[index] >= 'A' && id[index] <= 'F')))
				return false;
		}
		return true;
	}

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
	enum class Operation { GetStatus, NewPresentation, SetSlideTitle, Http };
	explicit PendingCall(Operation operation) : operation(operation) {}

	Operation operation;
	std::string text;
	int slideIndex = 0;
	long slideId = 0;
	HttpCommand httpCommand = HttpCommand::Version;
	bool force = false;
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
	// Empty the member before releasing COM; Release may itself reenter disconnect.
	std::vector<DocumentTarget> targets;
	m_targets.swap(targets);
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
		HttpCommand httpCommand = HttpCommand::Version;
		if (method == "PowerPoint.getStatus")
			operation = PendingCall::Operation::GetStatus;
		else if (method == "PowerPoint.newPresentation")
			operation = PendingCall::Operation::NewPresentation;
		else if (method == "PowerPoint.setSlideTitle")
			operation = PendingCall::Operation::SetSlideTitle;
		else
		{
			operation = PendingCall::Operation::Http;
			if (method == "PowerPoint.getPresentationState")
				httpCommand = HttpCommand::Presentation;
			else if (method == "PowerPoint.getSlides")
				httpCommand = HttpCommand::Slides;
			else if (method == "PowerPoint.getSlideState")
				httpCommand = HttpCommand::Slide;
			else if (method == "PowerPoint.getViewState")
				httpCommand = HttpCommand::View;
			else if (method == "PowerPoint.getSlideShowState")
				httpCommand = HttpCommand::SlideShow;
			else
				return invalid(-32601, "Unknown PowerPoint method");
		}

		static const nlohmann::json emptyParams = nlohmann::json::object();
		const auto &params = request.contains("params") ? request["params"] : emptyParams;
		if (operation == PendingCall::Operation::NewPresentation && params.contains("title") &&
			!params["title"].is_string())
			return invalid(-32602, "title must be a string");
		if (operation == PendingCall::Operation::SetSlideTitle &&
			(!params.contains("slideIndex") || !IntegerInRange(params["slideIndex"], 1, INT_MAX) ||
			 !params.contains("text") || !params["text"].is_string()))
			return invalid(-32602, "setSlideTitle requires a positive int32 slideIndex and string text");
		if (operation == PendingCall::Operation::Http &&
			(!params.contains("targetId") || !params["targetId"].is_string() ||
			 !IsTargetId(params["targetId"].get<std::string>())))
			return invalid(-32602, "State reads require a canonical GUID targetId");
		if (httpCommand == HttpCommand::Slide &&
			(!params.contains("slideId") || !IntegerInRange(params["slideId"], 1, INT_MAX)))
			return invalid(-32602, "getSlideState requires a positive int32 slideId");

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
		else if (operation == PendingCall::Operation::Http)
		{
			call->httpCommand = httpCommand;
			call->text = params["targetId"].get<std::string>();
			if (httpCommand == HttpCommand::Slide)
				call->slideId = params["slideId"].get<long>();
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

nlohmann::json AutomationDispatcher::HandleHttpRequest(HttpCommand command,
	const std::string &argument, long slideId, bool force,
	const std::shared_ptr<std::atomic_bool> &cancelled)
{
	const auto received = std::chrono::steady_clock::now();
	auto reply = [](const Status &status)
	{
		auto response = ErrorReply(nullptr, status);
		response.erase("id");
		return response;
	};
	try
	{
		if (argument.size() > MaximumHttpArgumentBytes || argument.find('\0') != std::string::npos)
			return reply(ErrorStatus(-32602, E_INVALIDARG, "HTTP document argument",
				"Argument exceeds 1 MiB or contains a null character"));
		if ((command == HttpCommand::Activate || command == HttpCommand::Close ||
			command == HttpCommand::Presentation || command == HttpCommand::Slides ||
			command == HttpCommand::Slide || command == HttpCommand::View ||
			command == HttpCommand::SlideShow) && !IsTargetId(argument))
			return reply(ErrorStatus(-32602, E_INVALIDARG, "HTTP document target",
				"Target must be a canonical GUID"));
		if (command == HttpCommand::Slide && slideId < 1)
			return reply(ErrorStatus(-32602, E_INVALIDARG, "HTTP slide id",
				"Slide ID must be a positive int32"));
		if (command != HttpCommand::Version && command != HttpCommand::List &&
			command != HttpCommand::New && command != HttpCommand::Activate && command != HttpCommand::Close &&
			command != HttpCommand::Presentation && command != HttpCommand::Slides &&
			command != HttpCommand::Slide && command != HttpCommand::View &&
			command != HttpCommand::SlideShow)
			return reply(ErrorStatus(-32602, E_INVALIDARG, "HTTP command"));
		auto call = std::make_shared<PendingCall>(PendingCall::Operation::Http);
		call->httpCommand = command;
		call->text = argument;
		call->slideId = slideId;
		call->force = force;
		call->cancelled = cancelled;
		call->deadline = received + std::chrono::milliseconds(10000);
		Status status = Dispatch(call);
		if (!status.ok())
			return reply(status);
		return { { "result", std::move(call->result) } };
	}
	catch (const std::bad_alloc &)
	{
		return reply(ErrorStatus(-32000, E_OUTOFMEMORY, "HTTP document request"));
	}
	catch (const std::exception &error)
	{
		return reply(ErrorStatus(-32000, E_FAIL, "HTTP document request", error.what()));
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
				else if (call->operation == PendingCall::Operation::Http)
					status = DispatchHttp(call->httpCommand, call->text, call->slideId,
						call->force, result);
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

AutomationDispatcher::Status AutomationDispatcher::GetString(IDispatch *object,
	const wchar_t *name, std::string &result)
{
	ATL::CComVariant value;
	Status status = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &value);
	if (!status.ok())
		return status;
	if (value.vt != VT_BSTR)
		return ErrorStatus(-32000, DISP_E_TYPEMISMATCH, MemberName(name));
	HRESULT hr = BstrToUtf8(value.bstrVal, result);
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, MemberName(name), "Invalid Unicode string");
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetExecutableBuild(IDispatch *application,
	std::string &result)
{
	ATL::CComVariant path;
	Status status = Invoke(application, L"Path", DISPATCH_PROPERTYGET, nullptr, 0, &path);
	if (!status.ok())
		return status;
	if (path.vt != VT_BSTR)
		return ErrorStatus(-32000, DISP_E_TYPEMISMATCH, "PowerPoint.Path");
	std::wstring executable(path.bstrVal, SysStringLen(path.bstrVal));
	if (executable.empty())
		return ErrorStatus(-32000, E_UNEXPECTED, "PowerPoint.Path", "Office returned an empty program path");
	if (executable.back() != L'\\' && executable.back() != L'/')
		executable.push_back(L'\\');
	executable += L"POWERPNT.EXE";
	DWORD ignored = 0;
	DWORD bytes = GetFileVersionInfoSizeW(executable.c_str(), &ignored);
	if (bytes == 0)
		return ErrorStatus(-32000, HRESULT_FROM_WIN32(GetLastError()), "Reading POWERPNT.EXE version");
	std::vector<BYTE> buffer(bytes);
	if (!GetFileVersionInfoW(executable.c_str(), 0, bytes, buffer.data()))
		return ErrorStatus(-32000, HRESULT_FROM_WIN32(GetLastError()), "Reading POWERPNT.EXE version");
	VS_FIXEDFILEINFO *version = nullptr;
	UINT size = 0;
	if (!VerQueryValueW(buffer.data(), L"\\", reinterpret_cast<void **>(&version), &size) ||
		version == nullptr || size < sizeof(VS_FIXEDFILEINFO) || version->dwSignature != 0xFEEF04BD)
		return ErrorStatus(-32000, E_UNEXPECTED, "Reading POWERPNT.EXE version",
			"Executable has no valid fixed file version");
	result = std::to_string(HIWORD(version->dwFileVersionMS)) + "." +
		std::to_string(LOWORD(version->dwFileVersionMS)) + "." +
		std::to_string(HIWORD(version->dwFileVersionLS)) + "." +
		std::to_string(LOWORD(version->dwFileVersionLS));
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ApplicationMetadata(nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> application = m_pApplication;
	std::string name;
	Status status = GetString(application, L"Name", name);
	if (!status.ok())
		return status;
	std::string version;
	status = GetString(application, L"Version", version);
	if (!status.ok())
		return status;
	std::string build;
	status = GetString(application, L"Build", build);
	if (!status.ok())
	{
		// Older Office type libraries may omit Build; never substitute the addin's version.
		if (status.data.value("hresult", std::string()) != "0x80020006" &&
			status.data.value("hresult", std::string()) != "0x80020003")
			return status;
		status = GetExecutableBuild(application, build);
		if (!status.ok())
			return status;
	}
#if defined(_WIN64)
	const char *architecture = "x64";
#else
	const char *architecture = "x86";
#endif
	result = { { "Application", name + "/" + version + " (build " + build + "; " + architecture + ")" },
		{ "V8-Version", nullptr } };
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::DescribeDocument(IDispatch *document,
	const std::string &id, nlohmann::json &result)
{
	std::string title;
	Status status = GetString(document, L"Name", title);
	if (!status.ok())
		return status;
	std::string path;
	status = GetString(document, L"Path", path);
	if (!status.ok())
		return status;
	std::string url;
	// FullName is only a title for never-saved documents, not a resolvable location.
	if (!path.empty())
	{
		status = GetString(document, L"FullName", url);
		if (!status.ok())
			return status;
	}
	result = { { "id", id }, { "type", "document" }, { "title", std::move(title) },
		{ "url", std::move(url) } };
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::RefreshTargets(nlohmann::json &result)
{
	ATLASSERT(m_ownerThreadId == GetCurrentThreadId());
	// A COM call can reenter Stop; never iterate the member cache across such a call.
	const auto previous = m_targets;
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> presentations;
	Status status = GetObject(application, L"Presentations", presentations);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(presentations, L"Count", count);
	if (!status.ok())
		return status;
	if (count < 0)
		return ErrorStatus(-32000, E_UNEXPECTED, "PowerPoint.Presentations.Count");
	std::vector<DocumentTarget> snapshot;
	snapshot.reserve(static_cast<size_t>(count));
	auto descriptors = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComVariant argument(index);
		ATL::CComVariant value;
		status = Invoke(presentations, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &argument, 1, &value);
		if (!status.ok())
			return status;
		DocumentTarget target;
		status = ReadObject(value, L"Presentations.Item", target.document, -32000);
		if (!status.ok())
			return status;
		HRESULT hr = target.document->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&target.identity));
		if (IsStopping())
			return StoppedStatus();
		if (FAILED(hr))
			return ErrorStatus(-32000, hr, "Presentation canonical identity");
		auto existing = std::find_if(previous.begin(), previous.end(), [&target](const DocumentTarget &entry)
		{
			return entry.identity.p == target.identity.p;
		});
		std::string id;
		if (existing != previous.end())
			id = existing->descriptor["id"].get<std::string>();
		else
		{
			GUID guid;
			hr = CoCreateGuid(&guid);
			if (FAILED(hr))
				return ErrorStatus(-32000, hr, "Creating document target id");
			wchar_t text[39];
			if (StringFromGUID2(guid, text, _countof(text)) == 0)
				return ErrorStatus(-32000, E_UNEXPECTED, "Formatting document target id");
			id.reserve(36);
			for (size_t character = 1; character <= 36; ++character)
			{
				char digit = static_cast<char>(text[character]);
				id.push_back(digit >= 'A' && digit <= 'F' ? digit + ('a' - 'A') : digit);
			}
		}
		status = DescribeDocument(target.document, id, target.descriptor);
		if (!status.ok())
			return status;
		descriptors.push_back(target.descriptor);
		snapshot.push_back(std::move(target));
	}
	if (IsStopping())
		return StoppedStatus();
	m_targets.swap(snapshot);
	result = std::move(descriptors);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::OpenDocument(const std::string &argument,
	nlohmann::json &result)
{
	HRESULT hr = S_OK;
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> presentations;
	Status status = GetObject(application, L"Presentations", presentations);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	if (argument.empty())
	{
		ATL::CComVariant withWindow(-1L); // msoTrue; Office's normal blank presentation, no slides added.
		status = Invoke(presentations, L"Add", DISPATCH_METHOD, &withWindow, 1, &value);
	}
	else
	{
		// IDispatch reverses FileName, ReadOnly, Untitled, WithWindow.
		ATL::CComVariant arguments[4] = { ATL::CComVariant(-1L), ATL::CComVariant(0L),
			ATL::CComVariant(0L), ATL::CComVariant() };
		hr = Utf8ToVariant(argument, arguments[3]);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Decoding document URL as UTF-8");
		status = Invoke(presentations, L"Open", DISPATCH_METHOD, arguments, 4, &value);
	}
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> document;
	status = ReadObject(value, argument.empty() ? L"Presentations.Add" : L"Presentations.Open", document, -32000);
	if (!status.ok())
		return status;
	ATL::CComPtr<IUnknown> identity;
	hr = document->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&identity));
	if (IsStopping())
		return StoppedStatus();
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, "Presentation canonical identity");
	nlohmann::json list;
	status = RefreshTargets(list);
	if (!status.ok())
		return status;
	auto target = std::find_if(m_targets.begin(), m_targets.end(), [&identity](const DocumentTarget &entry)
	{
		return entry.identity.p == identity.p;
	});
	if (target == m_targets.end())
		return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "New document target",
			"The opened document is no longer available");
	result = target->descriptor;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ActivateDocument(IDispatch *document)
{
	ATL::CComPtr<IDispatch> windows;
	Status status = GetObject(document, L"Windows", windows);
	if (!status.ok())
		return status;
	ATL::CComVariant index(1L);
	ATL::CComVariant value;
	status = Invoke(windows, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET, &index, 1, &value);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> window;
	status = ReadObject(value, L"Windows.Item", window, -32000);
	if (!status.ok())
		return status;
	// Bring Office to the foreground before selecting the document. Activating
	// the previous frame afterward would switch back to its presentation.
	HWND nativeWindow = GetActiveWindow();
	if (!IsWindow(nativeWindow))
		return ErrorStatus(-32000, HRESULT_FROM_WIN32(ERROR_INVALID_WINDOW_HANDLE), "Document window activation");
	HWND root = GetAncestor(nativeWindow, GA_ROOT);
	if (IsIconic(root))
		ShowWindow(root, SW_RESTORE);
	if (GetForegroundWindow() != root)
		SetForegroundWindow(root);
	if (GetForegroundWindow() != root)
	{
		// The Office STA can temporarily share foreground input to honor explicit activation.
		DWORD foregroundThread = GetWindowThreadProcessId(GetForegroundWindow(), nullptr);
		bool attached = foregroundThread != 0 && foregroundThread != GetCurrentThreadId() &&
			AttachThreadInput(GetCurrentThreadId(), foregroundThread, TRUE) != FALSE;
		SetForegroundWindow(root);
		if (attached)
			AttachThreadInput(GetCurrentThreadId(), foregroundThread, FALSE);
	}
	if (IsStopping())
		return StoppedStatus();
	if (GetForegroundWindow() != root)
		return ErrorStatus(-32000, HRESULT_FROM_WIN32(ERROR_ACCESS_DENIED), "Document window activation",
			"Windows refused to bring the document window to the foreground");
	status = Invoke(window, L"Activate", DISPATCH_METHOD, nullptr, 0, nullptr);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> activeDocument;
	status = GetObject(application, L"ActivePresentation", activeDocument);
	if (!status.ok())
		return status;
	ATL::CComPtr<IUnknown> activeIdentity, targetIdentity;
	HRESULT hr = activeDocument->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&activeIdentity));
	if (SUCCEEDED(hr))
		hr = document->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&targetIdentity));
	if (IsStopping())
		return StoppedStatus();
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, "Document activation identity");
	if (activeIdentity == nullptr || activeIdentity.p != targetIdentity.p ||
		GetForegroundWindow() != GetAncestor(GetActiveWindow(), GA_ROOT))
		return ErrorStatus(-32000, E_FAIL, "Document window activation",
			"Office did not activate the requested document in the foreground");
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::DispatchHttp(HttpCommand command,
	const std::string &argument, long slideId, bool force, nlohmann::json &result)
{
	if (command == HttpCommand::Version)
		return ApplicationMetadata(result);
	if (command == HttpCommand::New)
		return OpenDocument(argument, result);
	nlohmann::json list;
	Status status = RefreshTargets(list);
	if (!status.ok())
		return status;
	if (command == HttpCommand::List)
	{
		result = std::move(list);
		return {};
	}
	std::string id = argument;
	for (char &character : id)
		if (character >= 'A' && character <= 'F')
			character += 'a' - 'A';
	auto position = std::find_if(m_targets.begin(), m_targets.end(), [&id](const DocumentTarget &target)
	{
		return target.descriptor["id"] == id;
	});
	if (position == m_targets.end())
		return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Document target",
			"Unknown or closed document target");
	// Retain both references locally before any reentrant COM call.
	ATL::CComPtr<IDispatch> document = position->document;
	ATL::CComPtr<IUnknown> identity = position->identity;
	if (command == HttpCommand::Presentation)
		return GetPresentationState(document, id, result);
	if (command == HttpCommand::Slides)
		return GetSlides(document, id, result);
	if (command == HttpCommand::Slide)
		return GetSlideState(document, id, slideId, result);
	if (command == HttpCommand::View)
		return GetViewState(document, id, result);
	if (command == HttpCommand::SlideShow)
		return GetSlideShowState(document, id, result);
	if (command == HttpCommand::Activate)
		return ActivateDocument(document);
	if (!force)
	{
		long saved = 0;
		status = GetInteger(document, L"Saved", saved);
		if (!status.ok())
			return status;
		if (saved != -1) // msoTrue
			return ErrorStatus(-32005, HRESULT_FROM_WIN32(ERROR_CANCELLED), "Closing document",
				"Document has unsaved changes; use the force query flag to discard them");
	}
	else
	{
		ATL::CComVariant saved(-1L); // Mark clean immediately before Close, without saving.
		status = Invoke(document, L"Saved", DISPATCH_PROPERTYPUT, &saved, 1, nullptr);
		if (!status.ok())
			return status;
	}
	status = Invoke(document, L"Close", DISPATCH_METHOD, nullptr, 0, nullptr);
	if (!status.ok())
		return status;
	status = RefreshTargets(list);
	if (!status.ok())
		return status;
	if (std::any_of(m_targets.begin(), m_targets.end(), [&identity](const DocumentTarget &target)
	{
		return target.identity.p == identity.p;
	}))
		return ErrorStatus(-32000, E_ABORT, "Closing document", "Office kept the document open");
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetDouble(IDispatch *object,
	const wchar_t *name, double &result)
{
	ATL::CComVariant value;
	Status status = Invoke(object, name, DISPATCH_PROPERTYGET, nullptr, 0, &value);
	if (!status.ok())
		return status;
	HRESULT hr = value.ChangeType(VT_R8);
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, MemberName(name));
	result = value.dblVal;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetBoolean(IDispatch *object,
	const wchar_t *name, bool &result)
{
	long value = 0;
	Status status = GetInteger(object, name, value);
	if (!status.ok())
		return status;
	result = value != 0;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetSlideById(IDispatch *document,
	long slideId, ATL::CComPtr<IDispatch> &slide, long &index)
{
	ATL::CComPtr<IDispatch> slides;
	Status status = GetObject(document, L"Slides", slides);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(slides, L"Count", count);
	if (!status.ok())
		return status;
	for (long position = 1; position <= count; ++position)
	{
		ATL::CComVariant argument(position);
		ATL::CComVariant value;
		status = Invoke(slides, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			&argument, 1, &value);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> candidate;
		status = ReadObject(value, L"Slides.Item", candidate, -32000);
		if (!status.ok())
			return status;
		long candidateId = 0;
		status = GetInteger(candidate, L"SlideID", candidateId);
		if (!status.ok())
			return status;
		if (candidateId == slideId)
		{
			slide = candidate;
			index = position;
			return {};
		}
	}
	return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Slide target",
		"Unknown slide ID in this presentation");
}

AutomationDispatcher::Status AutomationDispatcher::ReadSlide(IDispatch *slide,
	long slideId, long index, nlohmann::json &result, bool includeShapes)
{
	std::string name;
	Status status = GetString(slide, L"Name", name);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> transition;
	status = GetObject(slide, L"SlideShowTransition", transition);
	if (!status.ok())
		return status;
	bool hidden = false;
	status = GetBoolean(transition, L"Hidden", hidden);
	if (!status.ok())
		return status;
	result = { { "slideId", slideId }, { "slideIndex", index }, { "name", std::move(name) },
		{ "hidden", hidden } };
	if (!includeShapes)
		return {};

	ATL::CComPtr<IDispatch> shapes;
	status = GetObject(slide, L"Shapes", shapes);
	if (!status.ok())
		return status;
	long shapeCount = 0;
	status = GetInteger(shapes, L"Count", shapeCount);
	if (!status.ok())
		return status;
	auto shapeResults = nlohmann::json::array();
	for (long shapeIndex = 1; shapeIndex <= shapeCount; ++shapeIndex)
	{
		ATL::CComVariant argument(shapeIndex);
		ATL::CComVariant value;
		status = Invoke(shapes, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			&argument, 1, &value);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> shape;
		status = ReadObject(value, L"Shapes.Item", shape, -32000);
		if (!status.ok())
			return status;
		long shapeId = 0, zOrder = 0, type = 0;
		status = GetInteger(shape, L"Id", shapeId);
		if (!status.ok())
			return status;
		status = GetInteger(shape, L"ZOrderPosition", zOrder);
		if (!status.ok())
			return status;
		status = GetInteger(shape, L"Type", type);
		if (!status.ok())
			return status;
		std::string shapeName;
		status = GetString(shape, L"Name", shapeName);
		if (!status.ok())
			return status;
		double left = 0, top = 0, width = 0, height = 0;
		status = GetDouble(shape, L"Left", left);
		if (!status.ok())
			return status;
		status = GetDouble(shape, L"Top", top);
		if (!status.ok())
			return status;
		status = GetDouble(shape, L"Width", width);
		if (!status.ok())
			return status;
		status = GetDouble(shape, L"Height", height);
		if (!status.ok())
			return status;
		nlohmann::json shapeResult = {
			{ "shapeId", shapeId }, { "zOrderPosition", zOrder }, { "name", std::move(shapeName) },
			{ "shapeType", type }, { "text", nullptr },
			{ "bounds", { { "left", left }, { "top", top }, { "width", width }, { "height", height } } }
		};
		if (type == 14) // msoPlaceholder
		{
			ATL::CComPtr<IDispatch> placeholder;
			Status optional = GetObject(shape, L"PlaceholderFormat", placeholder);
			if (optional.ok())
			{
				long placeholderType = 0;
				optional = GetInteger(placeholder, L"Type", placeholderType);
				if (optional.ok())
					shapeResult["placeholderType"] = placeholderType;
			}
			else if (IsStopping())
				return StoppedStatus();
		}
		ATL::CComPtr<IDispatch> textFrame;
		Status textStatus = GetObject(shape, L"TextFrame", textFrame);
		if (textStatus.ok())
		{
			bool hasText = false;
			textStatus = GetBoolean(textFrame, L"HasText", hasText);
			if (textStatus.ok() && hasText)
			{
				ATL::CComPtr<IDispatch> textRange;
				textStatus = GetObject(textFrame, L"TextRange", textRange);
				if (textStatus.ok())
				{
					std::string text;
					textStatus = GetString(textRange, L"Text", text);
					if (textStatus.ok())
						shapeResult["text"] = std::move(text);
				}
			}
		}
		if (!textStatus.ok() && IsStopping())
			return StoppedStatus();
		shapeResults.push_back(std::move(shapeResult));
	}
	result["shapes"] = std::move(shapeResults);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetPresentationState(IDispatch *document,
	const std::string &id, nlohmann::json &result)
{
	nlohmann::json descriptor;
	Status status = DescribeDocument(document, id, descriptor);
	if (!status.ok())
		return status;
	long saved = 0, readOnly = 0, slideCount = 0;
	status = GetInteger(document, L"Saved", saved);
	if (!status.ok())
		return status;
	status = GetInteger(document, L"ReadOnly", readOnly);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slides;
	status = GetObject(document, L"Slides", slides);
	if (!status.ok())
		return status;
	status = GetInteger(slides, L"Count", slideCount);
	if (!status.ok())
		return status;
	result = { { "id", id }, { "name", descriptor["title"] }, { "url", descriptor["url"] },
		{ "saved", saved != 0 }, { "readOnly", readOnly != 0 }, { "slideCount", slideCount } };
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetSlides(IDispatch *document,
	const std::string &id, nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> slides;
	Status status = GetObject(document, L"Slides", slides);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(slides, L"Count", count);
	if (!status.ok())
		return status;
	auto slideResults = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComVariant argument(index);
		ATL::CComVariant value;
		status = Invoke(slides, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			&argument, 1, &value);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> slide;
		status = ReadObject(value, L"Slides.Item", slide, -32000);
		if (!status.ok())
			return status;
		long slideId = 0;
		status = GetInteger(slide, L"SlideID", slideId);
		if (!status.ok())
			return status;
		nlohmann::json summary;
		status = ReadSlide(slide, slideId, index, summary, false);
		if (!status.ok())
			return status;
		slideResults.push_back(std::move(summary));
	}
	result = { { "id", id }, { "slides", std::move(slideResults) } };
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetSlideState(IDispatch *document,
	const std::string &id, long slideId, nlohmann::json &result)
{
	long index = 0;
	ATL::CComPtr<IDispatch> slide;
	Status status = GetSlideById(document, slideId, slide, index);
	if (!status.ok())
		return status;
	nlohmann::json slideState;
	status = ReadSlide(slide, slideId, index, slideState, true);
	if (!status.ok())
		return status;
	result = { { "id", id }, { "slide", std::move(slideState) } };
	return {};
}
AutomationDispatcher::Status AutomationDispatcher::GetViewState(IDispatch *document,
	const std::string &id, nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> windows;
	Status status = GetObject(document, L"Windows", windows);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(windows, L"Count", count);
	if (!status.ok())
		return status;
	auto windowResults = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComVariant argument(index);
		ATL::CComVariant value;
		status = Invoke(windows, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			&argument, 1, &value);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> window;
		status = ReadObject(value, L"Windows.Item", window, -32000);
		if (!status.ok())
			return status;
		long viewType = 0;
		status = GetInteger(window, L"ViewType", viewType);
		if (!status.ok())
			return status;
		nlohmann::json windowResult = { { "viewType", viewType }, { "currentSlideId", nullptr },
			{ "currentSlideIndex", nullptr }, { "selection", nullptr } };
		ATL::CComPtr<IDispatch> view;
		Status optional = GetObject(window, L"View", view);
		if (optional.ok())
		{
			ATL::CComPtr<IDispatch> currentSlide;
			optional = GetObject(view, L"Slide", currentSlide);
			if (optional.ok())
			{
				long currentSlideId = 0, currentSlideIndex = 0;
				optional = GetInteger(currentSlide, L"SlideID", currentSlideId);
				if (optional.ok())
					optional = GetInteger(currentSlide, L"SlideIndex", currentSlideIndex);
				if (optional.ok())
				{
					windowResult["currentSlideId"] = currentSlideId;
					windowResult["currentSlideIndex"] = currentSlideIndex;
				}
			}
		}
		if (IsStopping())
			return StoppedStatus();
		ATL::CComPtr<IDispatch> selection;
		optional = GetObject(window, L"Selection", selection);
		if (optional.ok())
		{
			long selectionType = 0;
			optional = GetInteger(selection, L"Type", selectionType);
			if (optional.ok())
			{
				nlohmann::json selectionResult = { { "type", selectionType } };
				if (selectionType == 2) // ppSelectionShapes
				{
					ATL::CComPtr<IDispatch> range;
					optional = GetObject(selection, L"ShapeRange", range);
					long selectedCount = 0;
					if (optional.ok())
						optional = GetInteger(range, L"Count", selectedCount);
					if (optional.ok())
					{
						auto selectedShapes = nlohmann::json::array();
						for (long selectedIndex = 1; selectedIndex <= selectedCount; ++selectedIndex)
						{
							ATL::CComVariant selectedArgument(selectedIndex);
							ATL::CComVariant selectedValue;
							optional = Invoke(range, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
								&selectedArgument, 1, &selectedValue);
							if (!optional.ok())
								break;
							ATL::CComPtr<IDispatch> selectedShape;
							optional = ReadObject(selectedValue, L"ShapeRange.Item", selectedShape, -32000);
							if (!optional.ok())
								break;
							long selectedShapeId = 0;
							optional = GetInteger(selectedShape, L"Id", selectedShapeId);
							if (!optional.ok())
								break;
							selectedShapes.push_back(selectedShapeId);
						}
						if (optional.ok())
							selectionResult["shapeIds"] = std::move(selectedShapes);
					}
				}
				windowResult["selection"] = std::move(selectionResult);
			}
		}
		if (IsStopping())
			return StoppedStatus();
		windowResults.push_back(std::move(windowResult));
	}
	result = { { "id", id }, { "available", count > 0 }, { "windows", std::move(windowResults) } };
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetSlideShowState(IDispatch *document,
	const std::string &id, nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> showWindows;
	Status status = GetObject(application, L"SlideShowWindows", showWindows);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(showWindows, L"Count", count);
	if (!status.ok())
		return status;
	ATL::CComPtr<IUnknown> targetIdentity;
	HRESULT hr = document->QueryInterface(IID_IUnknown, reinterpret_cast<void **>(&targetIdentity));
	if (IsStopping())
		return StoppedStatus();
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, "Presentation identity");
	auto showResults = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComVariant argument(index);
		ATL::CComVariant value;
		status = Invoke(showWindows, L"Item", DISPATCH_METHOD | DISPATCH_PROPERTYGET,
			&argument, 1, &value);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> showWindow;
		status = ReadObject(value, L"SlideShowWindows.Item", showWindow, -32000);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> showPresentation;
		status = GetObject(showWindow, L"Presentation", showPresentation);
		if (!status.ok())
			return status;
		ATL::CComPtr<IUnknown> showIdentity;
		hr = showPresentation->QueryInterface(IID_IUnknown,
			reinterpret_cast<void **>(&showIdentity));
		if (IsStopping())
			return StoppedStatus();
		if (FAILED(hr))
			return ErrorStatus(-32000, hr, "Slide-show presentation identity");
		if (showIdentity.p != targetIdentity.p)
			continue;
		ATL::CComPtr<IDispatch> view;
		status = GetObject(showWindow, L"View", view);
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> currentSlide;
		status = GetObject(view, L"Slide", currentSlide);
		if (!status.ok())
			return status;
		long currentSlideId = 0, currentSlideIndex = 0, showPosition = 0;
		status = GetInteger(currentSlide, L"SlideID", currentSlideId);
		if (!status.ok())
			return status;
		status = GetInteger(currentSlide, L"SlideIndex", currentSlideIndex);
		if (!status.ok())
			return status;
		status = GetInteger(view, L"CurrentShowPosition", showPosition);
		if (!status.ok())
			return status;
		showResults.push_back({ { "currentSlideId", currentSlideId },
			{ "currentSlideIndex", currentSlideIndex }, { "currentShowPosition", showPosition } });
	}
	result = { { "id", id }, { "running", !showResults.empty() }, { "windows", std::move(showResults) } };
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
