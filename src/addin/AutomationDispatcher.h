#pragma once

#include "pch.h"

#include <nlohmann/json.hpp>
#include <atomic>
#include <chrono>
#include <deque>
#include <memory>
#include <mutex>
#include <string>

// Start/Stop and every Automation reference belong to the Office STA.
// HTTP/WebSocket workers may only call HandleRequest.
class AutomationDispatcher : public std::enable_shared_from_this<AutomationDispatcher>
{
public:
	AutomationDispatcher() = default;
	~AutomationDispatcher();

	HRESULT Start(IDispatch *app);
	void Stop();
	nlohmann::json HandleRequest(const nlohmann::json &request,
		const std::shared_ptr<std::atomic_bool> &cancelled);

private:
	struct Status
	{
		int code = 0;
		std::string message;
		nlohmann::json data;
		bool ok() const { return code == 0; }
	};
	struct PendingCall;

	static Status ErrorStatus(int code, HRESULT hr, const std::string &operation,
		const std::string &details = std::string());
	static Status StoppedStatus();
	static Status CancellationStatus(std::chrono::steady_clock::time_point deadline);
	static nlohmann::json ErrorReply(const nlohmann::json &id, const Status &status);
	static LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
	void DispatchOnSta();
	Status Dispatch(const std::shared_ptr<PendingCall> &call);
	bool IsStopping();

	Status CreatePresentation(const std::string &title, nlohmann::json &result);
	Status UpdateSlideTitle(int slideIndex, const std::string &text);
	Status PutSlideTitle(IDispatch *slide, ATL::CComVariant &text);
	Status Invoke(IDispatch *object, const wchar_t *name, WORD flags,
		ATL::CComVariant *arguments, UINT argumentCount, ATL::CComVariant *result,
		int failureCode = -32000);
	Status GetObject(IDispatch *object, const wchar_t *name,
		ATL::CComPtr<IDispatch> &result, int failureCode = -32000);
	Status GetInteger(IDispatch *object, const wchar_t *name, long &result);
	Status ReadObject(ATL::CComVariant &value, const wchar_t *name,
		ATL::CComPtr<IDispatch> &result, int failureCode);

	ATL::CComPtr<IDispatch> m_pApplication;
	DWORD m_ownerThreadId = 0;
	HINSTANCE m_hInstance = nullptr;
	ATOM m_windowClass = 0;
	HWND m_hWnd = nullptr;
	bool m_dispatching = false;

	std::mutex m_mutex;
	bool m_stopping = true;
	std::deque<std::shared_ptr<PendingCall>> m_pendingCalls;
	std::shared_ptr<PendingCall> m_activeCall;
};
