#pragma once

#include "pch.h"

#include <nlohmann/json.hpp>
#include <atomic>
#include <chrono>
#include <deque>
#include <memory>
#include <mutex>
#include <string>
#include <vector>

// Start/Stop and every Automation reference belong to the Office STA.
// HTTP/WebSocket workers may only call HandleRequest or HandleHttpRequest.
class AutomationDispatcher : public std::enable_shared_from_this<AutomationDispatcher>
{
public:
	AutomationDispatcher() = default;
	enum class HttpCommand {
		Version, List, New, NamedNew, Activate, Close, Presentation, Slides, Slide, View, SlideShow,
		AddSlide, AddShape, SetShapeText, DeleteShape, DeleteSlide, SetCurrentSlide,
		StartSlideShow, StopSlideShow, NavigateSlideShow
	};
	~AutomationDispatcher();

	HRESULT Start(IDispatch *app);
	void Stop();
	nlohmann::json HandleRequest(const nlohmann::json &request,
		const std::shared_ptr<std::atomic_bool> &cancelled);
	nlohmann::json HandleHttpRequest(HttpCommand command, const std::string &argument,
		long slideId, long shapeId, bool force, const nlohmann::json &parameters,
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
	struct DocumentTarget
	{
		ATL::CComPtr<IUnknown> identity;
		ATL::CComPtr<IDispatch> document;
		nlohmann::json descriptor;
	};

	static Status ErrorStatus(int code, HRESULT hr, const std::string &operation,
		const std::string &details = std::string());
	static Status StoppedStatus();
	static Status CancellationStatus(std::chrono::steady_clock::time_point deadline);
	static nlohmann::json ErrorReply(const nlohmann::json &id, const Status &status);
	static LRESULT CALLBACK WindowProc(HWND window, UINT message, WPARAM wParam, LPARAM lParam);
	void DispatchOnSta();
	Status Dispatch(const std::shared_ptr<PendingCall> &call);
	bool IsStopping();

	Status DispatchHttp(HttpCommand command, const std::string &argument, long slideId,
		long shapeId, bool force, const nlohmann::json &parameters, nlohmann::json &result);
	Status GetPresentationState(IDispatch *document, const std::string &id, nlohmann::json &result);
	Status GetSlides(IDispatch *document, const std::string &id, nlohmann::json &result);
	Status GetSlideState(IDispatch *document, const std::string &id, long slideId, nlohmann::json &result);
	Status GetViewState(IDispatch *document, const std::string &id, nlohmann::json &result);
	Status GetSlideShowState(IDispatch *document, const std::string &id, nlohmann::json &result);
	Status ReadSlide(IDispatch *slide, long slideId, long index, nlohmann::json &result, bool includeShapes);
	Status GetDouble(IDispatch *object, const wchar_t *name, double &result);
	Status GetBoolean(IDispatch *object, const wchar_t *name, bool &result);
	Status GetSlideById(IDispatch *document, long slideId, ATL::CComPtr<IDispatch> &slide, long &index);
	Status MutatePresentation(HttpCommand command, IDispatch *document, const std::string &id,
		long slideId, long shapeId, const nlohmann::json &parameters, nlohmann::json &result);
	Status GetShapeById(IDispatch *slide, long shapeId, ATL::CComPtr<IDispatch> &shape);
	Status ApplicationMetadata(nlohmann::json &result);
	Status RefreshTargets(nlohmann::json &result);
	Status DescribeDocument(IDispatch *document, const std::string &id, nlohmann::json &result);
	Status OpenDocument(const std::string &argument, nlohmann::json &result);
	Status ActivateDocument(IDispatch *document);
	Status GetString(IDispatch *object, const wchar_t *name, std::string &result);
	Status GetExecutableBuild(IDispatch *application, std::string &result);
	Status CreatePresentation(const std::string &name, const std::string &directory,
		nlohmann::json &result);
	Status ValidateShutdown(bool force);
	Status Invoke(IDispatch *object, const wchar_t *name, WORD flags,
		ATL::CComVariant *arguments, UINT argumentCount, ATL::CComVariant *result,
		int failureCode = -32000);
	Status GetObject(IDispatch *object, const wchar_t *name,
		ATL::CComPtr<IDispatch> &result, int failureCode = -32000);
	Status GetInteger(IDispatch *object, const wchar_t *name, long &result);
	Status ReadObject(ATL::CComVariant &value, const wchar_t *name,
		ATL::CComPtr<IDispatch> &result, int failureCode);

	ATL::CComPtr<IDispatch> m_pApplication;
	// Canonical IUnknown identities and their documents are owned/released only on STA.
	std::vector<DocumentTarget> m_targets;
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
