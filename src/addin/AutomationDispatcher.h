#pragma once

#include "pch.h"

#include "PowerPointCommands.h"

#include <nlohmann/json.hpp>
#include <optional>
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
		StartSlideShow, StopSlideShow, NavigateSlideShow,
#define NETOFFICE_COMMAND_ENUM(name, method, verb, scope, suffix) name,
		NETOFFICE_POWERPOINT_COMMANDS(NETOFFICE_COMMAND_ENUM)
#undef NETOFFICE_COMMAND_ENUM
	};
	struct CommandInfo
	{
		HttpCommand command;
		const char *method;
		const char *verb;
		CommandScope scope;
		const char *suffix;
	};
	// Every table-driven command, in PowerPointCommands.h order.
	static const std::vector<CommandInfo> &Commands();
	static const CommandInfo *FindCommand(HttpCommand command);
	static const CommandInfo *FindCommand(const std::string &method);
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
	// Resolved Office objects for one table-driven command. The parameters are the
	// complete request params, including routing members (targetId, slideId, shapeId, master).
	struct CommandContext
	{
		IDispatch *document = nullptr; // null for Application scope
		const std::string &id;
		const nlohmann::json &parameters;
		bool master = false;           // container is the slide master
		long customLayout = 0;         // container is SlideMaster.CustomLayouts(customLayout)
		long slideId = 0;              // 0 for the master or a custom layout
		long slideIndex = 0;           // 0 for the master or a custom layout
		IDispatch *container = nullptr; // Slide, Container, and Shape scopes: slide, SlideMaster, or CustomLayout
		long shapeId = 0;
		IDispatch *shape = nullptr;    // Shape scope
	};

	// Table-driven handlers. The result arrives prefilled with the routing identity
	// (id, slideId or master, shapeId). Mutations then receive every accepted
	// non-routing parameter that the handler did not set itself.
#define NETOFFICE_COMMAND_HANDLER(name, method, verb, scope, suffix) \
	Status name(const CommandContext &context, nlohmann::json &result);
	NETOFFICE_POWERPOINT_COMMANDS(NETOFFICE_COMMAND_HANDLER)
#undef NETOFFICE_COMMAND_HANDLER
	Status RunCommand(const CommandInfo &info, IDispatch *document, const std::string &id,
		long slideId, long shapeId, const nlohmann::json &parameters, nlohmann::json &result);
	// Exactly one selector: slideId, master, or a 1-based customLayout index.
	Status ResolveContainer(IDispatch *document, long slideId, bool master, long customLayout,
		ATL::CComPtr<IDispatch> &container, long &slideIndex);
	// SlideMaster.CustomLayouts(index), -32004 when the index exceeds the layout count.
	Status GetCustomLayout(IDispatch *document, long index, ATL::CComPtr<IDispatch> &layout);
	// The first CustomLayouts entry whose Name equals name exactly; -32004 when absent.
	Status FindCustomLayout(IDispatch *document, const std::string &name,
		ATL::CComPtr<IDispatch> &layout, long &index);

	// Parameter readers: absent optional members leave the output empty; invalid or
	// absent required members return -32602 naming the member.
	static Status InvalidParameter(const std::string &message);
	static Status ReadNumber(const nlohmann::json &parameters, const char *name,
		std::optional<double> &value, double minimum, double maximum, bool required = false);
	static Status ReadInteger(const nlohmann::json &parameters, const char *name,
		std::optional<long> &value, long minimum, long maximum, bool required = false);
	static Status ReadBoolean(const nlohmann::json &parameters, const char *name,
		std::optional<bool> &value, bool required = false);
	static Status ReadString(const nlohmann::json &parameters, const char *name,
		std::optional<std::string> &value, bool required = false, bool allowEmpty = true);
	// Absolute local or UNC path without NUL characters.
	static Status ReadPath(const nlohmann::json &parameters, const char *name,
		std::optional<std::string> &value, bool required = false);
	// "#RRGGBB" -> Office RGB (red | green << 8 | blue << 16).
	static Status ReadColor(const nlohmann::json &parameters, const char *name,
		std::optional<long> &value, bool required = false);
	static Status ReadChoice(const nlohmann::json &parameters, const char *name,
		std::initializer_list<std::pair<const char *, long>> choices,
		std::optional<long> &value, bool required = false);
	// Array of distinct positive int32 shape IDs with at least minimumCount entries.
	static Status ReadShapeIds(const nlohmann::json &parameters, const char *name,
		std::vector<long> &value, size_t minimumCount);
	static std::string ColorToHex(long rgb);
	static HRESULT Utf8ToVariant(const std::string &text, ATL::CComVariant &result);
	static HRESULT BstrToUtf8(BSTR text, std::string &result);
	// VT_ERROR/DISP_E_PARAMNOTFOUND: an omitted optional positional argument.
	static ATL::CComVariant MissingArgument();

	// COM helpers for handlers. CallMethod takes arguments in declaration order.
	Status SetProperty(IDispatch *object, const wchar_t *name, const ATL::CComVariant &value);
	Status CallMethod(IDispatch *object, const wchar_t *name,
		std::vector<ATL::CComVariant> arguments, ATL::CComVariant *result = nullptr);
	Status CallObject(IDispatch *object, const wchar_t *name,
		std::vector<ATL::CComVariant> arguments, ATL::CComPtr<IDispatch> &result);
	Status GetItem(IDispatch *collection, long index, ATL::CComPtr<IDispatch> &item);
	// Shapes.Range over the container's shapes with these IDs, in the given order.
	Status GetShapeRange(IDispatch *container, const std::vector<long> &shapeIds,
		ATL::CComPtr<IDispatch> &range);
	// The slide-show shape summary: shapeId, zOrderPosition, name, shapeType, bounds,
	// text, optional placeholderType, and recursive groupItems for groups.
	Status DescribeShape(IDispatch *shape, nlohmann::json &result);

	// TextCommands.cpp helpers
	// The shape's TextFrame; -32602 when the shape has no text frame.
	Status GetShapeTextFrame(IDispatch *shape, ATL::CComPtr<IDispatch> &textFrame);
	Status DescribeTextParagraph(IDispatch *paragraph, long index, nlohmann::json &result);
	// Optional state reads: a failed read omits target[key]; only a stopping add-in fails.
	Status ReadStateInteger(IDispatch *object, const wchar_t *name, nlohmann::json &target,
		const char *key, nlohmann::json (*convert)(long));
	Status ReadStateNumber(IDispatch *object, const wchar_t *name, nlohmann::json &target,
		const char *key);
	Status ReadStateString(IDispatch *object, const wchar_t *name, nlohmann::json &target,
		const char *key);
	// object.<name>.RGB as "#RRGGBB".
	Status ReadStateColor(IDispatch *object, const wchar_t *name, nlohmann::json &target,
		const char *key);
	// DataCommands.cpp helpers
	// left, top: -10000..10000; width, height: 0.01..10000; all required.
	static Status ReadDataBounds(const nlohmann::json &parameters, double &left, double &top,
		double &width, double &height);
	// Reads a Shapes.AddXxx result and stores its shapeId and name in result.
	Status ReadAddedDataShape(ATL::CComVariant &value, const wchar_t *operation,
		ATL::CComPtr<IDispatch> &shape, nlohmann::json &result);
	// Returns shape.<member> when shape.<flag> is true, else -32602 naming kind.
	Status GetDataObject(IDispatch *shape, const wchar_t *flag, const wchar_t *member,
		const char *kind, ATL::CComPtr<IDispatch> &result);

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
