#pragma once

#include "pch.h"
#include "netoffice.grpc.pb.h"

#include <grpcpp/grpcpp.h>
#include <deque>
#include <memory>
#include <mutex>
#include <string>

// Start and Stop run on the Office STA. RPC threads only use the request queue.
class GrpcServer :
	public netoffice::PowerPoint::Service,
	public std::enable_shared_from_this<GrpcServer>
{
public:
	GrpcServer() = default;
	~GrpcServer();

	HRESULT Start(IDispatch *pApplication);
	void Stop();

	grpc::Status GetStatus(grpc::ServerContext *pContext,
		const netoffice::Empty *pRequest, netoffice::StatusReply *pReply) override;
	grpc::Status NewPresentation(grpc::ServerContext *pContext,
		const netoffice::NewPresentationRequest *pRequest,
		netoffice::PresentationReply *pReply) override;
	grpc::Status SetSlideTitle(grpc::ServerContext *pContext,
		const netoffice::SetSlideTitleRequest *pRequest, netoffice::Empty *pReply) override;

private:
	struct PendingCall;

	static LRESULT CALLBACK WindowProc(HWND hWnd, UINT message, WPARAM wParam, LPARAM lParam);
	void DispatchOnSta();
	grpc::Status Dispatch(const std::shared_ptr<PendingCall> &call, grpc::ServerContext *pContext);
	bool IsStopping();

	grpc::Status CreatePresentation(const std::string &title, netoffice::PresentationReply *pReply);
	grpc::Status UpdateSlideTitle(int slideIndex, const std::string &text);
	grpc::Status PutSlideTitle(IDispatch *pSlide, ATL::CComVariant &text);
	grpc::Status Invoke(IDispatch *pObject, const wchar_t *pName, WORD flags,
		ATL::CComVariant *pArguments, UINT argumentCount, ATL::CComVariant *pResult,
		grpc::StatusCode failureCode = grpc::StatusCode::INTERNAL);
	grpc::Status GetObject(IDispatch *pObject, const wchar_t *pName,
		ATL::CComPtr<IDispatch> &result,
		grpc::StatusCode failureCode = grpc::StatusCode::INTERNAL);
	grpc::Status GetInteger(IDispatch *pObject, const wchar_t *pName, long &result);
	grpc::Status ReadObject(ATL::CComVariant &value, const wchar_t *pName,
		ATL::CComPtr<IDispatch> &result, grpc::StatusCode failureCode);

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
	std::unique_ptr<grpc::Server> m_server;
};
