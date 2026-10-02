#pragma once

#include "pch.h"

#include <memory>
#include <mutex>

// Start and Stop belong to the Office STA. Transport threads never call Office COM.
class HttpServer : public std::enable_shared_from_this<HttpServer>
{
public:
	HttpServer() = default;
	~HttpServer();

	HRESULT Start(IDispatch *app, unsigned short port);
	void Stop();

private:
	struct State;
	std::mutex m_mutex;
	std::shared_ptr<State> m_state;
	bool m_stopping = false;
};
