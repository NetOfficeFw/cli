#include "pch.h"
#include "HttpServer.h"
#include "AutomationDispatcher.h"

#include <civetweb.h>
#include <nlohmann/json.hpp>

#include <algorithm>
#include <array>
#include <atomic>
#include <condition_variable>
#include <cstdint>
#include <deque>
#include <exception>
#include <new>
#include <string>
#include <string_view>
#include <system_error>
#include <thread>
#include <unordered_map>
#include <utility>
#include <vector>

namespace
{
	using Json = nlohmann::json;
	constexpr size_t MaxMessageBytes = 1024 * 1024;
	constexpr size_t MaxSessions = 32;
	constexpr size_t MaxQueuedMessages = 16;
	constexpr size_t MaxQueuedBytes = 4 * MaxMessageBytes;
	constexpr char NativeFailure[] = "{\"error\":{\"code\":-32000,\"message\":\"Native request handling failed\"}}";
	std::mutex LibraryMutex;

	bool EqualAscii(std::string_view left, std::string_view right)
	{
		if (left.size() != right.size()) return false;
		for (size_t i = 0; i < left.size(); ++i)
		{
			char a = left[i], b = right[i];
			if (a >= 'A' && a <= 'Z') a += 'a' - 'A';
			if (b >= 'A' && b <= 'Z') b += 'a' - 'A';
			if (a != b) return false;
		}
		return true;
	}

	std::string_view Trim(std::string_view value)
	{
		while (!value.empty() && (value.front() == ' ' || value.front() == '\t')) value.remove_prefix(1);
		while (!value.empty() && (value.back() == ' ' || value.back() == '\t')) value.remove_suffix(1);
		return value;
	}

	std::string_view Header(const mg_connection *connection, const char *name)
	{
		const char *value = mg_get_header(connection, name);
		return value == nullptr ? std::string_view() : std::string_view(value);
	}

	bool HasToken(std::string_view values, std::string_view token)
	{
		while (!values.empty())
		{
			size_t comma = values.find(',');
			if (EqualAscii(Trim(values.substr(0, comma)), token)) return true;
			if (comma == std::string_view::npos) break;
			values.remove_prefix(comma + 1);
		}
		return false;
	}

	bool AllowedOrigin(const mg_request_info &request, unsigned short port)
	{
		bool seen = false;
		const std::string suffix = ":" + std::to_string(port);
		for (int i = 0; i < request.num_headers; ++i)
		{
			const auto &header = request.http_headers[i];
			if (!EqualAscii(header.name, "Origin")) continue;
			if (seen) return false;
			seen = true;
			std::string_view origin(header.value);
			if (origin != "http://127.0.0.1" + suffix && origin != "http://localhost" + suffix)
				return false;
		}
		return true;
	}

	Json RequestId(const Json &request)
	{
		if (!request.is_object()) return nullptr;
		auto id = request.find("id");
		if (id == request.end()) return nullptr;
		if (id->is_number_unsigned())
			return id->get<uint64_t>() <= 9007199254740991ULL ? *id : Json(nullptr);
		if (id->is_number_integer())
		{
			int64_t value = id->get<int64_t>();
			if (value >= 0 && value <= 9007199254740991LL) return *id;
		}
		return nullptr;
	}

	Json ErrorReply(const Json &id, int code, const char *message)
	{
		return Json{{"id", id}, {"error", {{"code", code}, {"message", message}}}};
	}

	Json HttpErrorReply(int code, const char *message)
	{
		return Json{{"error", {{"code", code}, {"message", message}}}};
	}

	enum class Endpoint { Unknown, Rpc, WebSocket, Version, List, New, Activate, Close };

	std::string_view RequestPath(const mg_request_info &request)
	{
		// Automatic URI decoding is disabled: retain the unnormalized path so
		// malformed targets cannot be cleaned or truncated into another route.
		return request.local_uri_raw == nullptr ? std::string_view() : request.local_uri_raw;
	}

	Endpoint IdentifyEndpoint(std::string_view path)
	{
		if (path == "/json/rpc") return Endpoint::Rpc;
		if (path == "/devtools/application") return Endpoint::WebSocket;
		if (path == "/json/version") return Endpoint::Version;
		if (path == "/json" || path == "/json/list") return Endpoint::List;
		if (path == "/json/new") return Endpoint::New;
		if (path == "/json/activate" || path.substr(0, 15) == "/json/activate/") return Endpoint::Activate;
		if (path == "/json/close" || path.substr(0, 12) == "/json/close/") return Endpoint::Close;
		return Endpoint::Unknown;
	}

	const char *EndpointMethod(Endpoint endpoint)
	{
		if (endpoint == Endpoint::Rpc) return "POST";
		if (endpoint == Endpoint::New || endpoint == Endpoint::Activate || endpoint == Endpoint::Close) return "PUT";
		return "GET";
	}

	bool IsHex(char value)
	{
		return (value >= '0' && value <= '9') || (value >= 'a' && value <= 'f') ||
			(value >= 'A' && value <= 'F');
	}

	bool ParseHttpQuery(const char *query, std::string &url, bool &force, const char *&error)
	{
		if (query == nullptr || *query == '\0') return true;
		std::string data(query);
		if (data.size() > MaxMessageBytes)
		{
			error = "Query exceeds the 1 MiB limit";
			return false;
		}
		// CivetWeb's decoder preserves malformed percent escapes and permits NUL.
		// Validate only the encoding here; the library splits and decodes the query.
		for (size_t i = 0; i < data.size(); ++i)
		{
			if (data[i] != '%') continue;
			if (data.size() - i < 3 || !IsHex(data[i + 1]) || !IsHex(data[i + 2]))
			{
				error = "Malformed percent encoding in query";
				return false;
			}
			if (data[i + 1] == '0' && data[i + 2] == '0')
			{
				error = "Query must not contain a NUL character";
				return false;
			}
			i += 2;
		}
		const int count = mg_split_form_urlencoded(data.data(), nullptr, 0);
		if (count <= 0)
		{
			error = "Unable to parse query";
			return false;
		}
		std::vector<mg_header> fields(static_cast<size_t>(count));
		const int parsed = mg_split_form_urlencoded(data.data(), fields.data(), static_cast<unsigned>(fields.size()));
		if (parsed < 0)
		{
			error = "Unable to parse query";
			return false;
		}
		bool hasUrl = false;
		for (int i = 0; i < parsed; ++i)
		{
			const auto &field = fields[static_cast<size_t>(i)];
			if (std::string_view(field.name) == "url")
			{
				if (hasUrl) { error = "Duplicate url query parameter"; return false; }
				hasUrl = true;
				url = field.value == nullptr ? "" : field.value;
			}
			else if (std::string_view(field.name) == "force")
			{
				if (force) { error = "Duplicate force query parameter"; return false; }
				// Presence authorizes discard, including ?force and ?force=false.
				force = true;
			}
		}
		return true;
	}

	int HttpErrorStatus(int code)
	{
		switch (code)
		{
		case -32602: return 400;
		case -32004: return 404;
		case -32005: return 409;
		case -32002: return 504;
		case -32003: return 503;
		default: return 500;
		}
	}

	// This is the documented CivetWeb response API, including upgrade denials.
	// v1.16's response-header builders reject connections already marked as WebSocket,
	// even before the handshake; mg_printf/mg_write remain available in begin_request.
	int SendHttp(mg_connection *connection, int status, std::string_view body, const char *allow = nullptr)
	{
		mg_disable_connection_keep_alive(connection);
		if (mg_printf(connection,
			"HTTP/1.1 %d %s\r\nContent-Type: application/json; charset=utf-8\r\n"
			"Content-Length: %u\r\nCache-Control: no-store\r\nConnection: close\r\n%s%s%s\r\n",
			status, mg_get_response_code_text(connection, status), static_cast<unsigned>(body.size()),
			allow == nullptr ? "" : "Allow: ", allow == nullptr ? "" : allow,
			allow == nullptr ? "" : "\r\n") > 0)
			mg_write(connection, body.data(), body.size());
		return status;
	}

	int ProtocolError(mg_connection *connection, int status, const char *message, const char *allow = nullptr)
	{
		const bool rpc = IdentifyEndpoint(RequestPath(*mg_get_request_info(connection))) == Endpoint::Rpc;
		return SendHttp(connection, status,
			(rpc ? ErrorReply(nullptr, -32600, message) : HttpErrorReply(-32600, message)).dump(), allow);
	}

	std::string HandleEnvelope(const std::string &text, const std::shared_ptr<AutomationDispatcher> &dispatcher,
		const std::shared_ptr<std::atomic_bool> &cancelled)
	{
		Json request, reply;
		try
		{
			request = Json::parse(text);
			reply = dispatcher->HandleRequest(request, cancelled);
		}
		catch (const Json::parse_error &)
		{
			reply = ErrorReply(nullptr, -32700, "Malformed JSON");
		}
		catch (const std::exception &)
		{
			reply = ErrorReply(RequestId(request), -32000, "Native request handling failed");
		}
		std::string output = reply.dump();
		if (output.size() > MaxMessageBytes)
			output = ErrorReply(RequestId(request), -32000, "Reply exceeds the 1 MiB limit").dump();
		return output;
	}

	struct Session : std::enable_shared_from_this<Session>
	{
		Session(mg_connection *connection, std::shared_ptr<AutomationDispatcher> dispatcher)
			: connection(connection), dispatcher(std::move(dispatcher)),
			cancelled(std::make_shared<std::atomic_bool>(false)) {}

		mg_connection *connection;
		std::shared_ptr<AutomationDispatcher> dispatcher;
		std::shared_ptr<std::atomic_bool> cancelled;
		std::thread commandThread, senderThread;
		std::mutex mutex, sendMutex;
		std::condition_variable changed;
		std::deque<std::string> requests, replies;
		size_t requestBytes = 0, replyBytes = 0;
		bool ready = false;
		std::atomic_bool closed{false};
		// Only the CivetWeb receiver callback accesses fragment state.
		bool fragmented = false;
		std::string message;

		void Cancel()
		{
			cancelled->store(true);
			{
				std::lock_guard<std::mutex> lock(mutex);
				requests.clear();
				replies.clear();
				requestBytes = replyBytes = 0;
			}
			changed.notify_all();
		}

		void CloseAndJoin()
		{
			if (closed.exchange(true)) return;
			Cancel();
			{
				// Do not let a sender retain a CivetWeb connection after its close callback.
				// The library's request timeout bounds an already active socket write.
				std::lock_guard<std::mutex> lock(sendMutex);
				connection = nullptr;
			}
			if (senderThread.joinable()) senderThread.join();
			if (commandThread.joinable()) commandThread.join();
		}

		void StartWorkers()
		{
			auto self = shared_from_this();
			try
			{
				commandThread = std::thread([self]
				{
					try { self->CommandLoop(); }
					catch (...) { self->Cancel(); }
				});
				senderThread = std::thread([self]
				{
					try { self->SenderLoop(); }
					catch (...) { self->Cancel(); }
				});
			}
			catch (...)
			{
				CloseAndJoin();
				throw;
			}
		}

		void MakeReady()
		{
			{
				std::lock_guard<std::mutex> lock(mutex);
				ready = true;
			}
			changed.notify_all();
		}

		bool Enqueue(std::string &text)
		{
			std::lock_guard<std::mutex> lock(mutex);
			if (cancelled->load() || requests.size() >= MaxQueuedMessages ||
				text.size() > MaxQueuedBytes - requestBytes) return false;
			requestBytes += text.size();
			requests.push_back(std::move(text));
			changed.notify_all();
			return true;
		}

		bool QueueReply(std::string text, bool wait)
		{
			std::unique_lock<std::mutex> lock(mutex);
			auto available = [&]
			{
				return cancelled->load() ||
					(replies.size() < MaxQueuedMessages && text.size() <= MaxQueuedBytes - replyBytes);
			};
			if (wait) changed.wait(lock, available);
			if (cancelled->load() || !available()) return false;
			replyBytes += text.size();
			replies.push_back(std::move(text));
			changed.notify_all();
			return true;
		}

		void CommandLoop()
		{
			for (;;)
			{
				std::string text;
				{
					std::unique_lock<std::mutex> lock(mutex);
					changed.wait(lock, [&] { return cancelled->load() || !requests.empty(); });
					if (cancelled->load()) return;
					text = std::move(requests.front());
					requests.pop_front();
					requestBytes -= text.size();
				}
				if (!QueueReply(HandleEnvelope(text, dispatcher, cancelled), true)) return;
			}
		}

		void SenderLoop()
		{
			for (;;)
			{
				std::string text;
				{
					std::unique_lock<std::mutex> lock(mutex);
					changed.wait(lock, [&] { return cancelled->load() || (ready && !replies.empty()); });
					if (cancelled->load()) return;
					text = std::move(replies.front());
					replies.pop_front();
					replyBytes -= text.size();
				}
				changed.notify_all();
				bool sent;
				{
					std::lock_guard<std::mutex> lock(sendMutex);
					if (cancelled->load() || connection == nullptr) return;
					sent = mg_websocket_write(connection, MG_WEBSOCKET_OPCODE_TEXT, text.data(), text.size()) > 0;
				}
				if (!sent) { Cancel(); return; }
			}
		}

		int Close(unsigned short status, std::string_view reason = {})
		{
			std::array<char, 125> payload{};
			payload[0] = static_cast<char>(status >> 8);
			payload[1] = static_cast<char>(status & 255);
			const size_t length = (std::min)(reason.size(), payload.size() - 2);
			std::copy_n(reason.data(), length, payload.data() + 2);
			Cancel();
			std::lock_guard<std::mutex> lock(sendMutex);
			if (connection != nullptr)
				mg_websocket_write(connection, MG_WEBSOCKET_OPCODE_CONNECTION_CLOSE, payload.data(), length + 2);
			return 0;
		}

		int Receive(int bits, const char *data, size_t length)
		{
			if (cancelled->load()) return 0;
			const int opcode = bits & 15;
			const bool final = (bits & 128) != 0;
			if ((bits & 112) != 0 || (opcode >= 8 && (!final || length > 125)))
				return Close(1002);
			if (opcode == MG_WEBSOCKET_OPCODE_CONNECTION_CLOSE)
			{
				if (length == 1) return Close(1002);
				Cancel();
				std::lock_guard<std::mutex> lock(sendMutex);
				if (connection != nullptr)
					mg_websocket_write(connection, MG_WEBSOCKET_OPCODE_CONNECTION_CLOSE, data, length);
				return 0;
			}
			if (opcode == MG_WEBSOCKET_OPCODE_PING || opcode == MG_WEBSOCKET_OPCODE_PONG)
				return 1; // CivetWeb handles ping/pong with enable_websocket_ping_pong.
			if (opcode == MG_WEBSOCKET_OPCODE_BINARY)
				return Close(1003, "JSON requests must be UTF-8 text");
			if ((opcode == MG_WEBSOCKET_OPCODE_TEXT && fragmented) ||
				(opcode == MG_WEBSOCKET_OPCODE_CONTINUATION && !fragmented) ||
				(opcode != MG_WEBSOCKET_OPCODE_TEXT && opcode != MG_WEBSOCKET_OPCODE_CONTINUATION))
				return Close(1002);
			if (length > MaxMessageBytes - message.size())
				return Close(1009, "Message exceeds the 1 MiB limit");
			if (length != 0) message.append(data, length);
			fragmented = !final;
			if (final)
			{
				if (!Enqueue(message))
				{
					Json request = Json::parse(message, nullptr, false);
					if (!QueueReply(ErrorReply(RequestId(request), -32000, "Session request queue is full").dump(), false))
						return Close(1008, "Client is not consuming replies");
				}
				message.clear();
			}
			return cancelled->load() ? 0 : 1;
		}
	};
}

struct HttpServer::State
{
	std::shared_ptr<AutomationDispatcher> dispatcher;
	mg_context *context = nullptr;
	unsigned short port = 0;
	bool libraryInitialized = false;
	std::atomic_bool stopping{false}, ready{false};
	std::mutex sessionsMutex;
	std::unordered_map<const mg_connection *, std::shared_ptr<Session>> sessions;

	static State &From(const mg_connection *connection)
	{
		return *static_cast<State *>(mg_get_user_context_data(connection));
	}

	std::shared_ptr<Session> Find(const mg_connection *connection)
	{
		std::lock_guard<std::mutex> lock(sessionsMutex);
		auto found = sessions.find(connection);
		return found == sessions.end() ? nullptr : found->second;
	}

	std::shared_ptr<Session> Add(mg_connection *connection)
	{
		std::lock_guard<std::mutex> lock(sessionsMutex);
		if (stopping.load() || sessions.size() >= MaxSessions) return nullptr;
		auto session = std::make_shared<Session>(connection, dispatcher);
		sessions.emplace(connection, session);
		return session;
	}

	void CloseSession(const mg_connection *connection)
	{
		auto session = Find(connection);
		if (session == nullptr) return;
		// The CivetWeb connection remains valid throughout this callback. Join outside
		// sessionsMutex: Stop can cancel every outstanding STA request concurrently.
		session->CloseAndJoin();
		std::lock_guard<std::mutex> lock(sessionsMutex);
		sessions.erase(connection);
	}

	static int BeginRequest(mg_connection *connection) noexcept
	{
		try
		{
			auto &state = From(connection);
			if (!state.ready.load() || state.stopping.load())
				return SendHttp(connection, 503, HttpErrorReply(-32003, "Server is stopping").dump());
			const auto &request = *mg_get_request_info(connection);
			// The literal listening address guarantees the local address; reject non-IPv4
			// loopback peers as well. IPv6 and TLS are not enabled in this library build.
			if (request.is_ssl || request.server_port != state.port ||
				std::string_view(request.remote_addr).substr(0, 4) != "127.")
				return ProtocolError(connection, 403, "Only the selected IPv4 loopback endpoint is allowed");
			if (!AllowedOrigin(request, state.port))
				return ProtocolError(connection, 403, "Origin is not local to this endpoint");
			const Endpoint endpoint = IdentifyEndpoint(RequestPath(request));
			const bool websocket = endpoint == Endpoint::WebSocket;
			if (endpoint == Endpoint::Unknown) return ProtocolError(connection, 404, "Unknown endpoint");
			const char *method = EndpointMethod(endpoint);
			if (std::string_view(request.request_method) != method)
				return ProtocolError(connection, 405, "HTTP method is not supported by this endpoint", method);
			if (websocket)
			{
				if (std::string_view(request.http_version) != "1.1" ||
					!EqualAscii(Trim(Header(connection, "Upgrade")), "websocket") ||
					!HasToken(Header(connection, "Connection"), "Upgrade") || request.content_length > 0 ||
					(!Header(connection, "Content-Length").empty() && Header(connection, "Content-Length") != "0") ||
					!Header(connection, "Transfer-Encoding").empty())
					return ProtocolError(connection, 400, "Expected an HTTP/1.1 WebSocket upgrade");
			}
			else if (!Header(connection, "Upgrade").empty())
				return ProtocolError(connection, 400, "This endpoint does not support protocol upgrades");
			auto session = state.Add(connection);
			if (session == nullptr)
				return SendHttp(connection, 503, HttpErrorReply(-32000, "Server session capacity is exhausted").dump());
			// Allocate all owned threads before accepting the upgrade. They cannot send
			// until the ready callback, and handshake failure is cleaned by end_request.
			if (websocket) session->StartWorkers();
			return 0;
		}
		catch (const std::bad_alloc &)
		{
			return SendHttp(connection, 503, NativeFailure);
		}
		catch (const std::system_error &)
		{
			return SendHttp(connection, 503, NativeFailure);
		}
		catch (...)
		{
			return SendHttp(connection, 500, NativeFailure);
		}
	}

	static int HttpRequest(mg_connection *connection, void *data) noexcept
	{
		try { return static_cast<State *>(data)->RouteHttp(connection); }
		catch (...) { return SendHttp(connection, 500, NativeFailure); }
	}

	int RouteHttp(mg_connection *connection)
	{
		const auto &request = *mg_get_request_info(connection);
		std::string_view path = RequestPath(request);
		const Endpoint endpoint = IdentifyEndpoint(path);
		if (endpoint == Endpoint::Rpc)
		{
			auto session = Find(connection);
			if (session == nullptr || session->cancelled->load()) return SendHttp(connection, 503, NativeFailure);
			std::string_view contentType = Header(connection, "Content-Type");
			if (!EqualAscii(Trim(contentType.substr(0, contentType.find(';'))), "application/json"))
				return ProtocolError(connection, 415, "Content-Type must be application/json");
			if (request.content_length > static_cast<long long>(MaxMessageBytes))
				return ProtocolError(connection, 413, "Request exceeds the 1 MiB limit");
			std::string body;
			if (request.content_length > 0) body.reserve(static_cast<size_t>(request.content_length));
			std::array<char, 16 * 1024> buffer{};
			for (;;)
			{
				if (session->cancelled->load()) return 503;
				int bytes = mg_read(connection, buffer.data(), buffer.size());
				if (bytes < 0) return ProtocolError(connection, 400, "Unable to read the request body");
				if (bytes == 0) break;
				if (static_cast<size_t>(bytes) > MaxMessageBytes - body.size())
					return ProtocolError(connection, 413, "Request exceeds the 1 MiB limit");
				body.append(buffer.data(), static_cast<size_t>(bytes));
			}
			if (request.content_length >= 0 && body.size() != static_cast<size_t>(request.content_length))
				return ProtocolError(connection, 400, "Incomplete request body");
			// CivetWeb has no safe server-socket disconnect probe while this HTTP callback
			// waits on the STA. WebSocket receivers remain independent of command waiters.
			std::string reply = HandleEnvelope(body, dispatcher, session->cancelled);
			return session->cancelled->load() ? 503 : SendHttp(connection, 200, reply);
		}
		auto session = Find(connection);
		if (session == nullptr || session->cancelled->load())
			return SendHttp(connection, 503, HttpErrorReply(-32003, "Server is stopping").dump());
		std::string argument;
		bool force = false;
		const char *queryError = nullptr;
		if (!ParseHttpQuery(request.query_string, argument, force, queryError))
			return SendHttp(connection, 400, HttpErrorReply(-32602, queryError).dump());
		AutomationDispatcher::HttpCommand command;
		switch (endpoint)
		{
		case Endpoint::Version: command = AutomationDispatcher::HttpCommand::Version; argument.clear(); break;
		case Endpoint::List: command = AutomationDispatcher::HttpCommand::List; argument.clear(); break;
		case Endpoint::New: command = AutomationDispatcher::HttpCommand::New; break;
		case Endpoint::Activate:
			command = AutomationDispatcher::HttpCommand::Activate;
			argument = path.size() > 15 ? std::string(path.substr(15)) : "";
			break;
		case Endpoint::Close:
			command = AutomationDispatcher::HttpCommand::Close;
			argument = path.size() > 12 ? std::string(path.substr(12)) : "";
			break;
		default: return ProtocolError(connection, 404, "Unknown endpoint");
		}
		if ((endpoint == Endpoint::Activate || endpoint == Endpoint::Close) && !argument.empty())
		{
			// Decode exactly once with the library, retaining the returned byte count
			// so an encoded NUL cannot truncate a target into an otherwise valid GUID.
			const int length = mg_url_decode(argument.data(), static_cast<int>(argument.size()),
				argument.data(), static_cast<int>(argument.size() + 1), 0);
			if (length < 0)
				return SendHttp(connection, 400, HttpErrorReply(-32602, "Unable to decode target").dump());
			argument.resize(static_cast<size_t>(length));
			if (argument.find('\0') != std::string::npos)
				return SendHttp(connection, 400, HttpErrorReply(-32602, "Target must not contain a NUL character").dump());
		}
		Json envelope = dispatcher->HandleHttpRequest(command, argument, endpoint == Endpoint::Close && force, session->cancelled);
		if (session->cancelled->load())
			return SendHttp(connection, 503, HttpErrorReply(-32003, "Server is stopping").dump());
		int status = 200;
		Json reply;
		auto error = envelope.find("error");
		if (error != envelope.end())
		{
			status = HttpErrorStatus(error->value("code", -32000));
			reply = {{"error", *error}};
		}
		else
		{
			reply = envelope.at("result");
			if (endpoint == Endpoint::Version)
			{
				reply["Protocol-Version"] = "1.0";
				reply["webSocketDebuggerUrl"] = "ws://127.0.0.1:" + std::to_string(port) + "/devtools/application";
			}
		}
		std::string body = reply.dump();
		if (body.size() > MaxMessageBytes)
			return SendHttp(connection, 500, HttpErrorReply(-32000, "Reply exceeds the 1 MiB limit").dump());
		return SendHttp(connection, status, body);
	}

	static int WebSocketConnect(const mg_connection *connection, void *data) noexcept
	{
		try
		{
			auto session = static_cast<State *>(data)->Find(connection);
			return session != nullptr && !session->cancelled->load() ? 0 : 1;
		}
		catch (...) { return 1; }
	}

	static void WebSocketReady(mg_connection *connection, void *data) noexcept
	{
		try
		{
			auto session = static_cast<State *>(data)->Find(connection);
			if (session != nullptr) session->MakeReady();
		}
		catch (...) { /* The close callback owns cancellation and joining. */ }
	}

	static int WebSocketData(mg_connection *connection, int bits, char *data, size_t length, void *userData) noexcept
	{
		std::shared_ptr<Session> session;
		try
		{
			session = static_cast<State *>(userData)->Find(connection);
			return session == nullptr ? 0 : session->Receive(bits, data, length);
		}
		catch (...) { return session == nullptr ? 0 : session->Close(1011, "Native request handling failed"); }
	}

	static void WebSocketClose(const mg_connection *connection, void *data) noexcept
	{
		static_cast<State *>(data)->CloseSession(connection);
	}

	static void EndRequest(const mg_connection *connection, int) noexcept
	{
		From(connection).CloseSession(connection);
	}

	static void ConnectionClose(const mg_connection *connection) noexcept
	{
		From(connection).CloseSession(connection);
	}

	static int HttpError(mg_connection *connection, int status, const char *) noexcept
	{
		try { ProtocolError(connection, status, status == 413 ? "Request exceeds the 1 MiB limit" : "Invalid HTTP request"); }
		catch (...) { SendHttp(connection, status, NativeFailure); }
		return 0;
	}

	void Stop()
	{
		if (stopping.exchange(true)) return;
		ready.store(false);
		// Reentrant Office disconnect must release STA waiters before any network or
		// command worker is joined, including a waiter for the currently active COM call.
		if (dispatcher != nullptr) dispatcher->Stop();
		{
			std::lock_guard<std::mutex> lock(sessionsMutex);
			for (const auto &entry : sessions) entry.second->Cancel();
		}
		if (context != nullptr)
		{
			mg_stop(context);
			context = nullptr;
		}
		// mg_stop has joined all callbacks. Retain and join even an early failed upgrade
		// whose library path did not reach a close/end callback.
		for (const auto &entry : sessions) entry.second->CloseAndJoin();
		sessions.clear();
		if (libraryInitialized)
		{
			std::lock_guard<std::mutex> lock(LibraryMutex);
			mg_exit_library();
			libraryInitialized = false;
		}
	}
};

HttpServer::~HttpServer()
{
	Stop();
}

HRESULT HttpServer::Start(IDispatch *app, unsigned short port)
{
	if (app == nullptr || port == 0) return E_INVALIDARG;
	std::shared_ptr<State> state;
	auto cleanup = [&]
	{
		if (state == nullptr) return;
		bool ownsShutdown = false;
		{
			std::lock_guard<std::mutex> lock(m_mutex);
			if (m_state == state)
			{
				m_state.reset();
				m_stopping = ownsShutdown = true;
			}
		}
		state->Stop();
		if (ownsShutdown)
		{
			std::lock_guard<std::mutex> lock(m_mutex);
			m_stopping = false;
		}
	};
	try
	{
		state = std::make_shared<State>();
		state->dispatcher = std::make_shared<AutomationDispatcher>();
		state->port = port;
		state->sessions.reserve(MaxSessions);
		{
			std::lock_guard<std::mutex> lock(m_mutex);
			if (m_state != nullptr || m_stopping) return HRESULT_FROM_WIN32(ERROR_ALREADY_INITIALIZED);
			m_state = state;
		}
		HRESULT hr = state->dispatcher->Start(app);
		if (FAILED(hr)) { cleanup(); return hr; }
		if (state->stopping.load())
		{
			// Start may have been reentered by Office while acquiring the application or
			// creating its STA window. Do not leave a newly initialized dispatcher alive.
			state->dispatcher->Stop();
			cleanup();
			return HRESULT_FROM_WIN32(ERROR_OPERATION_ABORTED);
		}
		{
			std::lock_guard<std::mutex> lock(LibraryMutex);
			state->libraryInitialized = mg_init_library(MG_FEATURES_WEBSOCKET) != 0;
		}
		if (!state->libraryInitialized) { cleanup(); return E_FAIL; }
		const std::string listener = "127.0.0.1:" + std::to_string(port);
		const char *options[] = {
			"listening_ports", listener.c_str(),
			"num_threads", "32",
			"connection_queue", "32",
			"listen_backlog", "32",
			"max_request_size", "65536",
			"request_timeout_ms", "2000",
			"websocket_timeout_ms", "2000",
			"enable_websocket_ping_pong", "yes",
			"enable_keep_alive", "no",
			"tcp_nodelay", "1",
			"decode_url", "no",
			"decode_query_string", "no",
			"access_control_allow_origin", "",
			"access_control_allow_methods", "",
			"access_control_allow_headers", "",
			nullptr
		};
		mg_callbacks callbacks{};
		callbacks.begin_request = State::BeginRequest;
		callbacks.end_request = State::EndRequest;
		callbacks.connection_close = State::ConnectionClose;
		callbacks.http_error = State::HttpError;
		state->context = mg_start(&callbacks, state.get(), options);
		if (state->context == nullptr) { cleanup(); return E_FAIL; }
		mg_server_port bound{};
		if (mg_get_server_ports(state->context, 1, &bound) != 1 || bound.port != port ||
			bound.protocol != 1 || bound.is_ssl)
		{
			cleanup();
			return E_FAIL;
		}
		mg_set_request_handler(state->context, "/", State::HttpRequest, state.get());
		mg_set_websocket_handler(state->context, "/devtools/application", State::WebSocketConnect,
			State::WebSocketReady, State::WebSocketData, State::WebSocketClose, state.get());
		// The pinned CivetWeb build rejects any frame payload above 1 MiB before it
		// allocates the frame buffer; Session::Receive also caps aggregate fragments.
		state->ready.store(true);
		return S_OK;
	}
	catch (const std::bad_alloc &)
	{
		cleanup();
		return E_OUTOFMEMORY;
	}
	catch (const std::system_error &error)
	{
		cleanup();
		return HRESULT_FROM_WIN32(error.code().value());
	}
	catch (...)
	{
		cleanup();
		return E_FAIL;
	}
}

void HttpServer::Stop()
{
	std::shared_ptr<State> state;
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		if (m_stopping || m_state == nullptr) return;
		m_stopping = true;
		state = std::move(m_state);
	}
	// Remove the published state for idempotent disconnect, but do not permit a
	// reentrant Start to race the old listener while dispatcher->Stop releases COM.
	state->Stop();
	{
		std::lock_guard<std::mutex> lock(m_mutex);
		m_stopping = false;
	}
}
