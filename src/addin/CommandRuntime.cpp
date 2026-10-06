#include "pch.h"
#include "AutomationDispatcher.h"

#include <algorithm>
#include <atlsafe.h>
#include <climits>
#include <cmath>
#include <cstdio>
#include <set>

const std::vector<AutomationDispatcher::CommandInfo> &AutomationDispatcher::Commands()
{
	static const std::vector<CommandInfo> commands = {
#define NETOFFICE_COMMAND_INFO(name, method, verb, scope, suffix) \
		{ HttpCommand::name, method, verb, CommandScope::scope, suffix },
		NETOFFICE_POWERPOINT_COMMANDS(NETOFFICE_COMMAND_INFO)
#undef NETOFFICE_COMMAND_INFO
	};
	return commands;
}

const AutomationDispatcher::CommandInfo *AutomationDispatcher::FindCommand(HttpCommand command)
{
	for (const auto &info : Commands())
		if (info.command == command)
			return &info;
	return nullptr;
}

const AutomationDispatcher::CommandInfo *AutomationDispatcher::FindCommand(const std::string &method)
{
	for (const auto &info : Commands())
		if (method == info.method)
			return &info;
	return nullptr;
}

AutomationDispatcher::Status AutomationDispatcher::GetCustomLayout(IDispatch *document, long index,
	ATL::CComPtr<IDispatch> &layout)
{
	ATL::CComPtr<IDispatch> master, layouts;
	Status status = GetObject(document, L"SlideMaster", master);
	if (!status.ok())
		return status;
	status = GetObject(master, L"CustomLayouts", layouts);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(layouts, L"Count", count);
	if (!status.ok())
		return status;
	if (index < 1 || index > count)
		return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Custom layout",
			"customLayout exceeds the slide master's layout count");
	return GetItem(layouts, index, layout);
}

AutomationDispatcher::Status AutomationDispatcher::FindCustomLayout(IDispatch *document,
	const std::string &name, ATL::CComPtr<IDispatch> &layout, long &index)
{
	ATL::CComPtr<IDispatch> master, layouts;
	Status status = GetObject(document, L"SlideMaster", master);
	if (!status.ok())
		return status;
	status = GetObject(master, L"CustomLayouts", layouts);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(layouts, L"Count", count);
	if (!status.ok())
		return status;
	for (long position = 1; position <= count; ++position)
	{
		ATL::CComPtr<IDispatch> candidate;
		status = GetItem(layouts, position, candidate);
		if (!status.ok())
			return status;
		std::string candidateName;
		status = GetString(candidate, L"Name", candidateName);
		if (!status.ok())
			return status;
		if (candidateName == name)
		{
			layout = candidate;
			index = position;
			return {};
		}
	}
	return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Custom layout",
		"No custom layout has that name; use layout list");
}

AutomationDispatcher::Status AutomationDispatcher::ResolveContainer(IDispatch *document, long slideId,
	bool master, long customLayout, ATL::CComPtr<IDispatch> &container, long &slideIndex)
{
	slideIndex = 0;
	if (customLayout > 0)
		return GetCustomLayout(document, customLayout, container);
	if (master)
		return GetObject(document, L"SlideMaster", container);
	return GetSlideById(document, slideId, container, slideIndex);
}

AutomationDispatcher::Status AutomationDispatcher::RunCommand(const CommandInfo &info, IDispatch *document,
	const std::string &id, long slideId, long shapeId, const nlohmann::json &parameters, nlohmann::json &result)
{
	const bool master = parameters.contains("master") && parameters["master"].is_boolean() &&
		parameters["master"].get<bool>();
	std::optional<long> customLayout;
	const bool containerScope = info.scope == CommandScope::Container || info.scope == CommandScope::Shape;
	if (containerScope)
	{
		Status status = ReadInteger(parameters, "customLayout", customLayout, 1, INT_MAX);
		if (!status.ok())
			return status;
	}
	if ((master || parameters.contains("customLayout")) && !containerScope)
		return InvalidParameter("master and customLayout are accepted only by slide, master, or layout commands");
	if (master && customLayout)
		return InvalidParameter("master and customLayout are mutually exclusive");
	CommandContext context{ document, id, parameters };
	result = nlohmann::json::object();
	if (info.scope != CommandScope::Application)
		result["id"] = id;
	ATL::CComPtr<IDispatch> container, shape;
	if (info.scope == CommandScope::Slide || containerScope)
	{
		Status status = ResolveContainer(document, slideId, master, customLayout.value_or(0),
			container, context.slideIndex);
		if (!status.ok())
			return status;
		context.container = container;
		context.master = master;
		context.customLayout = customLayout.value_or(0);
		if (customLayout)
			result["customLayout"] = *customLayout;
		else if (master)
			result["master"] = true;
		else
		{
			context.slideId = slideId;
			result["slideId"] = slideId;
		}
	}
	if (info.scope == CommandScope::Shape)
	{
		Status status = GetShapeById(container, shapeId, shape);
		if (!status.ok())
			return status;
		context.shape = shape;
		context.shapeId = shapeId;
		result["shapeId"] = shapeId;
	}

	Status status;
	switch (info.command)
	{
#define NETOFFICE_COMMAND_CASE(name, method, verb, scope, suffix) \
	case HttpCommand::name: status = name(context, result); break;
		NETOFFICE_POWERPOINT_COMMANDS(NETOFFICE_COMMAND_CASE)
#undef NETOFFICE_COMMAND_CASE
	default:
		return ErrorStatus(-32601, E_NOTIMPL, "PowerPoint command");
	}
	if (!status.ok())
		return status;
	if (std::string_view(info.verb) != "GET" && result.is_object())
	{
		for (auto item = parameters.begin(); item != parameters.end(); ++item)
		{
			const std::string &key = item.key();
			if (key == "targetId" || key == "slideId" || key == "shapeId" || key == "master" ||
				result.contains(key))
				continue;
			result[key] = item.value();
		}
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::InvalidParameter(const std::string &message)
{
	return ErrorStatus(-32602, E_INVALIDARG, "PowerPoint command parameters", message);
}

AutomationDispatcher::Status AutomationDispatcher::ReadNumber(const nlohmann::json &parameters,
	const char *name, std::optional<double> &value, double minimum, double maximum, bool required)
{
	value.reset();
	if (!parameters.contains(name))
		return required ? InvalidParameter(std::string(name) + " is required") : Status{};
	const auto &member = parameters[name];
	if (!member.is_number() || !std::isfinite(member.get<double>()) ||
		member.get<double>() < minimum || member.get<double>() > maximum)
	{
		char message[160];
		std::snprintf(message, sizeof(message), "%s must be a finite number from %g through %g",
			name, minimum, maximum);
		return InvalidParameter(message);
	}
	value = member.get<double>();
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadInteger(const nlohmann::json &parameters,
	const char *name, std::optional<long> &value, long minimum, long maximum, bool required)
{
	value.reset();
	if (!parameters.contains(name))
		return required ? InvalidParameter(std::string(name) + " is required") : Status{};
	const auto &member = parameters[name];
	bool valid = false;
	long long number = 0;
	if (member.is_number_unsigned())
	{
		const auto parsed = member.get<uint64_t>();
		valid = parsed <= static_cast<uint64_t>(LLONG_MAX);
		number = static_cast<long long>(parsed);
	}
	else if (member.is_number_integer())
	{
		valid = true;
		number = member.get<int64_t>();
	}
	if (!valid || number < minimum || number > maximum)
		return InvalidParameter(std::string(name) + " must be an integer from " +
			std::to_string(minimum) + " through " + std::to_string(maximum));
	value = static_cast<long>(number);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadBoolean(const nlohmann::json &parameters,
	const char *name, std::optional<bool> &value, bool required)
{
	value.reset();
	if (!parameters.contains(name))
		return required ? InvalidParameter(std::string(name) + " is required") : Status{};
	if (!parameters[name].is_boolean())
		return InvalidParameter(std::string(name) + " must be a boolean");
	value = parameters[name].get<bool>();
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadString(const nlohmann::json &parameters,
	const char *name, std::optional<std::string> &value, bool required, bool allowEmpty)
{
	value.reset();
	if (!parameters.contains(name))
		return required ? InvalidParameter(std::string(name) + " is required") : Status{};
	const auto &member = parameters[name];
	if (!member.is_string() || (!allowEmpty && member.get_ref<const std::string &>().empty()))
		return InvalidParameter(std::string(name) + (allowEmpty ? " must be a string" : " must be a non-empty string"));
	if (member.get_ref<const std::string &>().find('\0') != std::string::npos)
		return InvalidParameter(std::string(name) + " must not contain a NUL character");
	value = member.get<std::string>();
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadPath(const nlohmann::json &parameters,
	const char *name, std::optional<std::string> &value, bool required)
{
	Status status = ReadString(parameters, name, value, required, false);
	if (!status.ok() || !value)
		return status;
	const std::string &path = *value;
	const bool drive = path.size() >= 3 && ((path[0] >= 'A' && path[0] <= 'Z') || (path[0] >= 'a' && path[0] <= 'z')) &&
		path[1] == ':' && (path[2] == '\\' || path[2] == '/');
	const bool unc = path.size() >= 3 && (path[0] == '\\' || path[0] == '/') && (path[1] == '\\' || path[1] == '/');
	if (!drive && !unc)
	{
		value.reset();
		return InvalidParameter(std::string(name) + " must be an absolute local or UNC path");
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadColor(const nlohmann::json &parameters,
	const char *name, std::optional<long> &value, bool required)
{
	value.reset();
	if (!parameters.contains(name))
		return required ? InvalidParameter(std::string(name) + " is required") : Status{};
	const auto &member = parameters[name];
	auto hex = [](char digit) -> int
	{
		if (digit >= '0' && digit <= '9') return digit - '0';
		if (digit >= 'a' && digit <= 'f') return digit - 'a' + 10;
		if (digit >= 'A' && digit <= 'F') return digit - 'A' + 10;
		return -1;
	};
	if (member.is_string())
	{
		const auto &text = member.get_ref<const std::string &>();
		if (text.size() == 7 && text[0] == '#')
		{
			int channels[3] = {};
			bool valid = true;
			for (int channel = 0; channel < 3 && valid; ++channel)
			{
				const int high = hex(text[1 + channel * 2]), low = hex(text[2 + channel * 2]);
				valid = high >= 0 && low >= 0;
				channels[channel] = high * 16 + low;
			}
			if (valid)
			{
				value = channels[0] | (channels[1] << 8) | (channels[2] << 16);
				return {};
			}
		}
	}
	return InvalidParameter(std::string(name) + " must be a #RRGGBB color");
}

AutomationDispatcher::Status AutomationDispatcher::ReadChoice(const nlohmann::json &parameters,
	const char *name, std::initializer_list<std::pair<const char *, long>> choices,
	std::optional<long> &value, bool required)
{
	value.reset();
	if (!parameters.contains(name))
		return required ? InvalidParameter(std::string(name) + " is required") : Status{};
	std::string allowed;
	if (parameters[name].is_string())
	{
		const auto &text = parameters[name].get_ref<const std::string &>();
		for (const auto &choice : choices)
		{
			if (text == choice.first)
			{
				value = choice.second;
				return {};
			}
		}
	}
	for (const auto &choice : choices)
	{
		if (!allowed.empty())
			allowed += ", ";
		allowed += choice.first;
	}
	return InvalidParameter(std::string(name) + " must be one of: " + allowed);
}

AutomationDispatcher::Status AutomationDispatcher::ReadShapeIds(const nlohmann::json &parameters,
	const char *name, std::vector<long> &value, size_t minimumCount)
{
	value.clear();
	const std::string message = std::string(name) + " must be an array of at least " +
		std::to_string(minimumCount) + " distinct positive int32 shape IDs";
	if (!parameters.contains(name) || !parameters[name].is_array() || parameters[name].size() < minimumCount)
		return InvalidParameter(message);
	std::set<long> seen;
	for (const auto &member : parameters[name])
	{
		long long number = 0;
		if (member.is_number_unsigned() && member.get<uint64_t>() <= INT_MAX)
			number = static_cast<long long>(member.get<uint64_t>());
		else if (member.is_number_integer())
			number = member.get<int64_t>();
		else
			return InvalidParameter(message);
		if (number < 1 || number > INT_MAX || !seen.insert(static_cast<long>(number)).second)
			return InvalidParameter(message);
		value.push_back(static_cast<long>(number));
	}
	return {};
}

std::string AutomationDispatcher::ColorToHex(long rgb)
{
	char text[8];
	std::snprintf(text, sizeof(text), "#%02X%02X%02X", static_cast<unsigned>(rgb & 0xFF),
		static_cast<unsigned>((rgb >> 8) & 0xFF), static_cast<unsigned>((rgb >> 16) & 0xFF));
	return text;
}

ATL::CComVariant AutomationDispatcher::MissingArgument()
{
	ATL::CComVariant missing;
	missing.vt = VT_ERROR;
	missing.scode = DISP_E_PARAMNOTFOUND;
	return missing;
}

AutomationDispatcher::Status AutomationDispatcher::SetProperty(IDispatch *object, const wchar_t *name,
	const ATL::CComVariant &value)
{
	ATL::CComVariant argument(value);
	return Invoke(object, name, DISPATCH_PROPERTYPUT, &argument, 1, nullptr);
}

AutomationDispatcher::Status AutomationDispatcher::CallMethod(IDispatch *object, const wchar_t *name,
	std::vector<ATL::CComVariant> arguments, ATL::CComVariant *result)
{
	// IDispatch::Invoke takes positional arguments in reverse order.
	std::reverse(arguments.begin(), arguments.end());
	return Invoke(object, name, DISPATCH_METHOD | (result != nullptr ? DISPATCH_PROPERTYGET : 0),
		arguments.empty() ? nullptr : arguments.data(), static_cast<UINT>(arguments.size()), result);
}

AutomationDispatcher::Status AutomationDispatcher::CallObject(IDispatch *object, const wchar_t *name,
	std::vector<ATL::CComVariant> arguments, ATL::CComPtr<IDispatch> &result)
{
	ATL::CComVariant value;
	Status status = CallMethod(object, name, std::move(arguments), &value);
	if (!status.ok())
		return status;
	return ReadObject(value, name, result, -32000);
}

AutomationDispatcher::Status AutomationDispatcher::GetItem(IDispatch *collection, long index,
	ATL::CComPtr<IDispatch> &item)
{
	return CallObject(collection, L"Item", { ATL::CComVariant(index) }, item);
}

AutomationDispatcher::Status AutomationDispatcher::GetShapeRange(IDispatch *container,
	const std::vector<long> &shapeIds, ATL::CComPtr<IDispatch> &range)
{
	ATL::CComPtr<IDispatch> shapes;
	Status status = GetObject(container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(shapes, L"Count", count);
	if (!status.ok())
		return status;
	std::vector<long> indices(shapeIds.size(), 0);
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComPtr<IDispatch> shape;
		status = GetItem(shapes, index, shape);
		if (!status.ok())
			return status;
		long id = 0;
		status = GetInteger(shape, L"Id", id);
		if (!status.ok())
			return status;
		for (size_t position = 0; position < shapeIds.size(); ++position)
			if (shapeIds[position] == id)
				indices[position] = index;
	}
	for (size_t position = 0; position < shapeIds.size(); ++position)
		if (indices[position] == 0)
			return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Shape ID",
				"Shape ID " + std::to_string(shapeIds[position]) + " was not found on the requested slide");
	ATL::CComSafeArray<VARIANT> array(static_cast<ULONG>(indices.size()));
	for (size_t position = 0; position < indices.size(); ++position)
	{
		HRESULT hr = array.SetAt(static_cast<LONG>(position), ATL::CComVariant(indices[position]));
		if (FAILED(hr))
			return ErrorStatus(-32000, hr, "Building shape range");
	}
	ATL::CComVariant argument(array.m_psa);
	return CallObject(shapes, L"Range", { argument }, range);
}
