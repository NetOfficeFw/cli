#include "pch.h"
#include "AutomationDispatcher.h"

#include <cctype>
#include <string>

namespace
{
	constexpr double CoordinateLimit = 10000.0;
}

AutomationDispatcher::Status AutomationDispatcher::AddLine(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<double> beginX, beginY, endX, endY;
	Status status = ReadNumber(context.parameters, "beginX", beginX, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "beginY", beginY, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "endX", endX, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "endY", endY, -CoordinateLimit, CoordinateLimit, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> shapes, shape;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	status = CallObject(shapes, L"AddLine", { ATL::CComVariant(static_cast<float>(*beginX)),
		ATL::CComVariant(static_cast<float>(*beginY)), ATL::CComVariant(static_cast<float>(*endX)),
		ATL::CComVariant(static_cast<float>(*endY)) }, shape);
	if (!status.ok())
		return status;
	long shapeId = 0;
	std::string name;
	status = GetInteger(shape, L"Id", shapeId);
	if (status.ok()) status = GetString(shape, L"Name", name);
	if (!status.ok())
		return status;
	result["shapeId"] = shapeId;
	result["name"] = name;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AddConnector(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> connectorType;
	std::optional<double> beginX, beginY, endX, endY;
	Status status = ReadChoice(context.parameters, "connectorType",
		{ { "straight", 1 }, { "elbow", 2 }, { "curve", 3 } }, connectorType, true);
	if (status.ok()) status = ReadNumber(context.parameters, "beginX", beginX, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "beginY", beginY, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "endX", endX, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "endY", endY, -CoordinateLimit, CoordinateLimit, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> shapes, shape;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	status = CallObject(shapes, L"AddConnector", { ATL::CComVariant(*connectorType),
		ATL::CComVariant(static_cast<float>(*beginX)), ATL::CComVariant(static_cast<float>(*beginY)),
		ATL::CComVariant(static_cast<float>(*endX)), ATL::CComVariant(static_cast<float>(*endY)) }, shape);
	if (!status.ok())
		return status;
	long shapeId = 0;
	std::string name;
	status = GetInteger(shape, L"Id", shapeId);
	if (status.ok()) status = GetString(shape, L"Name", name);
	if (!status.ok())
		return status;
	result["shapeId"] = shapeId;
	result["name"] = name;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ConnectConnector(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<long> beginShapeId, endShapeId, beginSite, endSite;
	Status status = ReadInteger(context.parameters, "beginShapeId", beginShapeId, 1, 2147483647L);
	if (status.ok()) status = ReadInteger(context.parameters, "endShapeId", endShapeId, 1, 2147483647L);
	if (status.ok()) status = ReadInteger(context.parameters, "beginSite", beginSite, 1, 2147483647L);
	if (status.ok()) status = ReadInteger(context.parameters, "endSite", endSite, 1, 2147483647L);
	if (!status.ok())
		return status;
	if (!beginShapeId && !endShapeId)
		return InvalidParameter("at least one of beginShapeId or endShapeId is required");
	if (beginSite && !beginShapeId)
		return InvalidParameter("beginSite requires beginShapeId");
	if (endSite && !endShapeId)
		return InvalidParameter("endSite requires endShapeId");
	if ((beginShapeId && *beginShapeId == context.shapeId) || (endShapeId && *endShapeId == context.shapeId))
		return InvalidParameter("A connector cannot connect to itself");

	long isConnector = 0;
	status = GetInteger(context.shape, L"Connector", isConnector);
	if (!status.ok())
		return status;
	if (isConnector == 0)
		return InvalidParameter("The shape is not a connector");

	// Resolve targets and validate sites before mutating.
	auto resolve = [&](const std::optional<long> &id, std::optional<long> &site, const char *siteName,
		ATL::CComPtr<IDispatch> &target) -> Status
	{
		if (!id)
			return {};
		Status resolved = GetShapeById(context.container, *id, target);
		if (!resolved.ok())
			return resolved;
		long siteCount = 0;
		resolved = GetInteger(target, L"ConnectionSiteCount", siteCount);
		if (!resolved.ok())
			return resolved;
		if (siteCount < 1)
			return InvalidParameter("Shape " + std::to_string(*id) + " has no connection sites");
		if (!site)
			site = 1;
		else if (*site > siteCount)
			return InvalidParameter(std::string(siteName) + " must be between 1 and " + std::to_string(siteCount));
		return {};
	};
	ATL::CComPtr<IDispatch> beginShape, endShape;
	status = resolve(beginShapeId, beginSite, "beginSite", beginShape);
	if (status.ok()) status = resolve(endShapeId, endSite, "endSite", endShape);
	if (!status.ok())
		return status;
	if (IsStopping())
		return StoppedStatus();

	ATL::CComPtr<IDispatch> format;
	status = GetObject(context.shape, L"ConnectorFormat", format);
	if (!status.ok())
		return status;
	if (beginShape)
	{
		status = CallMethod(format, L"BeginConnect", { ATL::CComVariant(beginShape.p), ATL::CComVariant(*beginSite) });
		if (!status.ok())
			return status;
	}
	if (endShape)
	{
		status = CallMethod(format, L"EndConnect", { ATL::CComVariant(endShape.p), ATL::CComVariant(*endSite) });
		if (!status.ok())
			return status;
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AddPicture(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> path;
	std::optional<double> left, top, width, height;
	Status status = ReadPath(context.parameters, "path", path, true);
	if (status.ok()) status = ReadNumber(context.parameters, "left", left, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "top", top, -CoordinateLimit, CoordinateLimit, true);
	if (status.ok()) status = ReadNumber(context.parameters, "width", width, -1.0, CoordinateLimit);
	if (status.ok()) status = ReadNumber(context.parameters, "height", height, -1.0, CoordinateLimit);
	if (!status.ok())
		return status;
	if (width.has_value() != height.has_value())
		return InvalidParameter("width and height must be given together");
	auto validSize = [](double value) { return value == -1.0 || value >= 0.01; };
	if (width && (!validSize(*width) || !validSize(*height)))
		return InvalidParameter("width and height must be greater than 0, or -1 for natural size");

	const size_t dot = path->find_last_of('.');
	std::string extension = dot == std::string::npos ? std::string() : path->substr(dot);
	for (char &character : extension)
		character = static_cast<char>(std::tolower(static_cast<unsigned char>(character)));
	static const char *const allowed[] = { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".emf",
		".wmf", ".tif", ".tiff", ".ico" };
	bool supported = false;
	for (const char *candidate : allowed)
		if (extension == candidate)
			supported = true;
	if (!supported)
		return InvalidParameter("path must end in .png, .jpg, .jpeg, .gif, .bmp, .svg, .emf, .wmf, .tif, .tiff, or .ico");
	ATL::CComVariant fileName;
	HRESULT hr = Utf8ToVariant(*path, fileName);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding picture path");
	if (GetFileAttributesW(fileName.bstrVal) == INVALID_FILE_ATTRIBUTES)
		return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Picture file", *path);

	ATL::CComPtr<IDispatch> shapes, shape;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	const float pictureWidth = width ? static_cast<float>(*width) : -1.0f;
	const float pictureHeight = height ? static_cast<float>(*height) : -1.0f;
	status = CallObject(shapes, L"AddPicture", { fileName, ATL::CComVariant(0L), ATL::CComVariant(-1L),
		ATL::CComVariant(static_cast<float>(*left)), ATL::CComVariant(static_cast<float>(*top)),
		ATL::CComVariant(pictureWidth), ATL::CComVariant(pictureHeight) }, shape);
	if (!status.ok())
		return status;
	long shapeId = 0;
	std::string name;
	status = GetInteger(shape, L"Id", shapeId);
	if (status.ok()) status = GetString(shape, L"Name", name);
	if (!status.ok())
		return status;
	result["shapeId"] = shapeId;
	result["name"] = name;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GroupShapes(const CommandContext &context,
	nlohmann::json &result)
{
	std::vector<long> shapeIds;
	Status status = ReadShapeIds(context.parameters, "shapeIds", shapeIds, 2);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> range, group;
	status = GetShapeRange(context.container, shapeIds, range);
	if (!status.ok())
		return status;
	status = CallObject(range, L"Group", {}, group);
	if (!status.ok())
		return status;
	long shapeId = 0;
	std::string name;
	status = GetInteger(group, L"Id", shapeId);
	if (status.ok()) status = GetString(group, L"Name", name);
	if (!status.ok())
		return status;
	result["shapeId"] = shapeId;
	result["name"] = name;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::UngroupShape(const CommandContext &context,
	nlohmann::json &result)
{
	long type = 0;
	Status status = GetInteger(context.shape, L"Type", type);
	if (!status.ok())
		return status;
	if (type != 6) // msoGroup
		return InvalidParameter("The shape is not a group");
	ATL::CComPtr<IDispatch> range;
	status = CallObject(context.shape, L"Ungroup", {}, range);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(range, L"Count", count);
	if (!status.ok())
		return status;
	nlohmann::json ids = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		if (IsStopping())
			return StoppedStatus();
		ATL::CComPtr<IDispatch> item;
		status = GetItem(range, index, item);
		if (!status.ok())
			return status;
		long id = 0;
		status = GetInteger(item, L"Id", id);
		if (!status.ok())
			return status;
		ids.push_back(id);
	}
	result["shapeIds"] = std::move(ids);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AlignShapes(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<long> alignment;
	std::optional<bool> relative;
	Status status = ReadChoice(context.parameters, "alignment", { { "left", 0 }, { "center", 1 },
		{ "right", 2 }, { "top", 3 }, { "middle", 4 }, { "bottom", 5 } }, alignment, true);
	if (status.ok()) status = ReadBoolean(context.parameters, "relativeToSlide", relative);
	if (!status.ok())
		return status;
	const bool relativeToSlide = relative.value_or(false);
	std::vector<long> shapeIds;
	status = ReadShapeIds(context.parameters, "shapeIds", shapeIds, relativeToSlide ? 1 : 2);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> range;
	status = GetShapeRange(context.container, shapeIds, range);
	if (!status.ok())
		return status;
	return CallMethod(range, L"Align", { ATL::CComVariant(*alignment),
		ATL::CComVariant(relativeToSlide ? -1L : 0L) });
}

AutomationDispatcher::Status AutomationDispatcher::DistributeShapes(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<long> direction;
	std::optional<bool> relative;
	Status status = ReadChoice(context.parameters, "direction", { { "horizontal", 0 }, { "vertical", 1 } },
		direction, true);
	if (status.ok()) status = ReadBoolean(context.parameters, "relativeToSlide", relative);
	if (!status.ok())
		return status;
	const bool relativeToSlide = relative.value_or(false);
	std::vector<long> shapeIds;
	status = ReadShapeIds(context.parameters, "shapeIds", shapeIds, relativeToSlide ? 1 : 3);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> range;
	status = GetShapeRange(context.container, shapeIds, range);
	if (!status.ok())
		return status;
	return CallMethod(range, L"Distribute", { ATL::CComVariant(*direction),
		ATL::CComVariant(relativeToSlide ? -1L : 0L) });
}
