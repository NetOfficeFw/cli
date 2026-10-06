#include "pch.h"
#include "AutomationDispatcher.h"

#include <climits>
#include <string>

namespace
{
	ATL::CComVariant TriState(bool value)
	{
		return ATL::CComVariant(value ? -1L : 0L);
	}
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeFill(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	const nlohmann::json &parameters = context.parameters;
	std::optional<bool> visible;
	std::optional<long> color;
	std::optional<long> color2;
	std::optional<long> gradientStyle;
	std::optional<double> transparency;
	std::optional<std::string> picture;
	Status status = ReadBoolean(parameters, "visible", visible);
	if (status.ok()) status = ReadColor(parameters, "color", color);
	if (status.ok()) status = ReadColor(parameters, "color2", color2);
	if (status.ok())
		status = ReadChoice(parameters, "gradientStyle", {
			{ "horizontal", 1 }, { "vertical", 2 }, { "diagonal-up", 3 },
			{ "diagonal-down", 4 }, { "from-corner", 5 }, { "from-center", 7 } }, gradientStyle);
	if (status.ok()) status = ReadNumber(parameters, "transparency", transparency, 0.0, 1.0);
	if (status.ok()) status = ReadPath(parameters, "picture", picture);
	if (!status.ok())
		return status;
	if (!visible && !color && !color2 && !gradientStyle && !transparency && !picture)
		return InvalidParameter("at least one of visible, color, color2, gradientStyle, transparency, or picture is required");
	if (color2 && !color)
		return InvalidParameter("color2 requires color");
	if (gradientStyle && !color2)
		return InvalidParameter("gradientStyle requires color2");
	if (picture && color)
		return InvalidParameter("picture cannot be combined with color");
	if (visible && !*visible && (color || color2 || transparency || picture))
		return InvalidParameter("visible false cannot be combined with other fill members");
	ATL::CComVariant pictureName;
	if (picture)
	{
		HRESULT hr = Utf8ToVariant(*picture, pictureName);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding picture path");
	}

	ATL::CComPtr<IDispatch> fill;
	status = GetObject(context.shape, L"Fill", fill);
	if (!status.ok())
		return status;
	if (visible && !*visible)
		return SetProperty(fill, L"Visible", TriState(false));
	if (visible || color || picture)
	{
		status = SetProperty(fill, L"Visible", TriState(true));
		if (!status.ok())
			return status;
	}
	if (picture)
	{
		status = CallMethod(fill, L"UserPicture", { pictureName });
		if (!status.ok())
			return status;
	}
	if (color)
	{
		if (color2)
			status = CallMethod(fill, L"TwoColorGradient", { ATL::CComVariant(gradientStyle.value_or(1L)), ATL::CComVariant(1L) });
		else
			status = CallMethod(fill, L"Solid", {});
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> foreColor;
		status = GetObject(fill, L"ForeColor", foreColor);
		if (status.ok()) status = SetProperty(foreColor, L"RGB", ATL::CComVariant(*color));
		if (!status.ok())
			return status;
		if (color2)
		{
			ATL::CComPtr<IDispatch> backColor;
			status = GetObject(fill, L"BackColor", backColor);
			if (status.ok()) status = SetProperty(backColor, L"RGB", ATL::CComVariant(*color2));
			if (!status.ok())
				return status;
		}
	}
	if (transparency)
		return SetProperty(fill, L"Transparency", ATL::CComVariant(*transparency));
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeLine(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	const nlohmann::json &parameters = context.parameters;
	std::optional<bool> visible;
	std::optional<long> color;
	std::optional<double> weight;
	std::optional<long> dash;
	std::optional<double> transparency;
	Status status = ReadBoolean(parameters, "visible", visible);
	if (status.ok()) status = ReadColor(parameters, "color", color);
	if (status.ok()) status = ReadNumber(parameters, "weight", weight, 0.0, 1584.0);
	if (status.ok())
		status = ReadChoice(parameters, "dash", {
			{ "solid", 1 }, { "square-dot", 2 }, { "round-dot", 3 }, { "dash", 4 },
			{ "dash-dot", 5 }, { "dash-dot-dot", 6 }, { "long-dash", 7 }, { "long-dash-dot", 8 } }, dash);
	if (status.ok()) status = ReadNumber(parameters, "transparency", transparency, 0.0, 1.0);
	if (!status.ok())
		return status;
	if (!visible && !color && !weight && !dash && !transparency)
		return InvalidParameter("at least one of visible, color, weight, dash, or transparency is required");
	if (visible && !*visible && (color || weight))
		return InvalidParameter("visible false cannot be combined with color or weight");

	ATL::CComPtr<IDispatch> line;
	status = GetObject(context.shape, L"Line", line);
	if (!status.ok())
		return status;
	if (visible || color || weight)
	{
		status = SetProperty(line, L"Visible", TriState(visible.value_or(true)));
		if (!status.ok())
			return status;
	}
	if (color)
	{
		ATL::CComPtr<IDispatch> foreColor;
		status = GetObject(line, L"ForeColor", foreColor);
		if (status.ok()) status = SetProperty(foreColor, L"RGB", ATL::CComVariant(*color));
		if (!status.ok())
			return status;
	}
	if (weight)
	{
		status = SetProperty(line, L"Weight", ATL::CComVariant(*weight));
		if (!status.ok())
			return status;
	}
	if (dash)
	{
		status = SetProperty(line, L"DashStyle", ATL::CComVariant(*dash));
		if (!status.ok())
			return status;
	}
	if (transparency)
		return SetProperty(line, L"Transparency", ATL::CComVariant(*transparency));
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeShadow(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	const nlohmann::json &parameters = context.parameters;
	std::optional<bool> visible;
	std::optional<long> color;
	std::optional<double> blur;
	std::optional<double> offsetX;
	std::optional<double> offsetY;
	std::optional<double> transparency;
	std::optional<double> size;
	Status status = ReadBoolean(parameters, "visible", visible);
	if (status.ok()) status = ReadColor(parameters, "color", color);
	if (status.ok()) status = ReadNumber(parameters, "blur", blur, 0.0, 100.0);
	if (status.ok()) status = ReadNumber(parameters, "offsetX", offsetX, -200.0, 200.0);
	if (status.ok()) status = ReadNumber(parameters, "offsetY", offsetY, -200.0, 200.0);
	if (status.ok()) status = ReadNumber(parameters, "transparency", transparency, 0.0, 1.0);
	if (status.ok()) status = ReadNumber(parameters, "size", size, 1.0, 200.0);
	if (!status.ok())
		return status;
	const bool styled = color || blur || offsetX || offsetY || transparency || size;
	if (!visible && !styled)
		return InvalidParameter("at least one of visible, color, blur, offsetX, offsetY, transparency, or size is required");
	if (visible && !*visible && styled)
		return InvalidParameter("visible false cannot be combined with other shadow members");

	ATL::CComPtr<IDispatch> shadow;
	status = GetObject(context.shape, L"Shadow", shadow);
	if (!status.ok())
		return status;
	status = SetProperty(shadow, L"Visible", TriState(visible.value_or(true)));
	if (!status.ok() || !visible.value_or(true))
		return status;
	if (color)
	{
		ATL::CComPtr<IDispatch> foreColor;
		status = GetObject(shadow, L"ForeColor", foreColor);
		if (status.ok()) status = SetProperty(foreColor, L"RGB", ATL::CComVariant(*color));
		if (!status.ok())
			return status;
	}
	const std::pair<const wchar_t *, const std::optional<double> *> members[] = {
		{ L"Blur", &blur }, { L"OffsetX", &offsetX }, { L"OffsetY", &offsetY },
		{ L"Transparency", &transparency }, { L"Size", &size } };
	for (const auto &member : members)
	{
		if (!*member.second)
			continue;
		status = SetProperty(shadow, member.first, ATL::CComVariant(**member.second));
		if (!status.ok())
			return status;
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeGlow(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<double> radius;
	std::optional<long> color;
	std::optional<double> transparency;
	Status status = ReadNumber(context.parameters, "radius", radius, 0.0, 150.0, true);
	if (status.ok()) status = ReadColor(context.parameters, "color", color);
	if (status.ok()) status = ReadNumber(context.parameters, "transparency", transparency, 0.0, 1.0);
	if (!status.ok())
		return status;

	ATL::CComPtr<IDispatch> glow;
	status = GetObject(context.shape, L"Glow", glow);
	if (status.ok()) status = SetProperty(glow, L"Radius", ATL::CComVariant(*radius));
	if (!status.ok())
		return status;
	if (color)
	{
		ATL::CComPtr<IDispatch> glowColor;
		status = GetObject(glow, L"Color", glowColor);
		if (status.ok()) status = SetProperty(glowColor, L"RGB", ATL::CComVariant(*color));
		if (!status.ok())
			return status;
	}
	if (transparency)
		return SetProperty(glow, L"Transparency", ATL::CComVariant(*transparency));
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeSoftEdge(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<double> radius;
	Status status = ReadNumber(context.parameters, "radius", radius, 0.0, 100.0, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> softEdge;
	status = GetObject(context.shape, L"SoftEdge", softEdge);
	if (!status.ok())
		return status;
	return SetProperty(softEdge, L"Radius", ATL::CComVariant(*radius));
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeThreeD(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	const nlohmann::json &parameters = context.parameters;
	std::optional<long> bevel;
	std::optional<double> bevelDepth;
	std::optional<double> bevelInset;
	std::optional<double> depth;
	std::optional<long> material;
	Status status = ReadInteger(parameters, "bevel", bevel, 1, 13);
	if (status.ok()) status = ReadNumber(parameters, "bevelDepth", bevelDepth, 0.0, 1584.0);
	if (status.ok()) status = ReadNumber(parameters, "bevelInset", bevelInset, 0.0, 1584.0);
	if (status.ok()) status = ReadNumber(parameters, "depth", depth, 0.0, 1584.0);
	if (status.ok()) status = ReadInteger(parameters, "material", material, 1, 20);
	if (!status.ok())
		return status;
	if (!bevel && !bevelDepth && !bevelInset && !depth && !material)
		return InvalidParameter("at least one of bevel, bevelDepth, bevelInset, depth, or material is required");

	ATL::CComPtr<IDispatch> threeD;
	status = GetObject(context.shape, L"ThreeD", threeD);
	if (!status.ok())
		return status;
	if (bevel)
	{
		status = SetProperty(threeD, L"BevelTopType", ATL::CComVariant(*bevel));
		if (!status.ok())
			return status;
	}
	const std::pair<const wchar_t *, const std::optional<double> *> members[] = {
		{ L"BevelTopDepth", &bevelDepth }, { L"BevelTopInset", &bevelInset }, { L"Depth", &depth } };
	for (const auto &member : members)
	{
		if (!*member.second)
			continue;
		status = SetProperty(threeD, member.first, ATL::CComVariant(**member.second));
		if (!status.ok())
			return status;
	}
	if (material)
		return SetProperty(threeD, L"PresetMaterial", ATL::CComVariant(*material));
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeStyle(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<long> style;
	Status status = ReadInteger(context.parameters, "style", style, 1, 42, true);
	if (!status.ok())
		return status;
	// msoShapeStylePreset1 = 1 .. msoShapeStylePreset42 = 42 (10001+ are line presets).
	return SetProperty(context.shape, L"ShapeStyle", ATL::CComVariant(*style));
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeAdjustment(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> index;
	std::optional<double> value;
	Status status = ReadInteger(context.parameters, "index", index, 1, LONG_MAX, true);
	if (status.ok()) status = ReadNumber(context.parameters, "value", value, -1000.0, 1000.0, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> adjustments;
	status = GetObject(context.shape, L"Adjustments", adjustments);
	long count = 0;
	if (status.ok()) status = GetInteger(adjustments, L"Count", count);
	if (!status.ok())
		return status;
	if (*index > count)
		return InvalidParameter("index must be between 1 and the shape's adjustment count (" + std::to_string(count) + ")");
	// DISPPARAMS order: rgvarg[0] is the named DISPID_PROPERTYPUT value, rgvarg[1] the index.
	ATL::CComVariant arguments[2] = { ATL::CComVariant(*value), ATL::CComVariant(*index) };
	status = Invoke(adjustments, L"Item", DISPATCH_PROPERTYPUT, arguments, 2, nullptr);
	if (!status.ok())
		return status;
	result["adjustmentCount"] = count;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeRotation(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<double> degrees;
	Status status = ReadNumber(context.parameters, "degrees", degrees, -360.0, 360.0, true);
	if (!status.ok())
		return status;
	return SetProperty(context.shape, L"Rotation", ATL::CComVariant(*degrees));
}

AutomationDispatcher::Status AutomationDispatcher::FlipShape(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<long> direction;
	Status status = ReadChoice(context.parameters, "direction",
		{ { "horizontal", 0 }, { "vertical", 1 } }, direction, true);
	if (!status.ok())
		return status;
	return CallMethod(context.shape, L"Flip", { ATL::CComVariant(*direction) });
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeName(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<std::string> name;
	Status status = ReadString(context.parameters, "name", name, true, false);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	HRESULT hr = Utf8ToVariant(*name, value);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding shape name");
	return SetProperty(context.shape, L"Name", value);
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeBounds(const CommandContext &context,
	nlohmann::json &result)
{
	const nlohmann::json &parameters = context.parameters;
	std::optional<double> left;
	std::optional<double> top;
	std::optional<double> width;
	std::optional<double> height;
	Status status = ReadNumber(parameters, "left", left, -10000.0, 10000.0);
	if (status.ok()) status = ReadNumber(parameters, "top", top, -10000.0, 10000.0);
	if (status.ok()) status = ReadNumber(parameters, "width", width, 0.01, 10000.0);
	if (status.ok()) status = ReadNumber(parameters, "height", height, 0.01, 10000.0);
	if (!status.ok())
		return status;
	if (!left && !top && !width && !height)
		return InvalidParameter("at least one of left, top, width, or height is required");
	const std::pair<const wchar_t *, const std::optional<double> *> members[] = {
		{ L"Left", &left }, { L"Top", &top }, { L"Width", &width }, { L"Height", &height } };
	for (const auto &member : members)
	{
		if (!*member.second)
			continue;
		status = SetProperty(context.shape, member.first, ATL::CComVariant(**member.second));
		if (!status.ok())
			return status;
	}
	nlohmann::json bounds = nlohmann::json::object();
	const std::pair<const wchar_t *, const char *> reads[] = {
		{ L"Left", "left" }, { L"Top", "top" }, { L"Width", "width" }, { L"Height", "height" } };
	for (const auto &read : reads)
	{
		double value = 0.0;
		status = GetDouble(context.shape, read.first, value);
		if (!status.ok())
			return status;
		bounds[read.second] = value;
	}
	result["bounds"] = bounds;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeZOrder(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> order;
	Status status = ReadChoice(context.parameters, "order", {
		{ "front", 0 }, { "back", 1 }, { "forward", 2 }, { "backward", 3 } }, order, true);
	if (!status.ok())
		return status;
	status = CallMethod(context.shape, L"ZOrder", { ATL::CComVariant(*order) });
	if (!status.ok())
		return status;
	long position = 0;
	status = GetInteger(context.shape, L"ZOrderPosition", position);
	if (!status.ok())
		return status;
	result["zOrderPosition"] = position;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::DuplicateShape(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<double> left;
	std::optional<double> top;
	Status status = ReadNumber(context.parameters, "left", left, -10000.0, 10000.0);
	if (status.ok()) status = ReadNumber(context.parameters, "top", top, -10000.0, 10000.0);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> range;
	status = CallObject(context.shape, L"Duplicate", {}, range);
	ATL::CComPtr<IDispatch> duplicate;
	if (status.ok()) status = GetItem(range, 1, duplicate);
	if (!status.ok())
		return status;
	if (left)
	{
		status = SetProperty(duplicate, L"Left", ATL::CComVariant(*left));
		if (!status.ok())
			return status;
	}
	if (top)
	{
		status = SetProperty(duplicate, L"Top", ATL::CComVariant(*top));
		if (!status.ok())
			return status;
	}
	long duplicateId = 0;
	std::string name;
	status = GetInteger(duplicate, L"Id", duplicateId);
	if (status.ok()) status = GetString(duplicate, L"Name", name);
	if (!status.ok())
		return status;
	result["duplicateShapeId"] = duplicateId;
	result["duplicateName"] = name;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::CopyShapeFormat(const CommandContext &context,
	nlohmann::json &result)
{
	(void)result;
	std::optional<long> sourceShapeId;
	Status status = ReadInteger(context.parameters, "sourceShapeId", sourceShapeId, 1, LONG_MAX, true);
	if (!status.ok())
		return status;
	if (*sourceShapeId == context.shapeId)
		return InvalidParameter("sourceShapeId must differ from shapeId");
	ATL::CComPtr<IDispatch> source;
	status = GetShapeById(context.container, *sourceShapeId, source);
	if (!status.ok())
		return status;
	status = CallMethod(source, L"PickUp", {});
	if (!status.ok())
		return status;
	return CallMethod(context.shape, L"Apply", {});
}
