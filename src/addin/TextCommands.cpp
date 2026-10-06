#include "pch.h"
#include "AutomationDispatcher.h"

#include <climits>
#include <string>
#include <utility>

namespace
{
	using Choices = std::initializer_list<std::pair<const char *, long>>;

	// Writable names first; read-only names only label values read back from Office.
	const Choices TextAlignments = { { "left", 1 }, { "center", 2 }, { "right", 3 }, { "justify", 4 } };
	const Choices TextAlignmentsReadOnly = { { "distribute", 5 }, { "thai-distribute", 6 },
		{ "justify-low", 7 } };
	const Choices TextAnchors = { { "top", 1 }, { "middle", 3 }, { "bottom", 4 } };
	const Choices TextAnchorsReadOnly = { { "top-baseline", 2 }, { "bottom-baseline", 5 } };
	const Choices TextAutoSizes = { { "none", 0 }, { "shape-to-fit-text", 1 } };
	const Choices TextOrientations = { { "horizontal", 1 }, { "upward", 2 }, { "downward", 3 },
		{ "vertical", 5 } };
	const Choices TextOrientationsReadOnly = { { "vertical-far-east", 4 },
		{ "horizontal-rotated-far-east", 6 } };
	const Choices BulletKinds = { { "none", 0 }, { "bullet", 1 }, { "number", 2 } };
	const Choices BulletKindsReadOnly = { { "picture", 3 } };

	// Office reports -2 for mixed values; unnamed values stay numeric.
	nlohmann::json ChoiceJson(long value, Choices choices, Choices readOnly)
	{
		for (const auto &choice : choices)
			if (choice.second == value)
				return choice.first;
		for (const auto &choice : readOnly)
			if (choice.second == value)
				return choice.first;
		if (value == -2)
			return "mixed";
		return value;
	}

	// MsoTriState: msoTrue -1 (msoCTrue 1), msoFalse 0, msoTriStateMixed -2.
	nlohmann::json TriStateJson(long value)
	{
		if (value == 0)
			return false;
		if (value == -1 || value == 1)
			return true;
		return "mixed";
	}

	nlohmann::json AlignmentJson(long value) { return ChoiceJson(value, TextAlignments, TextAlignmentsReadOnly); }
	nlohmann::json AnchorJson(long value) { return ChoiceJson(value, TextAnchors, TextAnchorsReadOnly); }
	nlohmann::json AutoSizeJson(long value) { return ChoiceJson(value, TextAutoSizes, {}); }
	nlohmann::json OrientationJson(long value) { return ChoiceJson(value, TextOrientations, TextOrientationsReadOnly); }
	nlohmann::json IntegerJson(long value) { return value; }

	ATL::CComVariant TriState(bool value)
	{
		return ATL::CComVariant(value ? -1L : 0L);
	}
}

AutomationDispatcher::Status AutomationDispatcher::GetShapeTextFrame(IDispatch *shape,
	ATL::CComPtr<IDispatch> &textFrame)
{
	long hasTextFrame = 0;
	Status status = GetInteger(shape, L"HasTextFrame", hasTextFrame);
	if (!status.ok())
		return status;
	if (hasTextFrame == 0)
		return InvalidParameter("The shape has no text frame");
	return GetObject(shape, L"TextFrame", textFrame);
}

AutomationDispatcher::Status AutomationDispatcher::ReadStateInteger(IDispatch *object,
	const wchar_t *name, nlohmann::json &target, const char *key, nlohmann::json (*convert)(long))
{
	long value = 0;
	Status status = GetInteger(object, name, value);
	if (status.ok())
		target[key] = convert(value);
	return status.ok() || !IsStopping() ? Status{} : StoppedStatus();
}

AutomationDispatcher::Status AutomationDispatcher::ReadStateNumber(IDispatch *object,
	const wchar_t *name, nlohmann::json &target, const char *key)
{
	double value = 0;
	Status status = GetDouble(object, name, value);
	if (status.ok())
		target[key] = value;
	return status.ok() || !IsStopping() ? Status{} : StoppedStatus();
}

AutomationDispatcher::Status AutomationDispatcher::ReadStateString(IDispatch *object,
	const wchar_t *name, nlohmann::json &target, const char *key)
{
	std::string value;
	Status status = GetString(object, name, value);
	if (status.ok())
		target[key] = std::move(value);
	return status.ok() || !IsStopping() ? Status{} : StoppedStatus();
}

AutomationDispatcher::Status AutomationDispatcher::ReadStateColor(IDispatch *object,
	const wchar_t *name, nlohmann::json &target, const char *key)
{
	ATL::CComPtr<IDispatch> color;
	Status status = GetObject(object, name, color);
	long rgb = 0;
	if (status.ok())
		status = GetInteger(color, L"RGB", rgb);
	if (status.ok())
		target[key] = ColorToHex(rgb);
	return status.ok() || !IsStopping() ? Status{} : StoppedStatus();
}

AutomationDispatcher::Status AutomationDispatcher::DescribeTextParagraph(IDispatch *paragraph,
	long index, nlohmann::json &result)
{
	std::string text;
	Status status = GetString(paragraph, L"Text", text);
	if (!status.ok())
		return status;
	while (!text.empty() && text.back() == '\r')
		text.pop_back();
	result = { { "index", index }, { "text", std::move(text) } };
	status = ReadStateInteger(paragraph, L"IndentLevel", result, "indentLevel", IntegerJson);
	if (!status.ok())
		return status;

	ATL::CComPtr<IDispatch> format;
	status = GetObject(paragraph, L"ParagraphFormat", format);
	if (status.ok())
	{
		// Spacing is in points when LineRule* is msoFalse, otherwise in lines.
		auto spacing = [&](const wchar_t *rule, const wchar_t *space, const char *pointsKey,
			const char *linesKey) -> Status
		{
			long lines = 0;
			Status read = GetInteger(format, rule, lines);
			if (!read.ok())
				return IsStopping() ? StoppedStatus() : Status{};
			return ReadStateNumber(format, space, result, lines != 0 ? linesKey : pointsKey);
		};
		status = ReadStateInteger(format, L"Alignment", result, "alignment", AlignmentJson);
		if (status.ok())
			status = spacing(L"LineRuleBefore", L"SpaceBefore", "spaceBefore", "spaceBeforeLines");
		if (status.ok())
			status = spacing(L"LineRuleAfter", L"SpaceAfter", "spaceAfter", "spaceAfterLines");
		if (status.ok())
			status = spacing(L"LineRuleWithin", L"SpaceWithin", "lineSpacingPoints", "lineSpacing");
		if (!status.ok())
			return status;

		ATL::CComPtr<IDispatch> bullet;
		long visible = 0, type = 0;
		Status read = GetObject(format, L"Bullet", bullet);
		if (read.ok())
			read = GetInteger(bullet, L"Visible", visible);
		if (read.ok())
			read = GetInteger(bullet, L"Type", type);
		if (read.ok())
		{
			result["bullet"] = visible == 0 ? nlohmann::json("none") :
				ChoiceJson(type, BulletKinds, BulletKindsReadOnly);
			long character = 0;
			if (visible != 0 && type == 1)
				read = GetInteger(bullet, L"Character", character);
			if (read.ok() && character > 0 && character <= 0xFFFF)
			{
				const OLECHAR unit = static_cast<OLECHAR>(character);
				ATL::CComBSTR value(1, &unit);
				std::string utf8;
				if (value.m_str != nullptr && SUCCEEDED(BstrToUtf8(value, utf8)))
					result["bulletCharacter"] = std::move(utf8);
			}
		}
		if (!read.ok() && IsStopping())
			return StoppedStatus();
	}
	else if (IsStopping())
		return StoppedStatus();

	ATL::CComPtr<IDispatch> runs;
	status = CallObject(paragraph, L"Runs", {}, runs);
	if (!status.ok())
		return status;
	long runCount = 0;
	status = GetInteger(runs, L"Count", runCount);
	if (!status.ok())
		return status;
	auto runResults = nlohmann::json::array();
	for (long runIndex = 1; runIndex <= runCount; ++runIndex)
	{
		ATL::CComPtr<IDispatch> run;
		status = CallObject(paragraph, L"Runs", { ATL::CComVariant(runIndex), ATL::CComVariant(1L) }, run);
		if (!status.ok())
			return status;
		long start = 0, length = 0;
		std::string runText;
		status = GetInteger(run, L"Start", start);
		if (status.ok())
			status = GetInteger(run, L"Length", length);
		if (status.ok())
			status = GetString(run, L"Text", runText);
		if (!status.ok())
			return status;
		nlohmann::json runResult = { { "start", start }, { "length", length }, { "text", std::move(runText) } };
		ATL::CComPtr<IDispatch> font;
		Status read = GetObject(run, L"Font", font);
		if (read.ok())
		{
			nlohmann::json fontResult = nlohmann::json::object();
			status = ReadStateString(font, L"Name", fontResult, "name");
			if (status.ok())
				status = ReadStateNumber(font, L"Size", fontResult, "size");
			if (status.ok())
				status = ReadStateInteger(font, L"Bold", fontResult, "bold", TriStateJson);
			if (status.ok())
				status = ReadStateInteger(font, L"Italic", fontResult, "italic", TriStateJson);
			if (status.ok())
				status = ReadStateInteger(font, L"Underline", fontResult, "underline", TriStateJson);
			if (status.ok())
				status = ReadStateColor(font, L"Color", fontResult, "color");
			if (!status.ok())
				return status;
			runResult["font"] = std::move(fontResult);
		}
		else if (IsStopping())
			return StoppedStatus();
		runResults.push_back(std::move(runResult));
	}
	result["runs"] = std::move(runResults);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetShapeState(const CommandContext &context,
	nlohmann::json &result)
{
	IDispatch *shape = context.shape;
	nlohmann::json description;
	Status status = DescribeShape(shape, description);
	if (!status.ok())
		return status;
	for (auto &item : description.items())
		result[item.key()] = std::move(item.value());

	status = ReadStateNumber(shape, L"Rotation", result, "rotation");
	if (!status.ok())
		return status;

	ATL::CComPtr<IDispatch> fill;
	Status read = GetObject(shape, L"Fill", fill);
	if (read.ok())
	{
		nlohmann::json fillResult = nlohmann::json::object();
		status = ReadStateInteger(fill, L"Visible", fillResult, "visible", TriStateJson);
		if (status.ok())
			status = ReadStateColor(fill, L"ForeColor", fillResult, "color");
		if (status.ok())
			status = ReadStateNumber(fill, L"Transparency", fillResult, "transparency");
		if (!status.ok())
			return status;
		result["fill"] = std::move(fillResult);
	}
	else if (IsStopping())
		return StoppedStatus();

	ATL::CComPtr<IDispatch> line;
	read = GetObject(shape, L"Line", line);
	if (read.ok())
	{
		nlohmann::json lineResult = nlohmann::json::object();
		status = ReadStateInteger(line, L"Visible", lineResult, "visible", TriStateJson);
		if (status.ok())
			status = ReadStateColor(line, L"ForeColor", lineResult, "color");
		if (status.ok())
			status = ReadStateNumber(line, L"Weight", lineResult, "weight");
		if (!status.ok())
			return status;
		result["line"] = std::move(lineResult);
	}
	else if (IsStopping())
		return StoppedStatus();

	long hasTextFrame = 0;
	status = GetInteger(shape, L"HasTextFrame", hasTextFrame);
	if (!status.ok())
		return status;
	result["hasTextFrame"] = hasTextFrame != 0;
	if (hasTextFrame == 0)
		return {};

	ATL::CComPtr<IDispatch> textFrame;
	status = GetObject(shape, L"TextFrame", textFrame);
	if (!status.ok())
		return status;
	nlohmann::json frameResult = nlohmann::json::object();
	status = ReadStateInteger(textFrame, L"VerticalAnchor", frameResult, "anchor", AnchorJson);
	if (status.ok())
		status = ReadStateInteger(textFrame, L"Orientation", frameResult, "orientation", OrientationJson);
	if (status.ok())
		status = ReadStateNumber(textFrame, L"MarginLeft", frameResult, "marginLeft");
	if (status.ok())
		status = ReadStateNumber(textFrame, L"MarginRight", frameResult, "marginRight");
	if (status.ok())
		status = ReadStateNumber(textFrame, L"MarginTop", frameResult, "marginTop");
	if (status.ok())
		status = ReadStateNumber(textFrame, L"MarginBottom", frameResult, "marginBottom");
	if (status.ok())
		status = ReadStateInteger(textFrame, L"AutoSize", frameResult, "autoSize", AutoSizeJson);
	if (status.ok())
		status = ReadStateInteger(textFrame, L"WordWrap", frameResult, "wordWrap", TriStateJson);
	if (!status.ok())
		return status;
	result["textFrame"] = std::move(frameResult);

	ATL::CComPtr<IDispatch> textRange, paragraphs;
	status = GetObject(textFrame, L"TextRange", textRange);
	if (!status.ok())
		return status;
	status = CallObject(textRange, L"Paragraphs", {}, paragraphs);
	if (!status.ok())
		return status;
	long paragraphCount = 0;
	status = GetInteger(paragraphs, L"Count", paragraphCount);
	if (!status.ok())
		return status;
	auto paragraphResults = nlohmann::json::array();
	for (long index = 1; index <= paragraphCount; ++index)
	{
		ATL::CComPtr<IDispatch> paragraph;
		status = CallObject(textRange, L"Paragraphs", { ATL::CComVariant(index), ATL::CComVariant(1L) },
			paragraph);
		if (!status.ok())
			return status;
		nlohmann::json paragraphResult;
		status = DescribeTextParagraph(paragraph, index, paragraphResult);
		if (!status.ok())
			return status;
		paragraphResults.push_back(std::move(paragraphResult));
	}
	result["paragraphs"] = std::move(paragraphResults);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AddTextbox(const CommandContext &context,
	nlohmann::json &result)
{
	const auto &parameters = context.parameters;
	std::optional<double> left, top, width, height;
	std::optional<long> orientation;
	Status status = ReadNumber(parameters, "left", left, -10000, 10000, true);
	if (status.ok())
		status = ReadNumber(parameters, "top", top, -10000, 10000, true);
	if (status.ok())
		status = ReadNumber(parameters, "width", width, 0, 10000, true);
	if (status.ok())
		status = ReadNumber(parameters, "height", height, 0, 10000, true);
	if (status.ok())
		status = ReadChoice(parameters, "orientation", TextOrientations, orientation);
	if (!status.ok())
		return status;
	if (*width <= 0 || *height <= 0)
		return InvalidParameter("width and height must be greater than 0");

	ATL::CComPtr<IDispatch> shapes, shape;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	status = CallObject(shapes, L"AddTextbox", { ATL::CComVariant(orientation.value_or(1L)),
		ATL::CComVariant(*left), ATL::CComVariant(*top), ATL::CComVariant(*width),
		ATL::CComVariant(*height) }, shape);
	if (!status.ok())
		return status;
	long shapeId = 0;
	std::string name;
	status = GetInteger(shape, L"Id", shapeId);
	if (status.ok())
		status = GetString(shape, L"Name", name);
	if (!status.ok())
		return status;
	result["shapeId"] = shapeId;
	result["name"] = std::move(name);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeFont(const CommandContext &context,
	nlohmann::json &result)
{
	const auto &parameters = context.parameters;
	std::optional<long> start, length, color;
	std::optional<std::string> name;
	std::optional<double> size;
	std::optional<bool> bold, italic, underline;
	Status status = ReadInteger(parameters, "start", start, 1, LONG_MAX);
	if (status.ok())
		status = ReadInteger(parameters, "length", length, 1, LONG_MAX);
	if (status.ok())
		status = ReadString(parameters, "name", name, false, false);
	if (status.ok())
		status = ReadNumber(parameters, "size", size, 1, 4000);
	if (status.ok())
		status = ReadBoolean(parameters, "bold", bold);
	if (status.ok())
		status = ReadBoolean(parameters, "italic", italic);
	if (status.ok())
		status = ReadBoolean(parameters, "underline", underline);
	if (status.ok())
		status = ReadColor(parameters, "color", color);
	if (!status.ok())
		return status;
	if (start.has_value() != length.has_value())
		return InvalidParameter("start and length must be given together");
	if (!name && !size && !bold && !italic && !underline && !color)
		return InvalidParameter("at least one of name, size, bold, italic, underline, or color is required");
	ATL::CComVariant fontName;
	if (name)
	{
		HRESULT hr = Utf8ToVariant(*name, fontName);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding font name");
	}

	ATL::CComPtr<IDispatch> textFrame, target;
	status = GetShapeTextFrame(context.shape, textFrame);
	if (status.ok())
		status = GetObject(textFrame, L"TextRange", target);
	if (!status.ok())
		return status;
	if (start)
	{
		long textLength = 0;
		status = GetInteger(target, L"Length", textLength);
		if (!status.ok())
			return status;
		if (*start > textLength)
			return InvalidParameter("start must not exceed the text length (" +
				std::to_string(textLength) + ")");
		ATL::CComPtr<IDispatch> characters;
		status = CallObject(target, L"Characters", { ATL::CComVariant(*start), ATL::CComVariant(*length) },
			characters);
		if (!status.ok())
			return status;
		target = characters;
	}
	ATL::CComPtr<IDispatch> font;
	status = GetObject(target, L"Font", font);
	if (!status.ok())
		return status;
	if (name)
		status = SetProperty(font, L"Name", fontName);
	if (status.ok() && size)
		status = SetProperty(font, L"Size", ATL::CComVariant(*size));
	if (status.ok() && bold)
		status = SetProperty(font, L"Bold", TriState(*bold));
	if (status.ok() && italic)
		status = SetProperty(font, L"Italic", TriState(*italic));
	if (status.ok() && underline)
		status = SetProperty(font, L"Underline", TriState(*underline));
	if (status.ok() && color)
	{
		ATL::CComPtr<IDispatch> colorFormat;
		status = GetObject(font, L"Color", colorFormat);
		if (status.ok())
			status = SetProperty(colorFormat, L"RGB", ATL::CComVariant(*color));
	}
	return status;
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeParagraph(const CommandContext &context,
	nlohmann::json &result)
{
	const auto &parameters = context.parameters;
	std::optional<long> paragraph, alignment, indentLevel, bullet;
	std::optional<double> spaceBefore, spaceAfter, lineSpacing;
	std::optional<std::string> bulletCharacter;
	Status status = ReadInteger(parameters, "paragraph", paragraph, 1, LONG_MAX);
	if (status.ok())
		status = ReadChoice(parameters, "alignment", TextAlignments, alignment);
	if (status.ok())
		status = ReadNumber(parameters, "spaceBefore", spaceBefore, 0, 1584);
	if (status.ok())
		status = ReadNumber(parameters, "spaceAfter", spaceAfter, 0, 1584);
	if (status.ok())
		status = ReadNumber(parameters, "lineSpacing", lineSpacing, 0, 1584);
	if (status.ok())
		status = ReadInteger(parameters, "indentLevel", indentLevel, 1, 9);
	if (status.ok())
		status = ReadChoice(parameters, "bullet", BulletKinds, bullet);
	if (status.ok())
		status = ReadString(parameters, "bulletCharacter", bulletCharacter, false, false);
	if (!status.ok())
		return status;
	if (!alignment && !spaceBefore && !spaceAfter && !lineSpacing && !indentLevel && !bullet &&
		!bulletCharacter)
		return InvalidParameter("at least one of alignment, spaceBefore, spaceAfter, lineSpacing, "
			"indentLevel, bullet, or bulletCharacter is required");
	long character = 0;
	if (bulletCharacter)
	{
		if (bullet && *bullet != 1)
			return InvalidParameter("bulletCharacter requires bullet to be bullet");
		ATL::CComVariant encoded;
		HRESULT hr = Utf8ToVariant(*bulletCharacter, encoded);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding bulletCharacter");
		// Valid UTF-8 yields two UTF-16 units for non-BMP characters, never a lone surrogate.
		if (encoded.vt != VT_BSTR || SysStringLen(encoded.bstrVal) != 1)
			return InvalidParameter("bulletCharacter must be a single Basic Multilingual Plane character");
		character = static_cast<long>(encoded.bstrVal[0]);
		bullet = 1;
	}

	ATL::CComPtr<IDispatch> textFrame, target;
	status = GetShapeTextFrame(context.shape, textFrame);
	if (status.ok())
		status = GetObject(textFrame, L"TextRange", target);
	if (!status.ok())
		return status;
	if (paragraph)
	{
		ATL::CComPtr<IDispatch> paragraphs, selected;
		long count = 0;
		status = CallObject(target, L"Paragraphs", {}, paragraphs);
		if (status.ok())
			status = GetInteger(paragraphs, L"Count", count);
		if (!status.ok())
			return status;
		if (*paragraph > count)
			return InvalidParameter("paragraph must not exceed the paragraph count (" +
				std::to_string(count) + ")");
		status = CallObject(target, L"Paragraphs", { ATL::CComVariant(*paragraph), ATL::CComVariant(1L) },
			selected);
		if (!status.ok())
			return status;
		target = selected;
	}

	if (alignment || spaceBefore || spaceAfter || lineSpacing || bullet)
	{
		ATL::CComPtr<IDispatch> format;
		status = GetObject(target, L"ParagraphFormat", format);
		if (status.ok() && alignment)
			status = SetProperty(format, L"Alignment", ATL::CComVariant(*alignment));
		if (status.ok() && spaceBefore)
		{
			status = SetProperty(format, L"LineRuleBefore", TriState(false));
			if (status.ok())
				status = SetProperty(format, L"SpaceBefore", ATL::CComVariant(*spaceBefore));
		}
		if (status.ok() && spaceAfter)
		{
			status = SetProperty(format, L"LineRuleAfter", TriState(false));
			if (status.ok())
				status = SetProperty(format, L"SpaceAfter", ATL::CComVariant(*spaceAfter));
		}
		if (status.ok() && lineSpacing)
		{
			status = SetProperty(format, L"LineRuleWithin", TriState(true));
			if (status.ok())
				status = SetProperty(format, L"SpaceWithin", ATL::CComVariant(*lineSpacing));
		}
		if (status.ok() && bullet)
		{
			ATL::CComPtr<IDispatch> bulletFormat;
			status = GetObject(format, L"Bullet", bulletFormat);
			if (status.ok())
				status = SetProperty(bulletFormat, L"Visible", TriState(*bullet != 0));
			if (status.ok() && *bullet != 0)
				status = SetProperty(bulletFormat, L"Type", ATL::CComVariant(*bullet));
			if (status.ok() && bulletCharacter)
				status = SetProperty(bulletFormat, L"Character", ATL::CComVariant(character));
		}
		if (!status.ok())
			return status;
	}
	if (indentLevel)
		status = SetProperty(target, L"IndentLevel", ATL::CComVariant(*indentLevel));
	return status;
}

AutomationDispatcher::Status AutomationDispatcher::SetShapeTextFrame(const CommandContext &context,
	nlohmann::json &result)
{
	const auto &parameters = context.parameters;
	std::optional<long> anchor, autoSize;
	std::optional<double> marginLeft, marginRight, marginTop, marginBottom;
	std::optional<bool> wordWrap;
	Status status = ReadChoice(parameters, "anchor", TextAnchors, anchor);
	if (status.ok())
		status = ReadNumber(parameters, "marginLeft", marginLeft, 0, 1584);
	if (status.ok())
		status = ReadNumber(parameters, "marginRight", marginRight, 0, 1584);
	if (status.ok())
		status = ReadNumber(parameters, "marginTop", marginTop, 0, 1584);
	if (status.ok())
		status = ReadNumber(parameters, "marginBottom", marginBottom, 0, 1584);
	if (status.ok())
		status = ReadChoice(parameters, "autoSize", TextAutoSizes, autoSize);
	if (status.ok())
		status = ReadBoolean(parameters, "wordWrap", wordWrap);
	if (!status.ok())
		return status;
	if (!anchor && !marginLeft && !marginRight && !marginTop && !marginBottom && !autoSize && !wordWrap)
		return InvalidParameter("at least one of anchor, marginLeft, marginRight, marginTop, "
			"marginBottom, autoSize, or wordWrap is required");

	ATL::CComPtr<IDispatch> textFrame;
	status = GetShapeTextFrame(context.shape, textFrame);
	if (status.ok() && anchor)
		status = SetProperty(textFrame, L"VerticalAnchor", ATL::CComVariant(*anchor));
	if (status.ok() && marginLeft)
		status = SetProperty(textFrame, L"MarginLeft", ATL::CComVariant(*marginLeft));
	if (status.ok() && marginRight)
		status = SetProperty(textFrame, L"MarginRight", ATL::CComVariant(*marginRight));
	if (status.ok() && marginTop)
		status = SetProperty(textFrame, L"MarginTop", ATL::CComVariant(*marginTop));
	if (status.ok() && marginBottom)
		status = SetProperty(textFrame, L"MarginBottom", ATL::CComVariant(*marginBottom));
	if (status.ok() && wordWrap)
		status = SetProperty(textFrame, L"WordWrap", TriState(*wordWrap));
	if (status.ok() && autoSize)
		status = SetProperty(textFrame, L"AutoSize", ATL::CComVariant(*autoSize));
	return status;
}
