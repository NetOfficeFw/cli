#include "pch.h"
#include "AutomationDispatcher.h"

#include <cctype>
#include <climits>
#include <string>

AutomationDispatcher::Status AutomationDispatcher::MoveSlide(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> index;
	Status status = ReadInteger(context.parameters, "index", index, 1, LONG_MAX, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> slides;
	status = GetObject(context.document, L"Slides", slides);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(slides, L"Count", count);
	if (!status.ok())
		return status;
	if (*index > count)
		return InvalidParameter("index must not exceed the slide count (" + std::to_string(count) + ")");
	status = CallMethod(context.container, L"MoveTo", { ATL::CComVariant(*index) });
	if (!status.ok())
		return status;
	long slideIndex = 0;
	status = GetInteger(context.container, L"SlideIndex", slideIndex);
	if (!status.ok())
		return status;
	result["slideIndex"] = slideIndex;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetSlideName(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> name;
	Status status = ReadString(context.parameters, "name", name, true, false);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	HRESULT hr = Utf8ToVariant(*name, value);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding slide name");
	return SetProperty(context.container, L"Name", value);
}

AutomationDispatcher::Status AutomationDispatcher::SetSlideNotes(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> text;
	Status status = ReadString(context.parameters, "text", text, true);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	HRESULT hr = Utf8ToVariant(*text, value);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding slide notes");
	ATL::CComPtr<IDispatch> notesPage, shapes, placeholders;
	status = GetObject(context.container, L"NotesPage", notesPage);
	if (!status.ok())
		return status;
	status = GetObject(notesPage, L"Shapes", shapes);
	if (!status.ok())
		return status;
	status = GetObject(shapes, L"Placeholders", placeholders);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(placeholders, L"Count", count);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> body;
	for (long index = 1; index <= count && !body; ++index)
	{
		ATL::CComPtr<IDispatch> placeholder, format;
		status = GetItem(placeholders, index, placeholder);
		if (!status.ok())
			return status;
		status = GetObject(placeholder, L"PlaceholderFormat", format);
		if (!status.ok())
			return status;
		long type = 0;
		status = GetInteger(format, L"Type", type);
		if (!status.ok())
			return status;
		if (type == 2) // ppPlaceholderBody
			body = placeholder;
	}
	if (!body)
		return ErrorStatus(-32004, HRESULT_FROM_WIN32(ERROR_NOT_FOUND), "Notes placeholder",
			"The slide's notes page has no body placeholder");
	ATL::CComPtr<IDispatch> textFrame, textRange;
	status = GetObject(body, L"TextFrame", textFrame);
	if (!status.ok())
		return status;
	status = GetObject(textFrame, L"TextRange", textRange);
	if (!status.ok())
		return status;
	return SetProperty(textRange, L"Text", value);
}

AutomationDispatcher::Status AutomationDispatcher::SetSlideTransition(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> effect;
	std::optional<double> duration, advanceAfter;
	std::optional<bool> advanceOnClick;
	Status status = ReadInteger(context.parameters, "effect", effect, 0, 4000);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "duration", duration, 0, 60);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "advanceOnClick", advanceOnClick);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "advanceAfter", advanceAfter, 0, 86400);
	if (!status.ok())
		return status;
	if (!effect && !duration && !advanceOnClick && !advanceAfter)
		return InvalidParameter("at least one of effect, duration, advanceOnClick, or advanceAfter is required");
	ATL::CComPtr<IDispatch> transition;
	status = GetObject(context.container, L"SlideShowTransition", transition);
	if (!status.ok())
		return status;
	// EntryEffect first: choosing an effect resets its duration to the effect's default.
	if (effect)
	{
		status = SetProperty(transition, L"EntryEffect", ATL::CComVariant(*effect));
		if (!status.ok())
			return status;
	}
	if (duration)
	{
		status = SetProperty(transition, L"Duration", ATL::CComVariant(*duration));
		if (!status.ok())
			return status;
	}
	if (advanceOnClick)
	{
		status = SetProperty(transition, L"AdvanceOnClick", ATL::CComVariant(*advanceOnClick ? -1L : 0L));
		if (!status.ok())
			return status;
	}
	if (advanceAfter)
	{
		status = SetProperty(transition, L"AdvanceOnTime", ATL::CComVariant(-1L));
		if (!status.ok())
			return status;
		status = SetProperty(transition, L"AdvanceTime", ATL::CComVariant(*advanceAfter));
		if (!status.ok())
			return status;
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetSlideHidden(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<bool> hidden;
	Status status = ReadBoolean(context.parameters, "hidden", hidden, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> transition;
	status = GetObject(context.container, L"SlideShowTransition", transition);
	if (!status.ok())
		return status;
	return SetProperty(transition, L"Hidden", ATL::CComVariant(*hidden ? -1L : 0L));
}

AutomationDispatcher::Status AutomationDispatcher::DuplicateSlide(const CommandContext &context,
	nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> range, duplicate;
	Status status = CallObject(context.container, L"Duplicate", {}, range);
	if (!status.ok())
		return status;
	status = GetItem(range, 1, duplicate);
	if (!status.ok())
		return status;
	long duplicateId = 0, duplicateIndex = 0;
	status = GetInteger(duplicate, L"SlideID", duplicateId);
	if (!status.ok())
		return status;
	status = GetInteger(duplicate, L"SlideIndex", duplicateIndex);
	if (!status.ok())
		return status;
	result["duplicateSlideId"] = duplicateId;
	result["duplicateSlideIndex"] = duplicateIndex;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ExportSlide(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> path;
	std::optional<long> widthPx, heightPx;
	Status status = ReadPath(context.parameters, "path", path, true);
	if (!status.ok())
		return status;
	status = ReadInteger(context.parameters, "widthPx", widthPx, 1, 10000);
	if (!status.ok())
		return status;
	status = ReadInteger(context.parameters, "heightPx", heightPx, 1, 10000);
	if (!status.ok())
		return status;
	const size_t dot = path->find_last_of('.');
	std::string extension = dot == std::string::npos ? std::string() : path->substr(dot);
	for (char &character : extension)
		character = static_cast<char>(std::tolower(static_cast<unsigned char>(character)));
	const wchar_t *filter = nullptr;
	if (extension == ".png") filter = L"PNG";
	else if (extension == ".jpg" || extension == ".jpeg") filter = L"JPG";
	else if (extension == ".gif") filter = L"GIF";
	else if (extension == ".bmp") filter = L"BMP";
	else if (extension == ".tif" || extension == ".tiff") filter = L"TIF";
	else if (extension == ".svg") filter = L"SVG";
	else
		return InvalidParameter("path must end in .png, .jpg, .jpeg, .gif, .bmp, .tif, .tiff, or .svg");
	ATL::CComVariant fileName;
	HRESULT hr = Utf8ToVariant(*path, fileName);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding export path");
	status = CallMethod(context.container, L"Export", {
		fileName, ATL::CComVariant(filter),
		widthPx ? ATL::CComVariant(*widthPx) : MissingArgument(),
		heightPx ? ATL::CComVariant(*heightPx) : MissingArgument() });
	if (!status.ok())
		return status;
	result["path"] = *path;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AddAnimation(const CommandContext &context,
	nlohmann::json &result)
{
	if (context.master || context.customLayout > 0)
		return InvalidParameter("animations can be added only to slide shapes, not the slide master or a layout");
	std::optional<long> effectId, trigger;
	std::optional<double> duration, delay;
	Status status = ReadInteger(context.parameters, "effect", effectId, 1, 200, true);
	if (!status.ok())
		return status;
	status = ReadChoice(context.parameters, "trigger",
		{ { "on-click", 1L }, { "with-previous", 2L }, { "after-previous", 3L } }, trigger);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "duration", duration, 0.01, 60);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "delay", delay, 0, 60);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> timeLine, sequence, effect;
	status = GetObject(context.container, L"TimeLine", timeLine);
	if (!status.ok())
		return status;
	status = GetObject(timeLine, L"MainSequence", sequence);
	if (!status.ok())
		return status;
	// Sequence.AddEffect(Shape, effectId, Level, trigger, Index)
	status = CallObject(sequence, L"AddEffect", {
		ATL::CComVariant(context.shape), ATL::CComVariant(*effectId), MissingArgument(),
		trigger ? ATL::CComVariant(*trigger) : MissingArgument() }, effect);
	if (!status.ok())
		return status;
	if (duration || delay)
	{
		ATL::CComPtr<IDispatch> timing;
		status = GetObject(effect, L"Timing", timing);
		if (!status.ok())
			return status;
		if (duration)
		{
			status = SetProperty(timing, L"Duration", ATL::CComVariant(*duration));
			if (!status.ok())
				return status;
		}
		if (delay)
		{
			status = SetProperty(timing, L"TriggerDelayTime", ATL::CComVariant(*delay));
			if (!status.ok())
				return status;
		}
	}
	long effectIndex = 0, effectCount = 0;
	status = GetInteger(effect, L"Index", effectIndex);
	if (!status.ok())
		return status;
	status = GetInteger(sequence, L"Count", effectCount);
	if (!status.ok())
		return status;
	result["effectIndex"] = effectIndex;
	result["effectCount"] = effectCount;
	return {};
}
