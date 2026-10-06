#include "pch.h"
#include "AutomationDispatcher.h"

#include <cctype>
#include <string>

AutomationDispatcher::Status AutomationDispatcher::SavePresentation(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> path;
	Status status = ReadPath(context.parameters, "path", path);
	if (!status.ok())
		return status;
	if (path)
	{
		// SaveAs selects the file format explicitly; Office's default ignores the extension.
		const size_t dot = path->find_last_of('.');
		std::string extension = dot == std::string::npos ? std::string() : path->substr(dot);
		for (char &character : extension)
			character = static_cast<char>(std::tolower(static_cast<unsigned char>(character)));
		long format = 0;
		if (extension == ".pptx") format = 24;      // ppSaveAsOpenXMLPresentation
		else if (extension == ".pptm") format = 25; // ppSaveAsOpenXMLPresentationMacroEnabled
		else if (extension == ".potx") format = 26; // ppSaveAsOpenXMLTemplate
		else if (extension == ".ppsx") format = 28; // ppSaveAsOpenXMLShow
		else if (extension == ".pdf") format = 32;  // ppSaveAsPDF
		else
			return InvalidParameter("path must end in .pptx, .pptm, .potx, .ppsx, or .pdf");
		ATL::CComVariant fileName;
		HRESULT hr = Utf8ToVariant(*path, fileName);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding presentation path");
		status = CallMethod(context.document, L"SaveAs", { fileName, ATL::CComVariant(format) });
	}
	else
	{
		std::string existing;
		status = GetString(context.document, L"Path", existing);
		if (!status.ok())
			return status;
		if (existing.empty())
			return InvalidParameter("The presentation has never been saved; provide path");
		status = CallMethod(context.document, L"Save", {});
	}
	if (!status.ok())
		return status;
	nlohmann::json descriptor;
	status = DescribeDocument(context.document, context.id, descriptor);
	if (!status.ok())
		return status;
	long saved = 0;
	status = GetInteger(context.document, L"Saved", saved);
	if (!status.ok())
		return status;
	result["name"] = descriptor["title"];
	result["url"] = descriptor["url"];
	result["saved"] = saved != 0;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetSlideSize(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<double> width, height;
	Status status = ReadNumber(context.parameters, "width", width, 1, 5760, true);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "height", height, 1, 5760, true);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> pageSetup;
	status = GetObject(context.document, L"PageSetup", pageSetup);
	if (!status.ok())
		return status;
	status = SetProperty(pageSetup, L"SlideWidth", ATL::CComVariant(static_cast<float>(*width)));
	if (!status.ok())
		return status;
	status = SetProperty(pageSetup, L"SlideHeight", ATL::CComVariant(static_cast<float>(*height)));
	if (!status.ok())
		return status;
	double actualWidth = 0, actualHeight = 0;
	status = GetDouble(pageSetup, L"SlideWidth", actualWidth);
	if (!status.ok())
		return status;
	status = GetDouble(pageSetup, L"SlideHeight", actualHeight);
	if (!status.ok())
		return status;
	result["slideWidth"] = actualWidth;
	result["slideHeight"] = actualHeight;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ApplyTheme(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> path;
	Status status = ReadPath(context.parameters, "path", path, true);
	if (!status.ok())
		return status;
	const size_t dot = path->find_last_of('.');
	std::string extension = dot == std::string::npos ? std::string() : path->substr(dot);
	for (char &character : extension)
		character = static_cast<char>(std::tolower(static_cast<unsigned char>(character)));
	if (extension != ".thmx" && extension != ".potx" && extension != ".pot" && extension != ".pptx")
		return InvalidParameter("path must end in .thmx, .potx, .pot, or .pptx");
	ATL::CComVariant fileName;
	HRESULT hr = Utf8ToVariant(*path, fileName);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding theme path");
	(void)result;
	return CallMethod(context.document, extension == ".pot" ? L"ApplyTemplate" : L"ApplyTheme", { fileName });
}

namespace
{
	const std::pair<const char *, long> kThemeColors[] = {
		{ "dark1", 1 }, { "light1", 2 }, { "dark2", 3 }, { "light2", 4 },
		{ "accent1", 5 }, { "accent2", 6 }, { "accent3", 7 }, { "accent4", 8 },
		{ "accent5", 9 }, { "accent6", 10 }, { "hyperlink", 11 }, { "followedHyperlink", 12 },
	};
}

AutomationDispatcher::Status AutomationDispatcher::SetThemeColors(const CommandContext &context,
	nlohmann::json &result)
{
	if (!context.parameters.contains("colors") || !context.parameters["colors"].is_object())
		return InvalidParameter("colors must be an object");
	const nlohmann::json &colors = context.parameters["colors"];
	if (colors.empty())
		return InvalidParameter("colors requires at least one theme color member");
	for (auto member = colors.begin(); member != colors.end(); ++member)
	{
		bool known = false;
		for (const auto &entry : kThemeColors)
			known = known || member.key() == entry.first;
		if (!known)
			return InvalidParameter("colors has unknown member " + member.key());
	}
	std::vector<std::pair<long, long>> updates;
	for (const auto &entry : kThemeColors)
	{
		std::optional<long> rgb;
		Status status = ReadColor(colors, entry.first, rgb);
		if (!status.ok())
			return InvalidParameter(std::string("colors.") + status.message);
		if (rgb)
			updates.emplace_back(entry.second, *rgb);
	}
	ATL::CComPtr<IDispatch> master, theme, scheme;
	Status status = GetObject(context.document, L"SlideMaster", master);
	if (!status.ok())
		return status;
	status = GetObject(master, L"Theme", theme);
	if (!status.ok())
		return status;
	status = GetObject(theme, L"ThemeColorScheme", scheme);
	if (!status.ok())
		return status;
	for (const auto &update : updates)
	{
		ATL::CComPtr<IDispatch> color;
		status = CallObject(scheme, L"Colors", { ATL::CComVariant(update.first) }, color);
		if (!status.ok())
			return status;
		status = SetProperty(color, L"RGB", ATL::CComVariant(update.second));
		if (!status.ok())
			return status;
	}
	nlohmann::json readBack = nlohmann::json::object();
	for (const auto &entry : kThemeColors)
	{
		ATL::CComPtr<IDispatch> color;
		status = CallObject(scheme, L"Colors", { ATL::CComVariant(entry.second) }, color);
		if (!status.ok())
			return status;
		long rgb = 0;
		status = GetInteger(color, L"RGB", rgb);
		if (!status.ok())
			return status;
		readBack[entry.first] = ColorToHex(rgb);
	}
	result["colors"] = std::move(readBack);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetThemeFonts(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> major, minor;
	Status status = ReadString(context.parameters, "major", major, false, false);
	if (!status.ok())
		return status;
	status = ReadString(context.parameters, "minor", minor, false, false);
	if (!status.ok())
		return status;
	if (!major && !minor)
		return InvalidParameter("at least one of major or minor is required");
	ATL::CComPtr<IDispatch> master, theme, scheme;
	status = GetObject(context.document, L"SlideMaster", master);
	if (!status.ok())
		return status;
	status = GetObject(master, L"Theme", theme);
	if (!status.ok())
		return status;
	status = GetObject(theme, L"ThemeFontScheme", scheme);
	if (!status.ok())
		return status;
	auto apply = [&](const wchar_t *member, const std::string &name) -> Status
	{
		ATL::CComVariant text;
		HRESULT hr = Utf8ToVariant(name, text);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding font name");
		ATL::CComPtr<IDispatch> fonts, font;
		Status applyStatus = GetObject(scheme, member, fonts);
		if (!applyStatus.ok())
			return applyStatus;
		applyStatus = GetItem(fonts, 1, font); // msoThemeLatin
		if (!applyStatus.ok())
			return applyStatus;
		return SetProperty(font, L"Name", text);
	};
	if (major)
	{
		status = apply(L"MajorFont", *major);
		if (!status.ok())
			return status;
	}
	if (minor)
	{
		status = apply(L"MinorFont", *minor);
		if (!status.ok())
			return status;
	}
	(void)result;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetLayouts(const CommandContext &context,
	nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> master, layouts;
	Status status = GetObject(context.document, L"SlideMaster", master);
	if (!status.ok())
		return status;
	status = GetObject(master, L"CustomLayouts", layouts);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(layouts, L"Count", count);
	if (!status.ok())
		return status;
	nlohmann::json items = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComPtr<IDispatch> layout, shapes, placeholders;
		status = GetItem(layouts, index, layout);
		if (!status.ok())
			return status;
		std::string name;
		status = GetString(layout, L"Name", name);
		if (!status.ok())
			return status;
		status = GetObject(layout, L"Shapes", shapes);
		if (!status.ok())
			return status;
		status = GetObject(shapes, L"Placeholders", placeholders);
		if (!status.ok())
			return status;
		long placeholderCount = 0;
		status = GetInteger(placeholders, L"Count", placeholderCount);
		if (!status.ok())
			return status;
		items.push_back({ { "index", index }, { "name", name }, { "placeholderCount", placeholderCount } });
	}
	result["layouts"] = std::move(items);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetMasterState(const CommandContext &context,
	nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> master, shapes, layouts;
	Status status = GetObject(context.document, L"SlideMaster", master);
	if (!status.ok())
		return status;
	std::string name;
	status = GetString(master, L"Name", name);
	if (!status.ok())
		return status;
	status = GetObject(master, L"Shapes", shapes);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(shapes, L"Count", count);
	if (!status.ok())
		return status;
	nlohmann::json items = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComPtr<IDispatch> shape;
		status = GetItem(shapes, index, shape);
		if (!status.ok())
			return status;
		nlohmann::json item;
		status = DescribeShape(shape, item);
		if (!status.ok())
			return status;
		items.push_back(std::move(item));
	}
	status = GetObject(master, L"CustomLayouts", layouts);
	if (!status.ok())
		return status;
	long layoutCount = 0;
	status = GetInteger(layouts, L"Count", layoutCount);
	if (!status.ok())
		return status;
	result["name"] = name;
	result["shapes"] = std::move(items);
	result["layouts"] = layoutCount;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetLayoutState(const CommandContext &context,
	nlohmann::json &result)
{
	if (context.customLayout == 0)
		return InvalidParameter("getLayoutState requires customLayout; use getSlideState or getMasterState");
	std::string name;
	Status status = GetString(context.container, L"Name", name);
	if (!status.ok())
		return status;
	bool followMaster = false;
	status = GetBoolean(context.container, L"FollowMasterBackground", followMaster);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> shapes;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(shapes, L"Count", count);
	if (!status.ok())
		return status;
	nlohmann::json items = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComPtr<IDispatch> shape;
		status = GetItem(shapes, index, shape);
		if (!status.ok())
			return status;
		nlohmann::json item;
		status = DescribeShape(shape, item);
		if (!status.ok())
			return status;
		items.push_back(std::move(item));
	}
	result["name"] = name;
	result["followMasterBackground"] = followMaster;
	result["shapes"] = std::move(items);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetBackground(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<bool> followMaster;
	std::optional<long> color, color2, gradientStyle;
	std::optional<double> transparency;
	std::optional<std::string> picture;
	Status status = ReadBoolean(context.parameters, "followMaster", followMaster);
	if (!status.ok())
		return status;
	status = ReadColor(context.parameters, "color", color);
	if (!status.ok())
		return status;
	status = ReadColor(context.parameters, "color2", color2);
	if (!status.ok())
		return status;
	status = ReadChoice(context.parameters, "gradientStyle", {
		{ "horizontal", 1 }, { "vertical", 2 }, { "diagonal-up", 3 }, { "diagonal-down", 4 },
		{ "from-corner", 5 }, { "from-center", 7 } }, gradientStyle);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "transparency", transparency, 0, 1);
	if (!status.ok())
		return status;
	status = ReadPath(context.parameters, "picture", picture);
	if (!status.ok())
		return status;
	const bool fill = color || color2 || gradientStyle || transparency || picture;
	if (!followMaster && !fill)
		return InvalidParameter("at least one of followMaster, color, color2, transparency, or picture is required");
	if (followMaster && context.master)
		return InvalidParameter("followMaster is not valid for the slide master");
	if (followMaster && *followMaster && fill)
		return InvalidParameter("followMaster true cannot be combined with other background members");
	if (color2 && !color)
		return InvalidParameter("color2 requires color");
	if (gradientStyle && !color2)
		return InvalidParameter("gradientStyle requires color2");
	if (picture && color)
		return InvalidParameter("picture cannot be combined with color");
	ATL::CComVariant picturePath;
	if (picture)
	{
		HRESULT hr = Utf8ToVariant(*picture, picturePath);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding picture path");
	}
	if (!context.master)
	{
		const bool follow = followMaster && *followMaster;
		status = SetProperty(context.container, L"FollowMasterBackground", ATL::CComVariant(follow ? -1L : 0L));
		if (!status.ok() || follow || !fill)
			return status;
	}
	ATL::CComPtr<IDispatch> background, backgroundFill;
	status = GetObject(context.container, L"Background", background);
	if (!status.ok())
		return status;
	status = GetObject(background, L"Fill", backgroundFill);
	if (!status.ok())
		return status;
	if (color2)
	{
		status = CallMethod(backgroundFill, L"TwoColorGradient",
			{ ATL::CComVariant(gradientStyle ? *gradientStyle : 1L), ATL::CComVariant(1L) });
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> foreColor, backColor;
		status = GetObject(backgroundFill, L"ForeColor", foreColor);
		if (!status.ok())
			return status;
		status = SetProperty(foreColor, L"RGB", ATL::CComVariant(*color));
		if (!status.ok())
			return status;
		status = GetObject(backgroundFill, L"BackColor", backColor);
		if (!status.ok())
			return status;
		status = SetProperty(backColor, L"RGB", ATL::CComVariant(*color2));
		if (!status.ok())
			return status;
	}
	else if (color)
	{
		status = CallMethod(backgroundFill, L"Solid", {});
		if (!status.ok())
			return status;
		ATL::CComPtr<IDispatch> foreColor;
		status = GetObject(backgroundFill, L"ForeColor", foreColor);
		if (!status.ok())
			return status;
		status = SetProperty(foreColor, L"RGB", ATL::CComVariant(*color));
		if (!status.ok())
			return status;
	}
	else if (picture)
	{
		status = CallMethod(backgroundFill, L"UserPicture", { picturePath });
		if (!status.ok())
			return status;
	}
	if (transparency)
	{
		status = SetProperty(backgroundFill, L"Transparency", ATL::CComVariant(static_cast<float>(*transparency)));
		if (!status.ok())
			return status;
	}
	(void)result;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetHeadersFooters(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> text, dateText;
	std::optional<bool> footer, slideNumber, date;
	Status status = ReadString(context.parameters, "text", text);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "footer", footer);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "slideNumber", slideNumber);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "date", date);
	if (!status.ok())
		return status;
	status = ReadString(context.parameters, "dateText", dateText);
	if (!status.ok())
		return status;
	if (!text && !footer && !slideNumber && !date && !dateText)
		return InvalidParameter("at least one of text, footer, slideNumber, date, or dateText is required");
	if (dateText && date && !*date)
		return InvalidParameter("dateText cannot be combined with date false");
	ATL::CComVariant footerText, fixedDate;
	if (text)
	{
		HRESULT hr = Utf8ToVariant(*text, footerText);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding footer text");
	}
	if (dateText)
	{
		HRESULT hr = Utf8ToVariant(*dateText, fixedDate);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding date text");
	}
	ATL::CComPtr<IDispatch> headersFooters;
	status = GetObject(context.container, L"HeadersFooters", headersFooters);
	if (!status.ok())
		return status;
	auto tri = [](bool value) { return ATL::CComVariant(value ? -1L : 0L); };
	if (text || footer)
	{
		ATL::CComPtr<IDispatch> footerObject;
		status = GetObject(headersFooters, L"Footer", footerObject);
		if (!status.ok())
			return status;
		status = SetProperty(footerObject, L"Visible", tri(footer ? *footer : true));
		if (!status.ok())
			return status;
		if (text)
		{
			status = SetProperty(footerObject, L"Text", footerText);
			if (!status.ok())
				return status;
		}
	}
	if (slideNumber)
	{
		ATL::CComPtr<IDispatch> number;
		status = GetObject(headersFooters, L"SlideNumber", number);
		if (!status.ok())
			return status;
		status = SetProperty(number, L"Visible", tri(*slideNumber));
		if (!status.ok())
			return status;
	}
	if (date || dateText)
	{
		ATL::CComPtr<IDispatch> dateAndTime;
		status = GetObject(headersFooters, L"DateAndTime", dateAndTime);
		if (!status.ok())
			return status;
		status = SetProperty(dateAndTime, L"Visible", tri(date ? *date : true));
		if (!status.ok())
			return status;
		if (dateText)
		{
			status = SetProperty(dateAndTime, L"UseFormat", tri(false));
			if (!status.ok())
				return status;
			status = SetProperty(dateAndTime, L"Text", fixedDate);
			if (!status.ok())
				return status;
		}
	}
	(void)result;
	return {};
}
