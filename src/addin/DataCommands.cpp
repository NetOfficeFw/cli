#include "pch.h"
#include "AutomationDispatcher.h"

#include <atlsafe.h>
#include <climits>
#include <cmath>
#include <string>

AutomationDispatcher::Status AutomationDispatcher::ReadDataBounds(const nlohmann::json &parameters,
	double &left, double &top, double &width, double &height)
{
	std::optional<double> values[4];
	const char *names[4] = { "left", "top", "width", "height" };
	for (int index = 0; index < 4; ++index)
	{
		const bool position = index < 2;
		Status status = ReadNumber(parameters, names[index], values[index],
			position ? -10000.0 : 0.01, 10000.0, true);
		if (!status.ok())
			return status;
	}
	left = *values[0];
	top = *values[1];
	width = *values[2];
	height = *values[3];
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::ReadAddedDataShape(ATL::CComVariant &value,
	const wchar_t *operation, ATL::CComPtr<IDispatch> &shape, nlohmann::json &result)
{
	Status status = ReadObject(value, operation, shape, -32000);
	if (!status.ok())
		return status;
	long shapeId = 0;
	status = GetInteger(shape, L"Id", shapeId);
	if (!status.ok())
		return status;
	std::string name;
	status = GetString(shape, L"Name", name);
	if (!status.ok())
		return status;
	result["shapeId"] = shapeId;
	result["name"] = std::move(name);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetDataObject(IDispatch *shape, const wchar_t *flag,
	const wchar_t *member, const char *kind, ATL::CComPtr<IDispatch> &result)
{
	bool has = false;
	Status status = GetBoolean(shape, flag, has);
	if (!status.ok())
		return status;
	if (!has)
		return InvalidParameter(std::string("shapeId does not identify a ") + kind);
	return GetObject(shape, member, result);
}

AutomationDispatcher::Status AutomationDispatcher::AddTable(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> rows, columns;
	Status status = ReadInteger(context.parameters, "rows", rows, 1, 75, true);
	if (!status.ok())
		return status;
	status = ReadInteger(context.parameters, "columns", columns, 1, 75, true);
	if (!status.ok())
		return status;
	double left = 0, top = 0, width = 0, height = 0;
	status = ReadDataBounds(context.parameters, left, top, width, height);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> shapes, shape;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	status = CallMethod(shapes, L"AddTable", {
		ATL::CComVariant(*rows), ATL::CComVariant(*columns), ATL::CComVariant(left),
		ATL::CComVariant(top), ATL::CComVariant(width), ATL::CComVariant(height) }, &value);
	if (!status.ok())
		return status;
	return ReadAddedDataShape(value, L"Shapes.AddTable", shape, result);
}

AutomationDispatcher::Status AutomationDispatcher::SetTableCell(const CommandContext &context,
	nlohmann::json &)
{
	std::optional<long> row, column, fillColor, fontColor;
	std::optional<double> fontSize;
	std::optional<bool> bold;
	std::optional<std::string> text;
	// Upper bounds come from the live table (it may have grown past AddTable's 75 limit).
	Status status = ReadInteger(context.parameters, "row", row, 1, INT_MAX, true);
	if (!status.ok())
		return status;
	status = ReadInteger(context.parameters, "column", column, 1, INT_MAX, true);
	if (!status.ok())
		return status;
	status = ReadString(context.parameters, "text", text);
	if (!status.ok())
		return status;
	status = ReadColor(context.parameters, "fillColor", fillColor);
	if (!status.ok())
		return status;
	status = ReadColor(context.parameters, "fontColor", fontColor);
	if (!status.ok())
		return status;
	status = ReadNumber(context.parameters, "fontSize", fontSize, 1, 4000);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "bold", bold);
	if (!status.ok())
		return status;
	if (!text && !fillColor && !fontColor && !fontSize && !bold)
		return InvalidParameter("at least one of text, fillColor, fontColor, fontSize, or bold is required");
	ATL::CComVariant textValue;
	if (text)
	{
		HRESULT hr = Utf8ToVariant(*text, textValue);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding table cell text");
	}

	ATL::CComPtr<IDispatch> table;
	status = GetDataObject(context.shape, L"HasTable", L"Table", "table", table);
	if (!status.ok())
		return status;
	struct Dimension { const wchar_t *collection; const char *name; long requested; };
	for (const Dimension &dimension : { Dimension{ L"Rows", "row", *row }, Dimension{ L"Columns", "column", *column } })
	{
		ATL::CComPtr<IDispatch> collection;
		status = GetObject(table, dimension.collection, collection);
		if (!status.ok())
			return status;
		long count = 0;
		status = GetInteger(collection, L"Count", count);
		if (!status.ok())
			return status;
		if (dimension.requested > count)
			return InvalidParameter(std::string(dimension.name) + " exceeds the table's " +
				dimension.name + " count of " + std::to_string(count));
	}
	ATL::CComPtr<IDispatch> cell, cellShape;
	status = CallObject(table, L"Cell", { ATL::CComVariant(*row), ATL::CComVariant(*column) }, cell);
	if (!status.ok())
		return status;
	status = GetObject(cell, L"Shape", cellShape);
	if (!status.ok())
		return status;

	if (text || fontColor || fontSize || bold)
	{
		ATL::CComPtr<IDispatch> textFrame, textRange;
		status = GetObject(cellShape, L"TextFrame", textFrame);
		if (!status.ok())
			return status;
		status = GetObject(textFrame, L"TextRange", textRange);
		if (!status.ok())
			return status;
		if (text)
		{
			status = SetProperty(textRange, L"Text", textValue);
			if (!status.ok())
				return status;
		}
		if (fontColor || fontSize || bold)
		{
			ATL::CComPtr<IDispatch> font;
			status = GetObject(textRange, L"Font", font);
			if (!status.ok())
				return status;
			if (fontColor)
			{
				ATL::CComPtr<IDispatch> color;
				status = GetObject(font, L"Color", color);
				if (!status.ok())
					return status;
				status = SetProperty(color, L"RGB", ATL::CComVariant(*fontColor));
				if (!status.ok())
					return status;
			}
			if (fontSize)
			{
				status = SetProperty(font, L"Size", ATL::CComVariant(*fontSize));
				if (!status.ok())
					return status;
			}
			if (bold)
			{
				status = SetProperty(font, L"Bold", ATL::CComVariant(*bold ? -1L : 0L));
				if (!status.ok())
					return status;
			}
		}
	}
	if (fillColor)
	{
		ATL::CComPtr<IDispatch> fill, foreColor;
		status = GetObject(cellShape, L"Fill", fill);
		if (!status.ok())
			return status;
		status = SetProperty(fill, L"Visible", ATL::CComVariant(-1L));
		if (!status.ok())
			return status;
		status = CallMethod(fill, L"Solid", {});
		if (!status.ok())
			return status;
		status = GetObject(fill, L"ForeColor", foreColor);
		if (!status.ok())
			return status;
		status = SetProperty(foreColor, L"RGB", ATL::CComVariant(*fillColor));
		if (!status.ok())
			return status;
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AddChart(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> chartType, style;
	Status status = ReadInteger(context.parameters, "chartType", chartType, -4170, 120, true);
	if (!status.ok())
		return status;
	status = ReadInteger(context.parameters, "style", style, -1, 400);
	if (!status.ok())
		return status;
	if (style && *style == 0)
		return InvalidParameter("style must be -1 or an integer from 1 through 400");
	double left = 0, top = 0, width = 0, height = 0;
	status = ReadDataBounds(context.parameters, left, top, width, height);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> shapes, shape;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	status = CallMethod(shapes, L"AddChart2", {
		ATL::CComVariant(style.value_or(-1L)), ATL::CComVariant(*chartType), ATL::CComVariant(left),
		ATL::CComVariant(top), ATL::CComVariant(width), ATL::CComVariant(height),
		ATL::CComVariant(-1L) }, &value);
	if (!status.ok())
		return status;
	status = ReadAddedDataShape(value, L"Shapes.AddChart2", shape, result);
	if (!status.ok())
		return status;
	// PowerPoint may open the chart data workbook; it is not necessarily open, so failures are ignored.
	ATL::CComPtr<IDispatch> chart, chartData, workbook;
	Status optional = GetObject(shape, L"Chart", chart);
	if (optional.ok())
		optional = GetObject(chart, L"ChartData", chartData);
	if (optional.ok())
		optional = GetObject(chartData, L"Workbook", workbook);
	if (optional.ok())
		optional = CallMethod(workbook, L"Close", {});
	if (!optional.ok() && IsStopping())
		return StoppedStatus();
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetChartData(const CommandContext &context,
	nlohmann::json &)
{
	const auto &parameters = context.parameters;
	if (!parameters.contains("categories") || !parameters["categories"].is_array() ||
		parameters["categories"].empty() || parameters["categories"].size() > 1000)
		return InvalidParameter("categories must be an array of 1 through 1000 strings");
	if (!parameters.contains("series") || !parameters["series"].is_array() ||
		parameters["series"].empty() || parameters["series"].size() > 50)
		return InvalidParameter("series must be an array of 1 through 50 {name, values} objects");
	const auto &categories = parameters["categories"];
	const auto &series = parameters["series"];
	const long rowCount = static_cast<long>(categories.size()) + 1;
	const long columnCount = static_cast<long>(series.size()) + 1;

	// Row 1: blank corner + series names; column A: categories; values below each name.
	SAFEARRAYBOUND bounds[2] = { { static_cast<ULONG>(rowCount), 1 }, { static_cast<ULONG>(columnCount), 1 } };
	ATL::CComSafeArray<VARIANT> cells;
	HRESULT hr = cells.Create(bounds, 2);
	if (FAILED(hr))
		return ErrorStatus(-32000, hr, "Building chart data");
	auto put = [&cells](long row, long column, const ATL::CComVariant &value)
	{
		LONG indices[2] = { row, column };
		return cells.MultiDimSetAt(indices, value);
	};
	for (long index = 0; index < rowCount - 1; ++index)
	{
		const auto &category = categories[static_cast<size_t>(index)];
		if (!category.is_string() || category.get_ref<const std::string &>().find('\0') != std::string::npos)
			return InvalidParameter("categories must contain only strings without NUL characters");
		ATL::CComVariant value;
		hr = Utf8ToVariant(category.get_ref<const std::string &>(), value);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding chart category");
		hr = put(index + 2, 1, value);
		if (FAILED(hr))
			return ErrorStatus(-32000, hr, "Building chart data");
	}
	for (long index = 0; index < columnCount - 1; ++index)
	{
		const auto &entry = series[static_cast<size_t>(index)];
		const std::string prefix = "series[" + std::to_string(index) + "]";
		if (!entry.is_object() || !entry.contains("name") || !entry["name"].is_string() ||
			entry["name"].get_ref<const std::string &>().find('\0') != std::string::npos)
			return InvalidParameter(prefix + ".name must be a string without NUL characters");
		if (!entry.contains("values") || !entry["values"].is_array() ||
			entry["values"].size() != categories.size())
			return InvalidParameter(prefix + ".values must be an array with one number per category");
		ATL::CComVariant name;
		hr = Utf8ToVariant(entry["name"].get_ref<const std::string &>(), name);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding chart series name");
		hr = put(1, index + 2, name);
		if (FAILED(hr))
			return ErrorStatus(-32000, hr, "Building chart data");
		long row = 2;
		for (const auto &number : entry["values"])
		{
			if (!number.is_number() || !std::isfinite(number.get<double>()))
				return InvalidParameter(prefix + ".values must contain only finite numbers");
			hr = put(row++, index + 2, ATL::CComVariant(number.get<double>()));
			if (FAILED(hr))
				return ErrorStatus(-32000, hr, "Building chart data");
		}
	}
	auto columnName = [](long column)
	{
		std::wstring name;
		for (; column > 0; column = (column - 1) / 26)
			name.insert(name.begin(), static_cast<wchar_t>(L'A' + (column - 1) % 26));
		return name;
	};
	const std::wstring lastColumn = columnName(columnCount);
	const std::wstring lastRow = std::to_wstring(rowCount);

	ATL::CComPtr<IDispatch> chart, chartData, workbook;
	Status status = GetDataObject(context.shape, L"HasChart", L"Chart", "chart", chart);
	if (!status.ok())
		return status;
	status = GetObject(chart, L"ChartData", chartData);
	if (!status.ok())
		return status;
	status = CallMethod(chartData, L"Activate", {});
	if (!status.ok())
		return status;
	status = GetObject(chartData, L"Workbook", workbook);
	if (!status.ok())
		return status;
	auto write = [&]() -> Status
	{
		ATL::CComPtr<IDispatch> worksheets, worksheet, allCells, header, categoryColumn, block;
		Status step = GetObject(workbook, L"Worksheets", worksheets);
		if (!step.ok())
			return step;
		step = GetItem(worksheets, 1, worksheet);
		if (!step.ok())
			return step;
		ATL::CComVariant sheetName;
		step = Invoke(worksheet, L"Name", DISPATCH_PROPERTYGET, nullptr, 0, &sheetName);
		if (!step.ok())
			return step;
		if (sheetName.vt != VT_BSTR || sheetName.bstrVal == nullptr)
			return ErrorStatus(-32000, DISP_E_TYPEMISMATCH, "Worksheet.Name");
		step = GetObject(worksheet, L"Cells", allCells);
		if (!step.ok())
			return step;
		step = CallMethod(allCells, L"Clear", {});
		if (!step.ok())
			return step;
		// Text format keeps names and categories such as "2024" or "1/2" from becoming numbers or dates.
		const ATL::CComVariant textFormat(L"@");
		step = CallObject(worksheet, L"Range",
			{ ATL::CComVariant((L"A1:" + lastColumn + L"1").c_str()) }, header);
		if (!step.ok())
			return step;
		step = SetProperty(header, L"NumberFormat", textFormat);
		if (!step.ok())
			return step;
		step = CallObject(worksheet, L"Range",
			{ ATL::CComVariant((L"A1:A" + lastRow).c_str()) }, categoryColumn);
		if (!step.ok())
			return step;
		step = SetProperty(categoryColumn, L"NumberFormat", textFormat);
		if (!step.ok())
			return step;
		step = CallObject(worksheet, L"Range",
			{ ATL::CComVariant((L"A1:" + lastColumn + lastRow).c_str()) }, block);
		if (!step.ok())
			return step;
		ATL::CComVariant values;
		values.vt = VT_ARRAY | VT_VARIANT;
		values.parray = cells.Detach();
		step = Invoke(block, L"Value2", DISPATCH_PROPERTYPUT, &values, 1, nullptr);
		if (!step.ok())
			return step;
		std::wstring quotedSheet;
		for (const wchar_t *character = sheetName.bstrVal; *character != L'\0'; ++character)
		{
			quotedSheet += *character;
			if (*character == L'\'')
				quotedSheet += L'\'';
		}
		const std::wstring source = L"='" + quotedSheet + L"'!$A$1:$" + lastColumn + L"$" + lastRow;
		// xlColumns (2): each worksheet column after A is one series.
		return CallMethod(chart, L"SetSourceData", { ATL::CComVariant(source.c_str()), ATL::CComVariant(2L) });
	};
	status = write();
	Status closed = CallMethod(workbook, L"Close", {});
	if (!status.ok())
		return status;
	if (!closed.ok())
		return closed;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetChartTitle(const CommandContext &context,
	nlohmann::json &)
{
	std::optional<std::string> text;
	std::optional<bool> visible;
	Status status = ReadString(context.parameters, "text", text);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "visible", visible);
	if (!status.ok())
		return status;
	if (!text && !visible)
		return InvalidParameter("at least one of text or visible is required");
	if (text && visible && !*visible)
		return InvalidParameter("text requires a visible title; omit visible or pass true");
	ATL::CComVariant textValue;
	if (text)
	{
		HRESULT hr = Utf8ToVariant(*text, textValue);
		if (FAILED(hr))
			return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding chart title");
	}
	ATL::CComPtr<IDispatch> chart;
	status = GetDataObject(context.shape, L"HasChart", L"Chart", "chart", chart);
	if (!status.ok())
		return status;
	const bool show = text ? true : *visible;
	status = SetProperty(chart, L"HasTitle", ATL::CComVariant(show ? -1L : 0L));
	if (!status.ok())
		return status;
	if (text)
	{
		ATL::CComPtr<IDispatch> title;
		status = GetObject(chart, L"ChartTitle", title);
		if (!status.ok())
			return status;
		status = SetProperty(title, L"Text", textValue);
		if (!status.ok())
			return status;
	}
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::GetSmartArtLayouts(const CommandContext &,
	nlohmann::json &result)
{
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> layouts;
	Status status = GetObject(application, L"SmartArtLayouts", layouts);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(layouts, L"Count", count);
	if (!status.ok())
		return status;
	auto entries = nlohmann::json::array();
	for (long index = 1; index <= count; ++index)
	{
		ATL::CComPtr<IDispatch> layout;
		status = GetItem(layouts, index, layout);
		if (!status.ok())
			return status;
		std::string layoutId, name, category;
		status = GetString(layout, L"Id", layoutId);
		if (!status.ok())
			return status;
		status = GetString(layout, L"Name", name);
		if (!status.ok())
			return status;
		status = GetString(layout, L"Category", category);
		if (!status.ok())
			return status;
		entries.push_back({ { "index", index }, { "id", std::move(layoutId) }, { "name", std::move(name) },
			{ "category", std::move(category) } });
	}
	result["layouts"] = std::move(entries);
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::AddSmartArt(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<long> layoutIndex;
	Status status = ReadInteger(context.parameters, "layout", layoutIndex, 1, INT_MAX, true);
	if (!status.ok())
		return status;
	double left = 0, top = 0, width = 0, height = 0;
	status = ReadDataBounds(context.parameters, left, top, width, height);
	if (!status.ok())
		return status;
	ATL::CComPtr<IDispatch> application = m_pApplication;
	ATL::CComPtr<IDispatch> layouts, layout, shapes, shape, smartArt, nodes;
	status = GetObject(application, L"SmartArtLayouts", layouts);
	if (!status.ok())
		return status;
	long count = 0;
	status = GetInteger(layouts, L"Count", count);
	if (!status.ok())
		return status;
	if (*layoutIndex > count)
		return InvalidParameter("layout must be an integer from 1 through " + std::to_string(count));
	status = GetItem(layouts, *layoutIndex, layout);
	if (!status.ok())
		return status;
	status = GetObject(context.container, L"Shapes", shapes);
	if (!status.ok())
		return status;
	ATL::CComVariant value;
	status = CallMethod(shapes, L"AddSmartArt", {
		ATL::CComVariant(layout.p), ATL::CComVariant(left), ATL::CComVariant(top),
		ATL::CComVariant(width), ATL::CComVariant(height) }, &value);
	if (!status.ok())
		return status;
	status = ReadAddedDataShape(value, L"Shapes.AddSmartArt", shape, result);
	if (!status.ok())
		return status;
	status = GetObject(shape, L"SmartArt", smartArt);
	if (!status.ok())
		return status;
	status = GetObject(smartArt, L"AllNodes", nodes);
	if (!status.ok())
		return status;
	long nodeCount = 0;
	status = GetInteger(nodes, L"Count", nodeCount);
	if (!status.ok())
		return status;
	result["nodeCount"] = nodeCount;
	return {};
}

AutomationDispatcher::Status AutomationDispatcher::SetSmartArtNode(const CommandContext &context,
	nlohmann::json &result)
{
	std::optional<std::string> text;
	std::optional<long> index;
	std::optional<bool> add;
	Status status = ReadString(context.parameters, "text", text, true);
	if (!status.ok())
		return status;
	status = ReadInteger(context.parameters, "index", index, 1, INT_MAX);
	if (!status.ok())
		return status;
	status = ReadBoolean(context.parameters, "add", add);
	if (!status.ok())
		return status;
	const bool append = add.value_or(false);
	if (append == index.has_value())
		return InvalidParameter("exactly one of index or add: true is required");
	ATL::CComVariant textValue;
	HRESULT hr = Utf8ToVariant(*text, textValue);
	if (FAILED(hr))
		return ErrorStatus(hr == E_OUTOFMEMORY ? -32000 : -32602, hr, "Encoding SmartArt node text");

	ATL::CComPtr<IDispatch> smartArt, nodes, node, textFrame, textRange;
	status = GetDataObject(context.shape, L"HasSmartArt", L"SmartArt", "SmartArt graphic", smartArt);
	if (!status.ok())
		return status;
	status = GetObject(smartArt, L"AllNodes", nodes);
	if (!status.ok())
		return status;
	long nodeCount = 0;
	status = GetInteger(nodes, L"Count", nodeCount);
	if (!status.ok())
		return status;
	long written = 0;
	if (append)
	{
		status = CallObject(nodes, L"Add", {}, node);
		if (!status.ok())
			return status;
		status = GetInteger(nodes, L"Count", nodeCount);
		if (!status.ok())
			return status;
		// Add appends a top-level node after every existing node in document order.
		written = nodeCount;
	}
	else
	{
		if (*index > nodeCount)
			return InvalidParameter("index must be an integer from 1 through " + std::to_string(nodeCount));
		status = GetItem(nodes, *index, node);
		if (!status.ok())
			return status;
		written = *index;
	}
	status = GetObject(node, L"TextFrame2", textFrame);
	if (!status.ok())
		return status;
	status = GetObject(textFrame, L"TextRange", textRange);
	if (!status.ok())
		return status;
	status = SetProperty(textRange, L"Text", textValue);
	if (!status.ok())
		return status;
	status = GetInteger(nodes, L"Count", nodeCount);
	if (!status.ok())
		return status;
	result["index"] = written;
	result["nodeCount"] = nodeCount;
	return {};
}
