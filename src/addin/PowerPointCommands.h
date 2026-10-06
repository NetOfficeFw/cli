#pragma once

// Table-driven PowerPoint commands. Each row is one granular capability:
//   NETOFFICE_COMMAND(Handler, WebSocket method, HTTP verb, scope, route suffix)
// Scopes resolve the Office object before the handler runs:
//   Application: no document; route /json/<suffix>
//   Target:      document;    route /json/{target}/<suffix>
//   Slide:       slide;       route /json/{target}/slides/{slide-id}/<suffix>
//   Container:   slide or slide master;
//                route /json/{target}/slides/{slide-id}/<suffix> or /json/{target}/master/<suffix>
//   Shape:       shape on a slide or the slide master;
//                route .../shapes/{shape-id}/<suffix>; an empty suffix is the shape route itself.
// Handlers are AutomationDispatcher members named after the first column.
#define NETOFFICE_POWERPOINT_COMMANDS(NETOFFICE_COMMAND) \
	/* Presentation, theme, master, and layouts (PresentationCommands.cpp) */ \
	NETOFFICE_COMMAND(SavePresentation, "PowerPoint.savePresentation", "POST", Target, "presentation/save") \
	NETOFFICE_COMMAND(SetSlideSize, "PowerPoint.setSlideSize", "PUT", Target, "presentation/page-setup") \
	NETOFFICE_COMMAND(ApplyTheme, "PowerPoint.applyTheme", "PUT", Target, "presentation/theme") \
	NETOFFICE_COMMAND(SetThemeColors, "PowerPoint.setThemeColors", "PUT", Target, "presentation/theme/colors") \
	NETOFFICE_COMMAND(SetThemeFonts, "PowerPoint.setThemeFonts", "PUT", Target, "presentation/theme/fonts") \
	NETOFFICE_COMMAND(GetLayouts, "PowerPoint.getLayouts", "GET", Target, "layouts") \
	NETOFFICE_COMMAND(GetMasterState, "PowerPoint.getMasterState", "GET", Target, "master") \
	NETOFFICE_COMMAND(GetLayoutState, "PowerPoint.getLayoutState", "GET", Container, "") \
	NETOFFICE_COMMAND(SetBackground, "PowerPoint.setBackground", "PUT", Container, "background") \
	NETOFFICE_COMMAND(SetHeadersFooters, "PowerPoint.setHeadersFooters", "PUT", Container, "headers-footers") \
	/* Slide operations (SlideCommands.cpp) */ \
	NETOFFICE_COMMAND(MoveSlide, "PowerPoint.moveSlide", "PUT", Slide, "position") \
	NETOFFICE_COMMAND(SetSlideName, "PowerPoint.setSlideName", "PUT", Slide, "name") \
	NETOFFICE_COMMAND(SetSlideNotes, "PowerPoint.setSlideNotes", "PUT", Slide, "notes") \
	NETOFFICE_COMMAND(SetSlideTransition, "PowerPoint.setSlideTransition", "PUT", Slide, "transition") \
	NETOFFICE_COMMAND(SetSlideHidden, "PowerPoint.setSlideHidden", "PUT", Slide, "hidden") \
	NETOFFICE_COMMAND(DuplicateSlide, "PowerPoint.duplicateSlide", "POST", Slide, "duplicate") \
	NETOFFICE_COMMAND(ExportSlide, "PowerPoint.exportSlide", "POST", Slide, "export") \
	NETOFFICE_COMMAND(AddAnimation, "PowerPoint.addAnimation", "POST", Shape, "animations") \
	/* Shape appearance and geometry (ShapeFormatCommands.cpp) */ \
	NETOFFICE_COMMAND(SetShapeFill, "PowerPoint.setShapeFill", "PUT", Shape, "fill") \
	NETOFFICE_COMMAND(SetShapeLine, "PowerPoint.setShapeLine", "PUT", Shape, "line") \
	NETOFFICE_COMMAND(SetShapeShadow, "PowerPoint.setShapeShadow", "PUT", Shape, "shadow") \
	NETOFFICE_COMMAND(SetShapeGlow, "PowerPoint.setShapeGlow", "PUT", Shape, "glow") \
	NETOFFICE_COMMAND(SetShapeSoftEdge, "PowerPoint.setShapeSoftEdge", "PUT", Shape, "soft-edge") \
	NETOFFICE_COMMAND(SetShapeThreeD, "PowerPoint.setShapeThreeD", "PUT", Shape, "three-d") \
	NETOFFICE_COMMAND(SetShapeStyle, "PowerPoint.setShapeStyle", "PUT", Shape, "style") \
	NETOFFICE_COMMAND(SetShapeAdjustment, "PowerPoint.setShapeAdjustment", "PUT", Shape, "adjustments") \
	NETOFFICE_COMMAND(SetShapeRotation, "PowerPoint.setShapeRotation", "PUT", Shape, "rotation") \
	NETOFFICE_COMMAND(FlipShape, "PowerPoint.flipShape", "POST", Shape, "flip") \
	NETOFFICE_COMMAND(SetShapeName, "PowerPoint.setShapeName", "PUT", Shape, "name") \
	NETOFFICE_COMMAND(SetShapeBounds, "PowerPoint.setShapeBounds", "PUT", Shape, "bounds") \
	NETOFFICE_COMMAND(SetShapeZOrder, "PowerPoint.setShapeZOrder", "PUT", Shape, "z-order") \
	NETOFFICE_COMMAND(DuplicateShape, "PowerPoint.duplicateShape", "POST", Shape, "duplicate") \
	NETOFFICE_COMMAND(CopyShapeFormat, "PowerPoint.copyShapeFormat", "POST", Shape, "format") \
	/* Text (TextCommands.cpp) */ \
	NETOFFICE_COMMAND(GetShapeState, "PowerPoint.getShapeState", "GET", Shape, "") \
	NETOFFICE_COMMAND(AddTextbox, "PowerPoint.addTextbox", "POST", Container, "textboxes") \
	NETOFFICE_COMMAND(SetShapeFont, "PowerPoint.setShapeFont", "PUT", Shape, "text/font") \
	NETOFFICE_COMMAND(SetShapeParagraph, "PowerPoint.setShapeParagraph", "PUT", Shape, "text/paragraphs") \
	NETOFFICE_COMMAND(SetShapeTextFrame, "PowerPoint.setShapeTextFrame", "PUT", Shape, "text/frame") \
	/* Lines, pictures, grouping, arrangement (ContentCommands.cpp) */ \
	NETOFFICE_COMMAND(AddLine, "PowerPoint.addLine", "POST", Container, "lines") \
	NETOFFICE_COMMAND(AddConnector, "PowerPoint.addConnector", "POST", Container, "connectors") \
	NETOFFICE_COMMAND(ConnectConnector, "PowerPoint.connectConnector", "PUT", Shape, "connections") \
	NETOFFICE_COMMAND(AddPicture, "PowerPoint.addPicture", "POST", Container, "pictures") \
	NETOFFICE_COMMAND(GroupShapes, "PowerPoint.groupShapes", "POST", Container, "groups") \
	NETOFFICE_COMMAND(UngroupShape, "PowerPoint.ungroupShape", "POST", Shape, "ungroup") \
	NETOFFICE_COMMAND(AlignShapes, "PowerPoint.alignShapes", "POST", Container, "alignment") \
	NETOFFICE_COMMAND(DistributeShapes, "PowerPoint.distributeShapes", "POST", Container, "distribution") \
	/* Tables, charts, SmartArt (DataCommands.cpp) */ \
	NETOFFICE_COMMAND(AddTable, "PowerPoint.addTable", "POST", Container, "tables") \
	NETOFFICE_COMMAND(SetTableCell, "PowerPoint.setTableCell", "PUT", Shape, "cells") \
	NETOFFICE_COMMAND(AddChart, "PowerPoint.addChart", "POST", Container, "charts") \
	NETOFFICE_COMMAND(SetChartData, "PowerPoint.setChartData", "PUT", Shape, "chart/data") \
	NETOFFICE_COMMAND(SetChartTitle, "PowerPoint.setChartTitle", "PUT", Shape, "chart/title") \
	NETOFFICE_COMMAND(GetSmartArtLayouts, "PowerPoint.getSmartArtLayouts", "GET", Application, "smartart-layouts") \
	NETOFFICE_COMMAND(AddSmartArt, "PowerPoint.addSmartArt", "POST", Container, "smartart") \
	NETOFFICE_COMMAND(SetSmartArtNode, "PowerPoint.setSmartArtNode", "PUT", Shape, "smartart/nodes")

enum class CommandScope { Application, Target, Slide, Container, Shape };
