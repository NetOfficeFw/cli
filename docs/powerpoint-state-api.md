# PowerPoint state and test-control API

## Scope and addressing

These target-scoped endpoints inspect and mutate the live PowerPoint object model for test automation. Mutations affect the open, unsaved presentation and never save automatically. HTTP paths use kebab-case for multiword components; JSON fields remain camelCase.

`{target}` is the opaque document ID returned by `/json/list`, not a file name. Reads and mutations never fall back to `Application.ActivePresentation`; this remains safe when several presentations are open. `{slide-id}` is PowerPoint's stable `Slide.SlideID`, not its one-based collection position. Microsoft documents that `SlideID` remains stable when slides are inserted or reordered, while `SlideIndex` represents current ordering ([`Slide.SlideID`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slideid)).

## HTTP routes

| Method | Purpose |
|---|---|
| `GET /json/{target}/presentation` | Presentation metadata, saved state, and slide count |
| `GET /json/{target}/slides` | Ordered slide summaries |
| `GET /json/{target}/slides/{slide-id}` | Slide metadata and top-level shape details |
| `GET /json/{target}/view` | Editing-window view and selection, when available |
| `GET /json/{target}/slide-show` | Slide-show windows for this presentation and their current slide |
| `POST /json/{target}/slides` | Create a slide |
| `POST /json/{target}/slides/{slide-id}/shapes` | Create an auto-shape |
| `PUT /json/{target}/slides/{slide-id}/shapes/{shape-id}` | Replace shape plain text |
| `DELETE /json/{target}/slides/{slide-id}/shapes/{shape-id}` | Delete a shape |
| `DELETE /json/{target}/slides/{slide-id}` | Delete a slide |
| `PUT /json/{target}/view` | Set the target's current editing slide |
| `POST /json/{target}/slide-show` | Start the target's slide show |
| `DELETE /json/{target}/slide-show` | Stop all slide shows for the target |
| `POST /json/{target}/slide-show/navigation` | Navigate a running show |

Mutation requests with bodies require `Content-Type: application/json` and a JSON object. Slide creation accepts optional `layout` from `1` through `12`, defaulting to PowerPoint's blank layout (`12`). Its result is `{ "id", "slideId", "slideIndex" }`.

Shape creation requires numeric `shapeType` (`1..255`), finite `left` and `top`, and positive `width` and `height`. `shapeType` is a PowerPoint `MsoAutoShapeType` numeric value. Creation returns `{ "id", "slideId", "shapeId" }` without changing text. Text replacement is a separate request requiring a string `text`; it replaces all existing plain text and fails if the shape has no writable text frame. Delete results include `deleted: true`.

Set the editing slide with `PUT /view` and `{ "slideId": 259 }`. The first document window for that presentation receives the change. Start runs the presentation's slide show; stop exits all matching show windows. Navigation takes `{ "action": "next" }`, `{ "action": "previous" }`, or `{ "action": "goto", "slideId": 259 }`. `goto` resolves the stable ID to the current slide index. Navigation applies to each running show window for the requested presentation; controlling a show that is not running returns `-32001`.

## WebSocket methods

The same reads and mutations are available on `/devtools/application` using the existing correlated request/response envelope:

| Method | Parameters |
|---|---|
| `PowerPoint.openPresentation` | `{ "path": "C:\\Documents\\Existing.pptx" }` |
| `PowerPoint.getPresentationState` | `{ "targetId": "<target>" }` |
| `PowerPoint.closePresentation` | `{ "targetId": "<target>", "force": true }` (`force` optional; defaults to `false`) |
| `PowerPoint.getSlides` | `{ "targetId": "<target>" }` |
| `PowerPoint.getSlideState` | `{ "targetId": "<target>", "slideId": 259 }` |
| `PowerPoint.getViewState` | `{ "targetId": "<target>" }` |
| `PowerPoint.getSlideShowState` | `{ "targetId": "<target>" }` |
| `PowerPoint.addSlide` | `{ "targetId": "<target>", "layout": 12 }` |
| `PowerPoint.createShape` | `{ "targetId": "<target>", "slideId": 259, "shapeType": 1, "left": 24, "top": 36, "width": 160, "height": 48 }` |
| `PowerPoint.setShapeText` | `{ "targetId": "<target>", "slideId": 259, "shapeId": 7, "text": "Updated" }` |
| `PowerPoint.deleteShape` | `{ "targetId": "<target>", "slideId": 259, "shapeId": 7 }` |
| `PowerPoint.deleteSlide` | `{ "targetId": "<target>", "slideId": 259 }` |
| `PowerPoint.setCurrentSlide` | `{ "targetId": "<target>", "slideId": 259 }` |
| `PowerPoint.startSlideShow` | `{ "targetId": "<target>" }` |
| `PowerPoint.stopSlideShow` | `{ "targetId": "<target>" }` |
| `PowerPoint.navigateSlideShow` | `{ "targetId": "<target>", "action": "goto", "slideId": 259 }` |

`targetId` must be a canonical target GUID for target-scoped methods; `openPresentation` instead requires a non-empty path and returns the opened document descriptor. `slideId` and `shapeId` must be positive int32 identifiers. Invalid parameters return `-32602`; an unknown/closed target, slide, or shape returns `-32004`. The WebSocket result has the same JSON shape as the corresponding HTTP response body. WebSocket is preferred for mutations because CivetWeb cannot detect an HTTP disconnect while its handler waits for PowerPoint's STA; an abandoned HTTP mutation may still execute until its queue deadline.

`PowerPoint.closePresentation` closes only its `targetId`. An unsaved presentation returns `-32005` unless `force` is `true`, which discards its edits; other presentations and the PowerPoint process remain open.

Speaker notes are intentionally not included.


## Response shapes

### Presentation summary

```json
{
  "id": "d17fd90d-90f8-4a29-b204-49495b78d064",
  "name": "Quarterly review.pptx",
  "url": "C:\\Reports\\Quarterly review.pptx",
  "saved": true,
  "readOnly": false,
  "slideCount": 12
}
```

The presentation object provides `Name`, `FullName`, `Saved`, `ReadOnly`, and the `Slides` collection ([`Presentation`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation), [`Presentation.Slides`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.slides)). `url` is empty for never-saved documents, matching the existing document descriptor behavior. `saved` reports Office's dirty/clean flag; this API does not save.

### Slide list

```json
{
  "id": "d17fd90d-90f8-4a29-b204-49495b78d064",
  "slides": [
    { "slideId": 256, "slideIndex": 1, "name": "Title Slide", "hidden": false },
    { "slideId": 259, "slideIndex": 2, "name": "Agenda", "hidden": false }
  ]
}
```

The listing includes `SlideID`, current `SlideIndex`, `Name`, and `SlideShowTransition.Hidden` ([`Presentation.Slides`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.presentation.slides), [`Slide.Name`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.name), [`SlideShowTransition.Hidden`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowtransition.hidden)). A hidden slide still exists; PowerPoint excludes it from slide shows. The list omits shape details to keep the common query bounded.

### Slide and shapes

```json
{
  "id": "d17fd90d-90f8-4a29-b204-49495b78d064",
  "slide": {
    "slideId": 259,
    "slideIndex": 2,
    "name": "Agenda",
    "hidden": false,
    "shapes": [
      {
        "shapeId": 7,
        "zOrderPosition": 1,
        "name": "Title 1",
        "shapeType": 14,
        "placeholderType": 1,
        "text": "Today's agenda",
        "bounds": { "left": 36.0, "top": 24.0, "width": 648.0, "height": 54.0 }
      },
      {
        "shapeId": 12,
        "zOrderPosition": 2,
        "name": "Rectangle 4",
        "shapeType": 1,
        "text": "Review\\nDecisions\\nNext steps",
        "bounds": { "left": 72.0, "top": 120.0, "width": 480.0, "height": 280.0 }
      },
      {
        "shapeId": 15,
        "zOrderPosition": 3,
        "name": "Picture 5",
        "shapeType": 13,
        "text": null,
        "bounds": { "left": 96.0, "top": 220.0, "width": 200.0, "height": 120.0 }
      }
    ]
  }
}
```

`Slide.Shapes` contains top-level placed shapes such as drawings, pictures, OLE objects, text objects, titles, and placeholders ([`Slide.Shapes`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.shapes)). Each entry reports `Shape.Id`, `ZOrderPosition`, `Name`, numeric `Type`, geometry, and plain text if readable ([`Shape.Id`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.id), [`Shape.ZOrderPosition`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.zorderposition), [`Shape.Type`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.type), [`Shape.Left`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.left), [`Shape.Top`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.top), [`Shape.Width`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.width), [`Shape.Height`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.shape.height)). IDs are numeric PowerPoint object-model values; names are included as human-readable metadata, not as unique keys.

Text is extracted as plain text, not lossless text runs or formatting. `text: null` means no plain text could be read for that shape. This endpoint does not promise recursive group members, table cells, chart data, embedded-object contents, alt text, image content, or rich formatting. Those require separate, explicitly bounded APIs.

### Editing-window state

```json
{
  "id": "d17fd90d-90f8-4a29-b204-49495b78d064",
  "available": true,
  "windows": [
    {
      "viewType": 1,
      "currentSlideId": 259,
      "currentSlideIndex": 2,
      "selection": { "type": 2, "shapeIds": [12] }
    }
  ]
}
```

The response includes one entry per `Presentation.Windows` item. `viewType` and selection `type` are PowerPoint enum values. Shape selections include `shapeIds`; other selection types are reported by `type` without an expanded range. The documented window view types include Normal, Slide Sorter, and Notes Page ([`DocumentWindow.ViewType`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.documentwindow.viewtype)); selection types include none, shapes, slides, and text ([`Selection.Type`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.selection.type)). `available` is false with an empty `windows` array when the presentation has no document window. Current-slide and selection fields may be `null` when that view does not expose them.

### Slide-show state

```json
{
  "id": "d17fd90d-90f8-4a29-b204-49495b78d064",
  "running": true,
  "windows": [
    { "currentSlideId": 259, "currentSlideIndex": 2, "currentShowPosition": 2 }
  ]
}
```

The implementation enumerates `Application.SlideShowWindows`, associates each show with its `Presentation`, and reads `SlideShowView.Slide` plus `CurrentShowPosition` ([`Application.SlideShowWindows`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.application.slideshowwindows), [`SlideShowView.Slide`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowview.slide), [`SlideShowView.CurrentShowPosition`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slideshowview.currentshowposition)). `currentShowPosition` is playback position; in a custom show it need not equal the slide's full-presentation index. `running` is true when at least one matching show window exists.

## Consistency and implementation limits

These are live COM reads and mutations, not snapshots parsed from the `.pptx` package. Presentation IDs keep every request scoped to one open document. View/selection/show state is transient and may change between calls, so responses are best-effort point-in-time reads, not a transactionally consistent snapshot. Calls must stay on PowerPoint's owning STA. Mutations remain unsaved until PowerPoint saves; responses are subject to the server's existing 1 MiB message limit. The `netoffice` CLI exposes every read and mutation as a subcommand using the same WebSocket methods; `netoffice presentation list` discovers target IDs via `/json/list`.
