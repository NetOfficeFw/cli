# PowerPoint state inspection API

## Scope and addressing

These are read-only APIs for testing the live PowerPoint object model. HTTP paths follow the requested kebab-case convention only for multiword path components. JSON fields remain camelCase, matching existing responses such as `processId` and `slideCount`.

Every read is document-scoped. `{target}` is the opaque document ID already returned by `/json/list`, not a file name. No route silently reads `Application.ActivePresentation`; that matters when several presentations are open. `{slide-id}` is PowerPoint's `Slide.SlideID`, not its one-based collection position. Microsoft documents that `SlideID` remains stable when slides are inserted or reordered, while `SlideIndex` represents current ordering ([`Slide.SlideID`](https://learn.microsoft.com/en-us/office/vba/api/powerpoint.slide.slideid)).

## HTTP methods

| Method | Purpose |
|---|---|
| `GET /json/{target}/presentation` | Presentation metadata, saved state, and slide count |
| `GET /json/{target}/slides` | Ordered slide summaries |
| `GET /json/{target}/slides/{slide-id}` | Slide metadata and top-level shape details |
| `GET /json/{target}/view` | Editing-window view and selection, when available |
| `GET /json/{target}/slide-show` | Slide-show windows for this presentation and their current slide |

Speaker notes are intentionally not included.

## WebSocket methods

The same reads are available on `/devtools/application` using the existing correlated request/response envelope:

| Method | Parameters |
|---|---|
| `PowerPoint.getPresentationState` | `{ "targetId": "<target>" }` |
| `PowerPoint.getSlides` | `{ "targetId": "<target>" }` |
| `PowerPoint.getSlideState` | `{ "targetId": "<target>", "slideId": 259 }` |
| `PowerPoint.getViewState` | `{ "targetId": "<target>" }` |
| `PowerPoint.getSlideShowState` | `{ "targetId": "<target>" }` |

`targetId` must be a canonical target GUID. `slideId` must be a positive int32 SlideID. Invalid parameters return `-32602`; an unknown/closed target or a slide ID not present in that presentation returns `-32004`. The WebSocket result has the same JSON shape as the corresponding HTTP response body.

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

These are live COM reads, not snapshots parsed from the `.pptx` package. Presentation IDs keep every request scoped to one open document. View/selection/show state is transient and may change between calls, so responses are best-effort point-in-time reads, not a transactionally consistent snapshot. Calls must stay on PowerPoint's owning STA. Responses are subject to the server's existing 1 MiB message limit. The current CLI does not expose these methods; callers use HTTP or WebSocket directly.
