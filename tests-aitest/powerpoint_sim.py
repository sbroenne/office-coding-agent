"""In-memory PowerPoint simulator for manifest-driven AI tool tests."""

from __future__ import annotations

import base64
import re
from dataclasses import dataclass, field
from typing import Any

from tool_result import ToolResult


@dataclass(slots=True)
class SlideShape:
    kind: str
    text: str = ""
    name: str = ""
    meta: dict[str, Any] = field(default_factory=dict)


@dataclass(slots=True)
class Slide:
    shapes: list[SlideShape] = field(default_factory=list)


class PowerPointSimulator:
    """Minimal stateful presentation simulator."""

    def __init__(self) -> None:
        self.slides: list[Slide] = [
            Slide(shapes=[SlideShape(kind="title", text="Office Coding Agent Overview", name="Title 1")])
        ]

    def _ok(self, value: Any) -> ToolResult:
        return ToolResult(success=True, value=value)

    def _error(self, message: str) -> ToolResult:
        return ToolResult(success=False, error=message)

    def _fake_png(self, label: str) -> str:
        payload = base64.b64encode(f"ppt:{label}".encode("utf-8")).decode("ascii")
        return f"data:image/png;base64,{payload}"

    def _ensure_slide(self, slide_index: int) -> Slide:
        while slide_index >= len(self.slides):
            self.slides.append(Slide())
        return self.slides[slide_index]

    def _preview(self, slide: Slide) -> str:
        parts = [shape.text or shape.name or shape.kind for shape in slide.shapes]
        return " | ".join(part for part in parts if part) or "(blank slide)"

    def get_presentation_overview(self, thumbnailWidth: float | None = None) -> ToolResult:  # noqa: N803
        del thumbnailWidth
        lines = ["Presentation Overview", "=" * 40, f"Total slides: {len(self.slides)}", ""]
        for index, slide in enumerate(self.slides, start=1):
            lines.append(f"Slide {index}: {self._preview(slide)}")
        return self._ok(
            {
                "text": "\n".join(lines),
                "slides": [
                    {"slideNumber": index, "image": self._fake_png(f"slide-{index}")}
                    for index in range(1, len(self.slides) + 1)
                ],
            }
        )

    def get_presentation_content(
        self,
        slideIndex: float | None = None,  # noqa: N803
        startIndex: float | None = None,  # noqa: N803
        endIndex: float | None = None,  # noqa: N803
    ) -> ToolResult:
        if not self.slides:
            return self._ok("Presentation has no slides.")

        if slideIndex is not None:
            start = end = int(slideIndex)
        elif startIndex is not None and endIndex is not None:
            start = int(startIndex)
            end = int(endIndex)
        else:
            start = 0
            end = len(self.slides) - 1

        if start < 0 or end >= len(self.slides) or start > end:
            return self._error("Requested slide range is invalid.")

        lines = []
        for idx in range(start, end + 1):
            lines.append(f"Slide {idx + 1}: {self._preview(self.slides[idx])}")
        return self._ok("\n".join(lines))

    def set_presentation_content(self, slideIndex: float, text: str) -> ToolResult:  # noqa: N803
        index = int(slideIndex)
        if index < 0:
            return self._error("slideIndex must be non-negative.")
        slide = self._ensure_slide(index)
        slide.shapes.append(SlideShape(kind="text", text=text, name=f"Text {len(slide.shapes)}"))
        return self._ok(f'Added text box to slide {index + 1}.')

    def add_geometric_shape(
        self,
        slideIndex: float,  # noqa: N803
        shapeType: str,  # noqa: N803
        left: float,
        top: float,
        width: float,
        height: float,
        fillColor: str | None = None,  # noqa: N803
        name: str | None = None,
    ) -> ToolResult:
        index = int(slideIndex)
        if index < 0 or index >= len(self.slides):
            return self._error("slideIndex is out of range.")
        self.slides[index].shapes.append(
            SlideShape(
                kind="shape",
                text=name or shapeType,
                name=name or shapeType,
                meta={
                    "shapeType": shapeType,
                    "left": left,
                    "top": top,
                    "width": width,
                    "height": height,
                    "fillColor": fillColor,
                },
            )
        )
        return self._ok(f'Added {shapeType} shape to slide {index + 1}.')

    def add_slide_from_code(self, code: str, replaceSlideIndex: float | None = None) -> ToolResult:  # noqa: N803
        shapes = self._parse_slide_code(code)
        slide = Slide(shapes=shapes)

        if replaceSlideIndex is None:
            self.slides.append(slide)
            return self._ok("Successfully added new slide to the presentation.")

        index = int(replaceSlideIndex)
        if index < 0 or index >= len(self.slides):
            return self._error("replaceSlideIndex is out of range.")
        self.slides[index] = slide
        return self._ok(f"Successfully replaced slide {index + 1}.")

    def fetch_image_as_base64(self, url: str) -> ToolResult:
        payload = base64.b64encode(url.encode("utf-8")).decode("ascii")
        return self._ok(f"data:image/png;base64,{payload}")

    def _parse_slide_code(self, code: str) -> list[SlideShape]:
        shapes: list[SlideShape] = []

        for text in re.findall(r"addText\(\s*['\"]([^'\"]+)['\"]", code):
            shapes.append(SlideShape(kind="text", text=text))

        for shape_type in re.findall(r"addShape\(\s*['\"]([^'\"]+)['\"]", code):
            shapes.append(SlideShape(kind="shape", text=shape_type, name=shape_type))

        if "addImage(" in code:
            shapes.append(SlideShape(kind="image", text="Image"))

        if "addChart(" in code:
            title_match = re.search(r"title\s*:\s*['\"]([^'\"]+)['\"]", code)
            chart_title = title_match.group(1) if title_match else "Chart"
            shapes.append(SlideShape(kind="chart", text=chart_title))

        if "addTable(" in code:
            shapes.append(SlideShape(kind="table", text="Table"))

        if not shapes:
            shapes.append(SlideShape(kind="code", text="Generated slide"))

        return shapes
