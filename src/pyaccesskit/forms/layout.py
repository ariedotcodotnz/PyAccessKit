"""Form layout engine: turns a :class:`FormSpec` into exact control geometry.

The engine is pure Python — it never talks to Access — so layouts can be unit-tested and previewed. Two
automatic layouts mirror Access's own control layouts:

* **stacked** (single/split forms): one control per row, attached label on the left;
* **tabular** (continuous/datasheet forms): labels in the form header, one row of controls in the detail.

Controls with an explicit ``at=`` keep their position; their attached label goes to their left.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from pyaccesskit.enums import FormView, LayoutKind, Section
from pyaccesskit.errors import SpecError
from pyaccesskit.forms.controls import ButtonSpec, CheckBoxSpec, ControlSpec, LabelSpec
from pyaccesskit.forms.spec import FormSpec, label_name_for
from pyaccesskit.forms.vba import EventBinding
from pyaccesskit.units import Length, cm, inch

__all__ = [
    "DEFAULT_METRICS",
    "MAX_FORM_EXTENT",
    "LayoutMetrics",
    "Rect",
    "ResolvedControl",
    "ResolvedForm",
    "ResolvedLabel",
    "layout_form",
]

MAX_FORM_EXTENT = inch(22)
"""Access limit for form width and for each section's height."""


@dataclass(frozen=True)
class LayoutMetrics:
    """Spacing and default sizes used by the automatic layouts."""

    margin: Length = field(default_factory=lambda: cm(0.5))
    label_width: Length = field(default_factory=lambda: cm(3.5))
    label_gap: Length = field(default_factory=lambda: cm(0.25))
    control_width: Length = field(default_factory=lambda: cm(6))
    row_height: Length = field(default_factory=lambda: cm(0.6))
    row_gap: Length = field(default_factory=lambda: cm(0.2))
    checkbox_size: Length = field(default_factory=lambda: cm(0.45))
    button_width: Length = field(default_factory=lambda: cm(3))
    button_height: Length = field(default_factory=lambda: cm(0.8))
    column_gap: Length = field(default_factory=lambda: cm(0.25))
    tabular_column_width: Length = field(default_factory=lambda: cm(3.5))


DEFAULT_METRICS = LayoutMetrics()


@dataclass(frozen=True)
class Rect:
    """A rectangle in twips-backed lengths."""

    left: Length
    top: Length
    width: Length
    height: Length

    @property
    def right(self) -> Length:
        """Right edge."""
        return self.left + self.width

    @property
    def bottom(self) -> Length:
        """Bottom edge."""
        return self.top + self.height

    def overlaps(self, other: Rect) -> bool:
        """Whether the two rectangles share any area."""
        return (
            self.left < other.right
            and other.left < self.right
            and self.top < other.bottom
            and other.top < self.bottom
        )


@dataclass(frozen=True)
class ResolvedLabel:
    """An attached label with its final geometry."""

    name: str
    caption: str
    section: Section
    rect: Rect


@dataclass(frozen=True)
class ResolvedControl:
    """A control with its final name, section and geometry."""

    spec: ControlSpec
    name: str
    section: Section
    rect: Rect
    label: ResolvedLabel | None = None


@dataclass(frozen=True)
class ResolvedForm:
    """Everything the Access materializer needs to build the form."""

    spec: FormSpec
    width: Length
    detail_height: Length
    header_height: Length | None
    """``None`` means the form has no header/footer sections."""
    footer_height: Length | None
    controls: tuple[ResolvedControl, ...]
    events: tuple[EventBinding, ...]
    module_text: str | None

    @property
    def has_header(self) -> bool:
        """Whether the form needs header/footer sections."""
        return self.header_height is not None


def _default_size(
    control: ControlSpec, m: LayoutMetrics, *, tabular: bool
) -> tuple[Length, Length]:
    if isinstance(control, CheckBoxSpec):
        width, height = m.checkbox_size, m.checkbox_size
    elif isinstance(control, ButtonSpec):
        width, height = m.button_width, m.button_height
    elif isinstance(control, LabelSpec):
        width = m.tabular_column_width if tabular else m.label_width + m.label_gap + m.control_width
        height = m.row_height
    else:
        width, height = (m.tabular_column_width if tabular else m.control_width), m.row_height
    return control.width or width, control.height or height


def _effective_layout(spec: FormSpec) -> LayoutKind:
    if spec.layout is not LayoutKind.AUTO:
        return spec.layout
    if spec.default_view in (FormView.CONTINUOUS, FormView.DATASHEET):
        return LayoutKind.TABULAR
    return LayoutKind.STACKED


def layout_form(spec: FormSpec, metrics: LayoutMetrics = DEFAULT_METRICS) -> ResolvedForm:
    """Compute the geometry of every control of ``spec``.

    Raises:
        SpecError: If a control cannot be placed or the form exceeds Access's size limits.
    """
    m = metrics
    kind = _effective_layout(spec)
    tabular = kind is LayoutKind.TABULAR
    controls = spec.resolved_controls()
    cursors: dict[Section, Length] = dict.fromkeys(Section, m.margin)
    tab_x = m.margin
    placed: list[ResolvedControl] = []
    label_x = m.margin
    control_x = m.margin + m.label_width + m.label_gap

    for control in controls:
        name = control.name or ""
        width, height = _default_size(control, m, tabular=tabular)
        caption = control.attached_label
        section = control.section

        if control.at is not None:
            left, top = control.at
            rect = Rect(left, top, width, height)
            label = None
            if caption is not None:
                label_left = left - m.label_gap - m.label_width
                if label_left.twips < 0:
                    raise SpecError(
                        f"no room for the label of {name!r} left of its position; move it right, "
                        "pass label=False, or add a free-standing label"
                    )
                label = ResolvedLabel(
                    label_name_for(name),
                    caption,
                    section,
                    Rect(label_left, top, m.label_width, m.row_height),
                )
            placed.append(ResolvedControl(control, name, section, rect, label))
            continue

        if kind is LayoutKind.NONE:
            raise SpecError(f"layout='none' requires at=(left, top) for control {name!r}")

        if tabular and section is Section.DETAIL:
            # A column is as wide as the wider of its header label and its control; narrow controls
            # (check boxes) are centred in their column.
            column_width = (
                width if isinstance(control, ButtonSpec) else max(width, m.tabular_column_width)
            )
            header_label = None
            if caption is not None or isinstance(control, LabelSpec):
                label_rect = Rect(tab_x, m.margin, column_width, m.row_height)
                if isinstance(control, LabelSpec):
                    placed.append(ResolvedControl(control, name, Section.HEADER, label_rect))
                    tab_x = tab_x + column_width + m.column_gap
                    continue
                header_label = ResolvedLabel(
                    label_name_for(name), caption or name, Section.HEADER, label_rect
                )
            left = tab_x + Length((column_width.twips - width.twips) // 2)
            placed.append(
                ResolvedControl(
                    control, name, section, Rect(left, m.row_gap, width, height), header_label
                )
            )
            tab_x = tab_x + column_width + m.column_gap
            continue

        # stacked (or tabular header/footer controls, which stack within their section)
        top = cursors[section]
        row_height = max(height, m.row_height)
        if isinstance(control, LabelSpec):
            rect = Rect(label_x, top, width, height)
            placed.append(ResolvedControl(control, name, section, rect))
        else:
            offset = (
                Length((row_height.twips - height.twips) // 2)
                if isinstance(control, CheckBoxSpec)
                else Length(0)
            )
            rect = Rect(control_x, top + offset, width, height)
            label = None
            if caption is not None:
                label = ResolvedLabel(
                    label_name_for(name),
                    caption,
                    section,
                    Rect(label_x, top, m.label_width, m.row_height),
                )
            placed.append(ResolvedControl(control, name, section, rect, label))
        cursors[section] = top + row_height + m.row_gap

    def extent(section: Section) -> Length | None:
        rects = [c.rect for c in placed if c.section is section]
        rects += [
            c.label.rect for c in placed if c.label is not None and c.label.section is section
        ]
        if not rects:
            return None
        return max(r.bottom for r in rects) + (
            m.row_gap if tabular and section is Section.DETAIL else m.margin
        )

    all_rects = [c.rect for c in placed] + [c.label.rect for c in placed if c.label is not None]
    width = max([r.right for r in all_rects], default=Length(0)) + m.margin
    if spec.width is not None and spec.width > width:
        width = spec.width

    detail = extent(Section.DETAIL) or m.margin * 2
    header_needed = any(
        c.section is not Section.DETAIL
        or (c.label is not None and c.label.section is not Section.DETAIL)
        for c in placed
    )
    has_header = spec.header if spec.header is not None else header_needed
    header_height: Length | None = None
    footer_height: Length | None = None
    if has_header:
        header_height = extent(Section.HEADER) or Length(0)
        footer_height = extent(Section.FOOTER) or Length(0)

    for label, value in (
        ("form width", width),
        ("detail height", detail),
        ("header height", header_height),
        ("footer height", footer_height),
    ):
        if value is not None and value > MAX_FORM_EXTENT:
            raise SpecError(
                f"{label} {value.format('cm')} exceeds Access's limit of {MAX_FORM_EXTENT.format('cm')}"
            )

    return ResolvedForm(
        spec=spec,
        width=width,
        detail_height=detail,
        header_height=header_height,
        footer_height=footer_height,
        controls=tuple(placed),
        events=spec.event_bindings(),
        module_text=spec.module_text(),
    )
