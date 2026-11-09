"""The op language: a closed set of typed operations against a deck.

POC implements the subset needed for translate + create + repair. The full design
(~20 ops) extends this union; the executor dispatch table is the only other place
that changes.
"""
from __future__ import annotations

from typing import Annotated, Literal, Union

from pydantic import BaseModel, ConfigDict, Field, TypeAdapter


class _Op(BaseModel):
    model_config = ConfigDict(extra="forbid")


class Expect(BaseModel):
    model_config = ConfigDict(extra="forbid")
    text_sha: str


class SetParagraphText(_Op):
    """Rewrite a paragraph from tagged text; runs keep their own formatting and
    may be reordered or dropped (needed for translation - word order changes)."""
    op: Literal["set_paragraph_text"]
    target: str
    expect: Expect
    value: str


class SetRunText(_Op):
    op: Literal["set_run_text"]
    target: str
    expect: Expect
    value: str


class SetFontSize(_Op):
    """Explicit size on every run of a shape. Deterministic overflow repair."""
    op: Literal["set_font_size"]
    target: str
    size_pt: float = Field(ge=6, le=96)


class TextContent(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["text"]
    text: str


class BulletsContent(BaseModel):
    model_config = ConfigDict(extra="forbid")
    kind: Literal["bullets"]
    items: list[str] = Field(min_length=1)


Content = Annotated[Union[TextContent, BulletsContent], Field(discriminator="kind")]


class AddSlide(_Op):
    op: Literal["add_slide"]
    layout: str                     # layout id from the template catalog, e.g. "M0L1"
    after: str | None = None        # slide ref; None = append
    fill: dict[int, Content]        # placeholder idx -> content (compiler resolved roles)
    notes: str | None = None
    drop_empty_placeholders: bool = True


class DeleteSlide(_Op):
    op: Literal["delete_slide"]
    target: str


class DeleteShape(_Op):
    op: Literal["delete_shape"]
    target: str


class SetGeometry(_Op):
    """Position/size in EMU. Emitted by the compiler, never by the model."""
    op: Literal["set_geometry"]
    target: str
    x: int = Field(ge=0)
    y: int = Field(ge=0)
    w: int = Field(gt=0)
    h: int = Field(ge=0)


Op = Annotated[
    Union[SetParagraphText, SetRunText, SetFontSize, AddSlide, DeleteSlide, DeleteShape, SetGeometry],
    Field(discriminator="op"),
]
OpList = TypeAdapter(list[Op])
