"""
io_schema.py — JSON contract for `extract_so_blocks_from_route` output.

This is the data-interchange surface between cad-agent-mcp (the producer,
C# plugin that reads .dwg via AutoCAD API) and pipe_engine (the consumer,
Python deterministic engine for ВОР/spec calculations).

Design principle: **thin extract**. The producer emits raw dwg facts only
(polylines + block instances + layer names + attribute strings). No
engineering interpretation — that is `pipe_engine.interpret`'s job.

Versioning: ``schema_version`` field on the top-level document. Bump on
any breaking change; loader rejects unknown versions.

Layout — see ``tests/fixtures/demo_extract.v1.json`` for a
hand-crafted example matching the structure below.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any


SCHEMA_VERSION = "1"


# ============================================================================
# Records
# ============================================================================

@dataclass(frozen=True)
class SourceMeta:
    """Provenance: which dwg was extracted, when, by which version."""
    dwg_path: str
    extracted_at_utc: str
    extractor_version: str
    insunits_raw: int                # what the dwg metadata says
    drawing_units_assumed: str       # what we actually treat coords as ("m", "mm", ...)
    extents_min: tuple[float, float, float]
    extents_max: tuple[float, float, float]

    @classmethod
    def from_dict(cls, d: dict) -> "SourceMeta":
        return cls(
            dwg_path=d["dwg_path"],
            extracted_at_utc=d["extracted_at_utc"],
            extractor_version=d["extractor_version"],
            insunits_raw=int(d["insunits_raw"]),
            drawing_units_assumed=d["drawing_units_assumed"],
            extents_min=tuple(d["extents_min"]),
            extents_max=tuple(d["extents_max"]),
        )


@dataclass(frozen=True)
class PolylineRecord:
    """One polyline from model space."""
    handle: str                      # hex DXF handle, unique within dwg
    layer: str                       # layer name (cp1251 string from dwg)
    length: float                    # 3D length in drawing units
    start: tuple[float, float, float]
    end: tuple[float, float, float]
    is_closed: bool
    vertex_count: int

    @classmethod
    def from_dict(cls, d: dict) -> "PolylineRecord":
        return cls(
            handle=d["handle"],
            layer=d["layer"],
            length=float(d["length"]),
            start=tuple(d["start"]),
            end=tuple(d["end"]),
            is_closed=bool(d["is_closed"]),
            vertex_count=int(d["vertex_count"]),
        )


@dataclass(frozen=True)
class BlockRecord:
    """One block instance (INSERT) from model space."""
    handle: str
    block_name_raw: str              # may be anonymous "*U37" for dynamic blocks
    block_name_effective: str        # resolved name, e.g. "СО_Отвод90_ПЭ+"
    layer: str
    position: tuple[float, float, float]
    rotation_rad: float
    scale: tuple[float, float, float]
    attributes: dict[str, str]       # ATTRIB tag → text value (raw, unparsed)

    @classmethod
    def from_dict(cls, d: dict) -> "BlockRecord":
        return cls(
            handle=d["handle"],
            block_name_raw=d["block_name_raw"],
            block_name_effective=d["block_name_effective"],
            layer=d["layer"],
            position=tuple(d["position"]),
            rotation_rad=float(d["rotation_rad"]),
            scale=tuple(d["scale"]),
            attributes=dict(d.get("attributes", {})),
        )


@dataclass(frozen=True)
class BlockFilterStats:
    """How blocks were filtered during extraction."""
    include_pattern: str
    include_anonymous_resolved: bool
    matched_count: int
    skipped_count: int

    @classmethod
    def from_dict(cls, d: dict) -> "BlockFilterStats":
        return cls(
            include_pattern=d["include_pattern"],
            include_anonymous_resolved=bool(d["include_anonymous_resolved"]),
            matched_count=int(d["matched_count"]),
            skipped_count=int(d["skipped_count"]),
        )


@dataclass(frozen=True)
class ExtractDocument:
    """Top-level dwg → JSON extract artifact."""
    schema_version: str
    source: SourceMeta
    polylines: list[PolylineRecord]
    blocks: list[BlockRecord]
    block_filter: BlockFilterStats

    @classmethod
    def from_dict(cls, d: dict) -> "ExtractDocument":
        v = d.get("schema_version")
        if v != SCHEMA_VERSION:
            raise ValueError(
                f"schema_version mismatch: extract has {v!r}, "
                f"pipe_engine expects {SCHEMA_VERSION!r}"
            )
        return cls(
            schema_version=v,
            source=SourceMeta.from_dict(d["source"]),
            polylines=[PolylineRecord.from_dict(p) for p in d.get("polylines", [])],
            blocks=[BlockRecord.from_dict(b) for b in d.get("blocks", [])],
            block_filter=BlockFilterStats.from_dict(d["block_filter"]),
        )

    @classmethod
    def from_json_file(cls, path: str | Path) -> "ExtractDocument":
        with Path(path).open("r", encoding="utf-8") as fh:
            return cls.from_dict(json.load(fh))
