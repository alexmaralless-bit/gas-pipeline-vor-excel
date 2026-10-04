"""
node_library.py — параметризованные шаблоны типовых узлов.

Идея: каждый «узел» в проекте (КТ, ЗУ кран, Узел А и т.д.) описан
один раз как параметризованный шаблон (YAML), из которого pipe_engine
автоматически генерирует строки ВР для каждого экземпляра соответствующего
СО_* блока в dwg.

Источник истины для compositions — реальный лист «ВР» эталонной книги,
где для каждого типа узла уже расписан полный список компонентов с
qty/unit (а не сами typical-dwg, которые показывают только geometry).

Шаблоны YAML живут в ``pipe_engine/data/node_templates/*.yaml``. Каждый
YAML описывает:

* ``id`` — стабильный id (используется в коде)
* ``name`` — человекочитаемое название для ВР section header
* ``detected_by_block`` — какой СО_* блок в dwg триггерит этот шаблон
* ``block_aliases`` — дополнительные реальные имена блоков, которые должны
  матчиться на тот же шаблон
* ``host_dn_source`` — откуда взять Ду основного газопровода
    (``"layer"`` = из имени слоя ближайшей трасс-полилинии,
     ``"attribute:DN"`` = из атрибута блока,
     ``"manual"`` = задаётся снаружи)
* ``variants_by_dn`` — таблица типоразмеров (опционально)
* ``components`` — список компонентов для ВР (каждый = одна строка ВР)
* ``work_label`` — строка-«заголовок группы» в ВР (опционально)

См. ``data/node_templates/kt_po_trasse.yaml`` как канонический пример.
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal, Optional


# ============================================================================
# Dataclasses
# ============================================================================

@dataclass(frozen=True)
class NodeVariant:
    """Один типоразмер узла (одна строка таблицы вариантов на чертеже)."""
    host_dn: int                  # Ду основного газопровода, мм
    suffix: str                   # суффикс обозначения, e.g. "00", "01", "10"
    L_mm: Optional[int] = None    # длина (для узлов где L зависит от DN)
    extra: dict = field(default_factory=dict)  # любые доп-параметры


@dataclass(frozen=True)
class NodeComponent:
    """Один компонент узла → одна строка ВР.

    ``qty`` может быть числом (literal) или строкой-формулой Excel.
    ``catalog_match`` — словарь критериев для VLOOKUP в Каталог при
    заполнении column «Наименование» (когда будет реализован Спец-лист).
    ``dn_hint`` — если задано, перебивает host_dn узла при VLOOKUP/spec
    lookup для этого компонента. Нужно для multi-DN узлов (например
    Узел А имеет основной кран DN63 + футляр Ду50 + контрольную трубку
    Ду20). Без dn_hint все компоненты использовали бы host_dn=63
    и VLOOKUP возвращал бы неверные позиции каталога.
    """
    tag: str                              # имя/тег компонента (для ВР column D)
    qty: float | int | str = 1            # количество — литерал или Excel formula
    unit: str = "шт"
    catalog_match: dict = field(default_factory=dict)
    kind: Literal["material", "work"] = "material"
    dn_hint: Optional[int] = None         # override host_dn для catalog lookup


@dataclass
class NodeTemplate:
    """Параметризованный шаблон одного типа узла.

    ``triggered_by`` определяет КАК узел детектируется в проекте:
      * ``"block"`` (default): по СО_* блоку в dwg (поле detected_by_block)
      * ``"polyline_layer"``: по полилиниям на слое-маркере
            (поле polyline_layer_pattern — regex, e.g. r"^!!ГСН.*_Вводы@\\d+$")
            count = number of polylines matching the pattern
      * ``"manual"``: количество задаётся снаружи (CLI flag или dict),
            поле manual_count_key — имя ключа в dict (e.g. "grpsh")

    Для polyline-trigger host_dn extract'ится из имени слоя через regex
    group (e.g. layer "!!ГСН3_Г1_ГП_Вводы@63" → host_dn=63).
    """
    id: str
    name: str
    components: list[NodeComponent]
    triggered_by: Literal["block", "polyline_layer", "manual"] = "block"
    # Trigger-specific fields:
    detected_by_block: str = ""             # for triggered_by="block"
    block_aliases: list[str] = field(default_factory=list)
    polyline_layer_pattern: str = ""        # regex for triggered_by="polyline_layer"
    manual_count_key: str = ""              # CLI key for triggered_by="manual"
    # Common fields:
    variants_by_dn: list[NodeVariant] = field(default_factory=list)
    host_dn_source: str = "layer"   # "layer" | "attribute:<TAG>" | "manual" | "polyline_layer"
    work_label: Optional[str] = None
    typical_drawing_ref: Optional[str] = None   # path / human-readable reference
    notes: str = ""

    def pick_variant(self, host_dn_mm: Optional[int]) -> Optional[NodeVariant]:
        """Найти подходящий типоразмер по DN основного газопровода.

        Если variants пустые — None (узел не зависит от DN).
        Если есть точное совпадение DN — возвращаем его.
        Иначе берём ближайший по DN снизу (или первый если все больше).
        """
        if not self.variants_by_dn:
            return None
        if host_dn_mm is None:
            return self.variants_by_dn[0]
        exact = [v for v in self.variants_by_dn if v.host_dn == host_dn_mm]
        if exact:
            return exact[0]
        # ближайший снизу
        smaller = [v for v in self.variants_by_dn if v.host_dn <= host_dn_mm]
        if smaller:
            return max(smaller, key=lambda v: v.host_dn)
        return self.variants_by_dn[0]


@dataclass
class InstantiatedNode:
    """Один конкретный экземпляр узла (= один блок СО_* в dwg)."""
    template_id: str
    template_name: str
    block_handle: str
    host_dn_mm: Optional[int]
    variant: Optional[NodeVariant]
    components: list[NodeComponent]
    work_label: Optional[str]


# ============================================================================
# Library
# ============================================================================

class NodeLibrary:
    """Коллекция NodeTemplate, загружается из директории YAML/JSON-файлов
    (или из встроенного pipe_engine/data/node_templates/)."""

    def __init__(self, templates: list[NodeTemplate]):
        self._templates = templates
        self._by_block: dict[str, NodeTemplate] = {}
        self._by_id: dict[str, NodeTemplate] = {}
        for t in templates:
            if t.triggered_by == "block" and t.detected_by_block:
                self._by_block[t.detected_by_block] = t
                for alias in t.block_aliases:
                    self._by_block[alias] = t
            self._by_id[t.id] = t

    def templates_by_trigger(self, trigger: str) -> list[NodeTemplate]:
        return [t for t in self._templates if t.triggered_by == trigger]

    @classmethod
    def load_from_directory(cls, yaml_dir: str | Path) -> "NodeLibrary":
        """Загрузить все *.yaml из директории.

        Если PyYAML не установлен — fallback на минимальный YAML-like
        парсер для нашего ограниченного формата. Для MVP мы поддерживаем
        более простой формат: каждый файл — JSON (быстро, без yaml dep).
        """
        yaml_dir = Path(yaml_dir)
        if not yaml_dir.is_dir():
            raise FileNotFoundError(f"node templates directory not found: {yaml_dir}")
        templates = []
        for f in sorted(yaml_dir.glob("*.json")):
            templates.append(cls._load_one_json(f))
        for f in sorted(yaml_dir.glob("*.yaml")):
            templates.append(cls._load_one_yaml(f))
        for f in sorted(yaml_dir.glob("*.yml")):
            templates.append(cls._load_one_yaml(f))
        return cls(templates)

    @classmethod
    def load_default(cls) -> "NodeLibrary":
        """Default templates shipped with pipe_engine."""
        default_dir = Path(__file__).resolve().parent / "data" / "node_templates"
        if not default_dir.is_dir():
            return cls([])
        return cls.load_from_directory(default_dir)

    @staticmethod
    def _load_one_json(path: Path) -> NodeTemplate:
        with path.open("r", encoding="utf-8") as fh:
            data = json.load(fh)
        return NodeLibrary._build_template(data, source=path.name)

    @staticmethod
    def _load_one_yaml(path: Path) -> NodeTemplate:
        try:
            import yaml  # type: ignore
        except ImportError:
            raise RuntimeError(
                f"YAML template {path.name} requires `pip install pyyaml`. "
                f"Or convert to .json — same schema works."
            )
        with path.open("r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh)
        return NodeLibrary._build_template(data, source=path.name)

    @staticmethod
    def _build_template(d: dict, source: str) -> NodeTemplate:
        try:
            components = [
                NodeComponent(
                    tag=c["tag"],
                    qty=c.get("qty", 1),
                    unit=c.get("unit", "шт"),
                    catalog_match=dict(c.get("catalog_match", {})),
                    kind=c.get("kind", "material"),
                    dn_hint=(int(c["dn_hint"]) if c.get("dn_hint") is not None else None),
                )
                for c in d.get("components", [])
            ]
            variants = [
                NodeVariant(
                    host_dn=int(v["host_dn"]),
                    suffix=str(v.get("suffix", "")),
                    L_mm=v.get("L_mm"),
                    extra={k: vv for k, vv in v.items()
                           if k not in ("host_dn", "suffix", "L_mm")},
                )
                for v in d.get("variants_by_dn", [])
            ]
            trigger = d.get("triggered_by", "block")
            # Backward compat: if detected_by_block is present without triggered_by,
            # implicitly trigger="block".
            return NodeTemplate(
                id=d["id"],
                name=d["name"],
                components=components,
                triggered_by=trigger,
                detected_by_block=d.get("detected_by_block", ""),
                block_aliases=list(d.get("block_aliases", [])),
                polyline_layer_pattern=d.get("polyline_layer_pattern", ""),
                manual_count_key=d.get("manual_count_key", ""),
                variants_by_dn=variants,
                host_dn_source=d.get("host_dn_source", "layer"),
                work_label=d.get("work_label"),
                typical_drawing_ref=d.get("typical_drawing_ref"),
                notes=d.get("notes", ""),
            )
        except KeyError as e:
            raise ValueError(f"template {source} missing required field: {e}") from e

    # ---- access ---------------------------------------------------------

    def __len__(self) -> int:
        return len(self._templates)

    @property
    def templates(self) -> list[NodeTemplate]:
        return list(self._templates)

    def find_by_block(self, block_name: str) -> Optional[NodeTemplate]:
        return self._by_block.get(block_name)

    def find_by_id(self, template_id: str) -> Optional[NodeTemplate]:
        return self._by_id.get(template_id)

    # ---- instantiation --------------------------------------------------

    def instantiate(self, block_name: str, *,
                    block_handle: str = "",
                    host_dn_mm: Optional[int] = None) -> Optional[InstantiatedNode]:
        """Создать экземпляр узла для блока. None если шаблона нет."""
        template = self.find_by_block(block_name)
        if template is None:
            return None
        variant = template.pick_variant(host_dn_mm)
        return InstantiatedNode(
            template_id=template.id,
            template_name=template.name,
            block_handle=block_handle,
            host_dn_mm=host_dn_mm,
            variant=variant,
            components=list(template.components),
            work_label=template.work_label,
        )
