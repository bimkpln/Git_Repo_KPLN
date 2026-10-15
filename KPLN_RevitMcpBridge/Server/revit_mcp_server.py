"""Codex <-> MCP stdio <-> loopback HTTP <-> KPLN_Loader / Revit Idling."""
from __future__ import annotations

from typing import Annotated, Literal
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field, StrictFloat, StrictInt, StrictStr, model_validator
from bridge_client import BridgeClient

mcp = FastMCP("KPLN_RevitMcpBridge", instructions=(
    "Работа с локальным Revit через KPLN_Loader. Сначала list_revit_sessions и get_revit_context. "
    "Для чтения/записи передавайте полученный document_id, для записи также expected_revision. "
    "set_revit_parameters и set_revit_family_type_parameters по умолчанию строят план; dry_run=False применяет его. После таймаута "
    "не повторяйте запись: get_revit_operation. Правила генерации семейств храните в KPLN_RevitRules, "
    "проверки кандидатов — в KPLN_JobCandidateModelCheckRules, оформления ОВ — в KPLN_OVModelDocumentationRules. "
    "Команды оформления по умолчанию выполняют создание и Commit с откатом: ID dry_run недействительны; обновите context перед записью. "
    "геометрия в мм относительно внутренних координат документа, Double-параметры в единицах Revit."
))
client = BridgeClient()
READ = ToolAnnotations(readOnlyHint=True, destructiveHint=False, idempotentHint=True, openWorldHint=False)
UI = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=True, openWorldHint=False)
POST_UI = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False)
EXPORT = ToolAnnotations(readOnlyHint=False, destructiveHint=False, idempotentHint=False, openWorldHint=False)
WRITE = ToolAnnotations(readOnlyHint=False, destructiveHint=True, idempotentHint=False, openWorldHint=False)
Limit = Annotated[int, Field(ge=1, le=200)]
Offset = Annotated[int, Field(ge=0, le=10000000)]
Ids = Annotated[list[str], Field(min_length=1, max_length=200)]
CategoryIds = Annotated[list[str], Field(min_length=1, max_length=64)]


def _validate_sheet_selection(ids):
    if not isinstance(ids, list) or not 1 <= len(ids) <= 200 or any(not isinstance(i, str) or not i.strip() for i in ids) or len(set(ids)) != len(ids):
        raise ValueError("Нужны от 1 до 200 непустых уникальных UniqueId листов.")


class BatchPrintSettings(BaseModel):
    model_config = ConfigDict(extra="forbid")
    printerName: Literal["PDFCreator"] = "PDFCreator"
    outputPDFFolder: str = "C:\\PDF_Print"
    pdfNameConstructor: str = "<Номер листа>_<Имя листа>.pdf"
    hiddenLineProcessing: Literal["VectorProcessing", "RasterProcessing"] = "VectorProcessing"
    colorsType: Literal["Color", "Monochrome", "GrayScale"] = "Color"
    rasterQuality: Literal["Low", "Medium", "High", "Presentation"] = "Medium"
    isPDFExport: Literal[True] = True
    isPrintToPaper: Literal[False] = False
    isMergePdfs: Literal[False] = False
    isUseOrientation: bool = False
    isRefreshSchedules: bool = True
    isExcludeBorders: Literal[False] = False
    isDWGExport: Literal[False] = False


@mcp.tool(annotations=EXPORT)
def print_revit_sheets_pdf(document_id: str, expected_revision: Annotated[int, Field(ge=0)], sheet_unique_ids: Ids, settings: BatchPrintSettings | None = None,
                           dry_run: bool = True, timeout_seconds: Annotated[int, Field(ge=5, le=60)] = 30, session_id: str | None = None) -> dict:
    """Только явно заданные sheet_unique_ids (1–200 листов) через KPLN Пакетную выдачу/PDFCreator без окон. dry_run строит план.

    Новая папка запуска внутри outputPDFFolder; модель не сохраняется, временные изменения откатываются.
    После outcome_pending читать get_revit_operation; при таймауте принтера не повторять печать автоматически.
    Нужны обновлённая KPLN_Publication и настроенный PDFCreator. Связи не печатаются.
    """
    _validate_sheet_selection(sheet_unique_ids)
    return client.command("print_selected_sheets_pdf", session_id, document_id=document_id, expected_revision=expected_revision, sheet_unique_ids=sheet_unique_ids,
                          settings=(settings or BatchPrintSettings()).model_dump(), dry_run=dry_run, timeout_seconds=timeout_seconds)


class ParameterUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid")
    element_unique_id: str
    parameter_id: str
    expected_value: StrictStr | StrictInt | StrictFloat | None = Field(description="Точное сырое value из get_revit_elements; null допустим для отсутствующего значения.")
    value: StrictStr | StrictInt | StrictFloat
    value_units: Literal["revit_internal"] | None = Field(default=None, description="Обязательно для Double. Длина: футы; угол: радианы. Не передавайте мм как сырые значения.")


class FamilyValueUpdate(BaseModel):
    model_config = ConfigDict(extra="forbid")
    parameter_id: str
    value: StrictStr | StrictInt | StrictFloat
    value_units: Literal["revit_internal"] | None = Field(default=None, description="Обязательно для Double.")


class FamilyValueSet(BaseModel):
    model_config = ConfigDict(extra="forbid")
    parameter_id: str
    expected_value: StrictStr | StrictInt | StrictFloat | None = Field(description="Точное сырое value из get_revit_family_document; null допустим для отсутствующего значения.")
    value: StrictStr | StrictInt | StrictFloat
    value_units: Literal["revit_internal"] | None = Field(default=None, description="Обязательно для Double.")


class FamilyTypeCreate(BaseModel):
    model_config = ConfigDict(extra="forbid")
    name: str
    values: Annotated[list[FamilyValueUpdate], Field(min_length=1, max_length=200)]


class FamilyTypeSet(BaseModel):
    model_config = ConfigDict(extra="forbid")
    name: str
    values: Annotated[list[FamilyValueSet], Field(min_length=1, max_length=200)]


@mcp.tool(annotations=READ)
def list_revit_sessions() -> dict:
    """Открытые экземпляры Revit 2020/2023/2024 с загруженным мостом; ключи доступа не выдаются."""
    return client.sessions()


@mcp.tool(annotations=READ)
def get_revit_bridge_health(session_id: str | None = None) -> dict:
    """Проверить транспорт без доступа к модели. Занятый Revit может отвечать health, но не выполнять команды."""
    return client.health(session_id)


@mcp.tool(annotations=READ)
def get_revit_operation(operation_id: str, session_id: str) -> dict:
    """Результат или состояние запроса после таймаута. Хранится минимум 15 минут до перезапуска моста."""
    return client.operation(operation_id, session_id)


@mcp.tool(annotations=READ)
def get_revit_documents(session_id: str | None = None) -> dict:
    """Список открытых документов и идентификатор активного. Чтение элементов выполняется в активном документе."""
    return client.command("get_documents", session_id)


@mcp.tool(annotations=READ)
def get_revit_context(session_id: str | None = None) -> dict:
    """Активный документ, document_id, revision, вид и единицы. Начальная точка работы с моделью."""
    return client.command("get_context", session_id)


@mcp.tool(annotations=READ)
def get_revit_selection(document_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Прочитать текущее выделение постранично; next_offset=null означает последнюю страницу."""
    return client.command("get_selection", session_id, document_id=document_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_categories(document_id: str, session_id: str | None = None) -> dict:
    """Категории активного документа с устойчивыми числовыми Id, независимо от языка Revit."""
    return client.command("get_categories", session_id, document_id=document_id)


@mcp.tool(annotations=READ)
def find_revit_elements(document_id: str, category_id: str | None = None, name_contains: str | None = None, element_types_only: bool = False, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Найти экземпляры либо типы по категории и части имени. Страницы сравнивайте при одной revision."""
    return client.command("find_elements", session_id, document_id=document_id, category_id=category_id, name_contains=name_contains, include_types=element_types_only, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_elements(document_id: str, unique_ids: Ids, session_id: str | None = None) -> dict:
    """Карточки элементов: параметры экземпляра/типа, расположение и габарит в мм. Габарит не является точной геометрией. Для параметров типа отдельно запросите элемент типа."""
    return client.command("get_elements", session_id, document_id=document_id, unique_ids=unique_ids)


@mcp.tool(annotations=READ)
def get_revit_views(document_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Виды активного документа без шаблонов, постранично."""
    return client.command("get_views", session_id, document_id=document_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_group_members(document_id: str, group_unique_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Состав экземпляра группы, тип, родитель и привязанные группы детализации. Только прямые члены; вложенные группы запрашиваются отдельно. Сохраняет member_index, страницы читать при одной revision."""
    return client.command("get_group_members", session_id, document_id=document_id, group_unique_id=group_unique_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_sheet_contents(document_id: str, sheet_unique_id: str, session_id: str | None = None) -> dict:
    """Размещённые на листе виды (тип, масштаб, шаблон), спецификации и основные надписи. Не доказывает фактическое отображение каждого элемента."""
    return client.command("get_sheet_contents", session_id, document_id=document_id, sheet_unique_id=sheet_unique_id)


@mcp.tool(annotations=READ)
def get_revit_view_elements(document_id: str, view_unique_id: str, category_id: str | None = None, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Потенциально видимые элементы вида с группами/рабочими наборами. Это не пиксельная видимость: обрезка и перекрытие могут дополнительно скрыть элементы. Для листа сначала получить размещённые виды."""
    return client.command("get_view_elements", session_id, document_id=document_id, view_unique_id=view_unique_id, category_id=category_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_view_visibility(document_id: str, view_unique_id: str, unique_ids: Ids, session_id: str | None = None) -> dict:
    """Для указанных элементов: потенциальная видимость, скрытие экземпляра/категории, рабочие наборы, фильтры и показ привязанных групп. Не заменяет визуальную проверку обрезки/перекрытия."""
    return client.command("get_view_visibility", session_id, document_id=document_id, view_unique_id=view_unique_id, unique_ids=unique_ids)


@mcp.tool(annotations=READ)
def get_revit_schedule_data(document_id: str, schedule_unique_id: str, section: Literal["body", "header"] = "body", offset: Offset = 0, limit: Annotated[int, Field(ge=1, le=100)] = 50, column_offset: Offset = 0, column_limit: Annotated[int, Field(ge=1, le=50)] = 50, session_id: str | None = None) -> dict:
    """Поля, фильтры, сортировка/группирование, итоги и ячейки спецификации. Обойти обе пагинации строк/столбцов при одной revision. Ячейки — отображаемый текст; формулы и вложенное определение не читаются. Оценка правильности — в слое правил."""
    return client.command("get_schedule_data", session_id, document_id=document_id, schedule_unique_id=schedule_unique_id, section=section, offset=offset, limit=limit, column_offset=column_offset, column_limit=column_limit)


@mcp.tool(annotations=READ)
def check_revit_intersections(document_id: str, max_pairs: Annotated[int, Field(ge=1, le=5000)] = 1000, time_budget_seconds: Annotated[int, Field(ge=1, le=60)] = 20, session_id: str | None = None) -> dict:
    """API-коллизии всех 3D-элементов текущего документа без фильтров категорий/видов. Мост >=1.0.11, contract_version=2. Solid-геометрия, включая члены групп и скрытые элементы; Mesh/поверхности, связи, закрытые РН и ошибки явно дают incomplete. Линии без объёма перечислены отдельно. Нет лимита 2000 элементов. Лимиты времени/пар никогда не дают ложного успеха. Только complete=true/status=no_intersections означает отсутствие найденных пересечений в пределах механизма Revit API. После таймаута читать get_revit_operation."""
    return client.command("check_intersections", session_id, document_id=document_id, max_pairs=max_pairs, time_budget_seconds=time_budget_seconds)


@mcp.tool(annotations=READ)
def get_revit_links(document_id: str, session_id: str | None = None) -> dict:
    """Экземпляры связей Revit, состояние загрузки и матрицы переноса из координат связи в хост. Не редактирует связи."""
    return client.command("get_links", session_id, document_id=document_id)


@mcp.tool(annotations=EXPORT)
def export_revit_sheets_pdf(document_id: str, expected_revision: Annotated[int, Field(ge=0)], sheet_unique_ids: Ids, dry_run: bool = True, session_id: str | None = None) -> dict:
    """Только явно заданные sheet_unique_ids (1–200 листов) в PDF штатным экспортом Revit 2023/2024. Сначала dry_run=True: перечень листов и пропусков. dry_run=False создаёт отдельный PDF каждого печатаемого листа и manifest.json в новом локальном каталоге моста; возвращает абсолютные пути, SHA256 и соответствие листам. Размер бумаги по листу, 100%, цвет. Документ не сохраняется, принтеры/настройки печати не меняются. Пропуски/ошибки дают complete=false. После таймаута читать get_revit_operation, НЕ повторять экспорт. Для визуального контроля прочитать реальные PDF, проверить страницы и размеры. Revit 2020 не поддерживается."""
    _validate_sheet_selection(sheet_unique_ids)
    return client.command("export_selected_sheets_pdf", session_id, document_id=document_id, expected_revision=expected_revision, sheet_unique_ids=sheet_unique_ids, dry_run=dry_run)


@mcp.tool(annotations=READ)
def get_revit_families(document_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Загруженные семейства Family в проекте (не системные семейства), постранично."""
    return client.command("get_families", session_id, document_id=document_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_family_types(document_id: str, family_unique_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Типоразмеры указанного загруженного семейства."""
    return client.command("get_family_types", session_id, document_id=document_id, family_unique_id=family_unique_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_family_document(document_id: str, offset: Offset = 0, limit: Annotated[int, Field(ge=1, le=100)] = 50, session_id: str | None = None) -> dict:
    """В редакторе семейства: параметры, формулы, значения типов и associated_parameters — прямые привязки к параметрам элементов/вложенных семейств. На параметр до 200 привязок; count/truncated показывают полноту. Пустой список не доказывает отсутствие использования в формулах или размерных метках. Double в единицах Revit. Только чтение."""
    return client.command("get_family_document", session_id, document_id=document_id, offset=offset, limit=limit)


@mcp.tool(annotations=UI)
def select_revit_elements(document_id: str, unique_ids: Ids, show: bool = False, session_id: str | None = None) -> dict:
    """Выделить элементы в Revit, при show=True также приблизить вид. Модель не редактируется."""
    return client.command("select_elements", session_id, document_id=document_id, unique_ids=unique_ids, show=show)


@mcp.tool(annotations=WRITE)
def set_revit_parameters(document_id: str, expected_revision: Annotated[int, Field(ge=0)], updates: Annotated[list[ParameterUpdate], Field(min_length=1, max_length=200)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Пакет параметров проекта по UniqueId и parameter_id: план по умолчанию, одна атомарная транзакция при dry_run=False. expected_value обязателен; Double требует value_units=revit_internal. Любое предупреждение Revit откатывает пакет. Параметры типа затрагивают все экземпляры типа. Документ не сохраняется."""
    return client.command("set_parameters", session_id, document_id=document_id, expected_revision=expected_revision, updates=[u.model_dump(exclude_none=False) for u in updates], dry_run=dry_run)


@mcp.tool(annotations=WRITE)
def create_revit_family_types(document_id: str, expected_revision: Annotated[int, Field(ge=0)], base_type_name: str, types: Annotated[list[FamilyTypeCreate], Field(min_length=1, max_length=20)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """В открытом редакторе семейства: атомарно создать типоразмеры из базового типа и задать FamilyManager-параметры. Имена должны отсутствовать; Double — только в единицах Revit. Сначала dry_run=True. Файл не сохраняется."""
    return client.command("create_family_types", session_id, document_id=document_id, expected_revision=expected_revision, base_type_name=base_type_name,
                          types=[t.model_dump(exclude_none=False) for t in types], dry_run=dry_run)


@mcp.tool(annotations=WRITE)
def set_revit_family_type_parameters(document_id: str, expected_revision: Annotated[int, Field(ge=0)], types: Annotated[list[FamilyTypeSet], Field(min_length=1, max_length=20)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """В открытом редакторе семейства: атомарно изменить параметры существующих типоразмеров FamilyManager. expected_value обязателен; Double — только в единицах Revit. Любое предупреждение откатывает весь пакет. Сначала dry_run=True. Файл не сохраняется."""
    return client.command("set_family_type_parameters", session_id, document_id=document_id, expected_revision=expected_revision,
                          types=[t.model_dump(exclude_none=False) for t in types], dry_run=dry_run)


Point3 = tuple[Annotated[float, Field(allow_inf_nan=False)], Annotated[float, Field(allow_inf_nan=False)], Annotated[float, Field(allow_inf_nan=False)]]


class DocumentationItem(BaseModel):
    model_config = ConfigDict(extra="forbid")


class SheetCreate(DocumentationItem):
    title_block_type_unique_id: str
    number: Annotated[str, Field(min_length=1)]
    name: Annotated[str, Field(min_length=1)]


class ViewPlacement(DocumentationItem):
    sheet_unique_id: str
    view_unique_id: str
    point_mm: Point3 | None = Field(default=None, description="Центр окна вида в координатах листа, мм, Z=0. Для спецификации — точка вставки.")
    reference_viewport_unique_id: str | None = Field(default=None, description="Окно вида на листе-образце в том же документе: взять центр, тип, поворот и положение заголовка.")

    @model_validator(mode="after")
    def placement_source(self):
        if (self.point_mm is None) == (self.reference_viewport_unique_id is None):
            raise ValueError("Укажите ровно один источник: point_mm или reference_viewport_unique_id")
        if self.point_mm is not None and self.point_mm[2] != 0:
            raise ValueError("Координаты листа требуют Z=0")
        return self


class SheetItemMove(DocumentationItem):
    sheet_unique_id: Annotated[str, Field(min_length=1)]
    instance_unique_id: Annotated[str, Field(min_length=1)]
    expected_point_mm: Point3 = Field(description="placement_point_mm из get_revit_sheet_layout; Z=0.")
    point_mm: Point3 = Field(description="Новая точка вставки спецификации или центр окна вида, мм листа, Z=0.")

    @model_validator(mode="after")
    def planar_points(self):
        if self.expected_point_mm[2] != 0 or self.point_mm[2] != 0:
            raise ValueError("Координаты листа требуют Z=0")
        return self


class ViewportLabelUpdate(DocumentationItem):
    sheet_unique_id: Annotated[str, Field(min_length=1)]
    viewport_unique_id: Annotated[str, Field(min_length=1)]
    expected_label_offset_mm: Point3 = Field(description="Прочитанный label_offset_mm из get_revit_sheet_layout; Z=0.")
    label_offset_mm: Point3 = Field(description="Смещение API LabelOffset в мм бумаги от нижнего левого угла окна при Rotation=None; не абсолютная точка текста. Z=0.")
    expected_label_line_length_mm: Annotated[float, Field(allow_inf_nan=False)] | None = None
    label_line_length_mm: Annotated[float, Field(allow_inf_nan=False)] | None = None

    @model_validator(mode="after")
    def validate_label(self):
        if self.expected_label_offset_mm[2] != 0 or self.label_offset_mm[2] != 0:
            raise ValueError("Смещение заголовка требует Z=0")
        if (self.label_line_length_mm is None) != (self.expected_label_line_length_mm is None):
            raise ValueError("Для длины линии нужны новое и ожидаемое значения")
        return self


class SheetReservedRegion(DocumentationItem):
    name: Annotated[str, Field(min_length=1)]
    min_mm: Point3
    max_mm: Point3

    @model_validator(mode="after")
    def rectangle(self):
        if self.min_mm[2] != 0 or self.max_mm[2] != 0 or any(self.min_mm[i] >= self.max_mm[i] for i in (0, 1)):
            raise ValueError("Нужен непустой прямоугольник листа с Z=0")
        return self


class TagCreate(DocumentationItem):
    view_unique_id: str
    tag_type_unique_id: str
    element_unique_id: str
    head_mm: Point3
    kind: Literal["independent", "room", "space"] = "independent"
    leader: bool = False
    orientation: Literal["Horizontal", "Vertical"] = "Horizontal"
    link_instance_unique_id: str | None = None


class LinkedRoomTagCreate(DocumentationItem):
    view_unique_id: Annotated[str, Field(min_length=1)]
    link_instance_unique_id: Annotated[str, Field(min_length=1)]
    tag_type_unique_id: Annotated[str, Field(min_length=1)]
    linked_level_unique_id: Annotated[str, Field(min_length=1)] | None = None
    orientation: Literal["Horizontal", "Vertical"] = "Horizontal"


class TagUpdate(DocumentationItem):
    view_unique_id: str
    tag_unique_id: str
    expected_head_mm: Point3
    head_mm: Point3
    leader: bool | None = None
    elbow_mm: Point3 | None = None
    leader_end_mm: Point3 | None = None

    @model_validator(mode="after")
    def leader_geometry(self):
        if (self.elbow_mm is not None or self.leader_end_mm is not None) and self.leader is False:
            raise ValueError("Изгиб требует включённой выноски")
        return self


class TagDelete(DocumentationItem):
    view_unique_id: str
    tag_unique_id: str


class RoomScheduleClone(DocumentationItem):
    source_schedule_unique_id: str
    view_unique_id: str
    link_instance_unique_id: str
    name: str
    number_prefix: Annotated[str, Field(min_length=3)]


class DimensionCreate(DocumentationItem):
    view_unique_id: str
    dimension_type_unique_id: str
    start_mm: Point3
    end_mm: Point3
    references: Annotated[list[str], Field(min_length=2, max_length=100, description="Уникальные stable_reference из get_revit_dimension_references или образца. Не выдумывать строки.")]

    @model_validator(mode="after")
    def unique_references(self):
        if len(set(self.references)) != len(self.references):
            raise ValueError("Повтор размерной привязки")
        if self.start_mm == self.end_mm:
            raise ValueError("Нулевая размерная линия")
        return self


class PlanCreate(DocumentationItem):
    view_type_unique_id: str
    level_unique_id: str
    name: Annotated[str, Field(min_length=1)]
    template_unique_id: str | None = None
    scale: Annotated[int, Field(ge=1, le=24000)] | None = None
    crop_min_mm: Point3 | None = None
    crop_max_mm: Point3 | None = None

    @model_validator(mode="after")
    def crop_pair(self):
        if (self.crop_min_mm is None) != (self.crop_max_mm is None):
            raise ValueError("Нужны обе границы crop")
        if self.crop_min_mm is not None and any(a >= b for a, b in zip(self.crop_min_mm, self.crop_max_mm)):
            raise ValueError("crop_min должен быть меньше crop_max")
        return self


class SectionCreate(DocumentationItem):
    view_type_unique_id: str
    name: Annotated[str, Field(min_length=1)]
    origin_mm: Point3
    up: Point3 = Field(description="Единичный вектор вверх; безразмерный.")
    direction: Point3 = Field(description="Единичный вектор взгляда, перпендикулярный up; безразмерный.")
    min_mm: Point3 = Field(description="Минимум объёма в локальном базисе (up×direction, up, direction), мм.")
    max_mm: Point3
    template_unique_id: str | None = None
    scale: Annotated[int, Field(ge=1, le=24000)] | None = None

    @model_validator(mode="after")
    def section_basis(self):
        if any(abs(sum(v*v for v in axis) - 1) > 1e-6 for axis in (self.up, self.direction)):
            raise ValueError("Требуются единичные векторы")
        if abs(sum(a*b for a, b in zip(self.up, self.direction))) > 1e-6:
            raise ValueError("up и direction должны быть перпендикулярны")
        if any(a >= b for a, b in zip(self.min_mm, self.max_mm)):
            raise ValueError("min должен быть меньше max")
        return self


def documentation_write(command, document_id, expected_revision, items, dry_run, session_id):
    return client.command(command, session_id, document_id=document_id, expected_revision=expected_revision,
                          items=[item.model_dump(mode="json", exclude_none=True) for item in items], dry_run=dry_run)


@mcp.tool(annotations=READ)
def get_revit_sheet_layout(document_id: str, sheet_unique_id: str, session_id: str | None = None) -> dict:
    """Измерить лист-образец: штампы, границы листа, центры/габариты/повороты окон, заголовки, точки вставки спецификаций. Всё в мм системы листа. В Revit 2020 смещения заголовков API не выдаёт."""
    return client.command("get_sheet_layout", session_id, document_id=document_id, sheet_unique_id=sheet_unique_id)


@mcp.tool(annotations=READ)
def get_revit_documentation_catalog(document_id: str, kind: Literal["title_blocks", "tag_types", "dimension_types", "view_types", "view_templates"],
                                    offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Типы оформления образца. title_blocks возвращает семейства 020_. Наличие типа марки не доказывает его применение: читайте get_revit_view_annotations."""
    return client.command("get_documentation_catalog", session_id, document_id=document_id, kind=kind, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_project_levels(document_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Уровни по отметке: building_story, structural, одинаковые отметки, планы/шаблоны/листы. Это признаки кандидатов; проектные этажи и типовые группы выбираются по образцу или с оператором."""
    return client.command("get_project_levels", session_id, document_id=document_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_view_annotations(document_id: str, view_unique_id: str, offset: Offset = 0, limit: Limit = 100, session_id: str | None = None) -> dict:
    """Существующие марки/размеры вида-образца: тип, цель, позиция головки, привязки размера и базис вида. Точки в мм внутренних координат. Не копировать project-specific ID в другой документ."""
    return client.command("get_view_annotations", session_id, document_id=document_id, view_unique_id=view_unique_id, offset=offset, limit=limit)


@mcp.tool(annotations=READ)
def get_revit_dimension_references(document_id: str, view_unique_id: str, element_unique_id: str,
                                   limit: Annotated[int, Field(ge=1, le=1000)] = 200, session_id: str | None = None) -> dict:
    """Кандидаты привязок элемента: datum, опорные плоскости семейства, плоские грани, рёбра. Геометрия в мм внутренних координат; стабильные ссылки проверяются Revit при dry_run размера. Геометрия связей не обходится."""
    return client.command("get_dimension_references", session_id, document_id=document_id, view_unique_id=view_unique_id, element_unique_id=element_unique_id, limit=limit)


@mcp.tool(annotations=WRITE)
def create_revit_sheets(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                        items: Annotated[list[SheetCreate], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Создать листы со штампами КПЛН 020_. Повтор номера отклоняется. dry_run выполняет создание/Commit и откат всей группы; пробные ID недействительны. Перед записью обновить context/revision."""
    return documentation_write("create_sheets", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def place_revit_views(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                      items: Annotated[list[ViewPlacement], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Разместить виды/спецификации на листе по измеренным координатам образца. Для вида point_mm=центр окна; reference_viewport_unique_id копирует положение из текущего документа. dry_run откатывает весь пакет."""
    return documentation_write("place_views", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def create_revit_tags(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                      items: Annotated[list[TagCreate], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Марки элементов, помещений и пространств по типам из образца. Для связи укажите link_instance_unique_id и UniqueId цели в связи; пространства только локальные. head_mm в текущем документе. dry_run проверяет Revit и откатывает пакет."""
    return documentation_write("create_tags", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def move_revit_sheet_items(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                           items: Annotated[list[SheetItemMove], Field(min_length=1, max_length=100)],
                           dry_run: bool = True, session_id: str | None = None) -> dict:
    """Переместить размещённые спецификации и окна видов, сохраняя ID. Точки — мм листа, Z=0; expected_point_mm взять из placement_point_mm в get_revit_sheet_layout. Защищает ручное положение, закрепление, группы и ведомости изменений. После Commit проверяет все габариты затронутых листов, включая аннотации/заголовки: конфликт или неизвестный габарит откатывает весь пакет. При нехватке места увеличить формат штампа через фактические параметры и повторно измерить поле. dry_run откатывает перемещение; обновить revision. RVT не сохраняется. Нужен мост 1.1.10."""
    if len({item.instance_unique_id for item in items}) != len(items):
        raise ValueError("Один экземпляр нельзя перемещать дважды в пакете")
    return documentation_write("move_sheet_items", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def update_revit_viewport_labels(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                                 items: Annotated[list[ViewportLabelUpdate], Field(min_length=1, max_length=100)],
                                 dry_run: bool = True, session_id: str | None = None) -> dict:
    """Изменить LabelOffset и при необходимости длину линии заголовка без перемещения плана. Ожидаемые значения читать из get_revit_sheet_layout. Смещения и длины в мм бумаги, Z=0; знак длины сохраняет направление линии. Защищает закрепление, группы и ручные изменения. Проверяет весь лист после Commit: конфликт или неизвестные габариты откатывают пакет. Проверка рамки/штампа дополнительно через check_revit_sheet_layout с измеренным полем. dry_run выполняет Commit и откат; обновить revision перед записью. RVT не сохраняется. Мост 1.1.12, Revit 2023/2024/2026; в 2020 API недоступен."""
    if len({item.viewport_unique_id for item in items}) != len(items):
        raise ValueError("Один заголовок нельзя изменять дважды в пакете")
    return documentation_write("update_viewport_labels", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=READ)
def check_revit_sheet_layout(document_id: str, sheet_unique_id: str,
                             minimum_gap_mm: Annotated[float, Field(ge=0, le=100, allow_inf_nan=False)] = 0,
                             content_min_mm: Point3 | None = None, content_max_mm: Point3 | None = None,
                             reserved_regions: Annotated[list[SheetReservedRegion], Field(max_length=100)] | None = None,
                             session_id: str | None = None) -> dict:
    """Проверить пересечения габаритов видов, их аннотаций/заголовков и спецификаций. content_min/max — измеренное поле внутри рамки, reserved_regions — штамп и служебные зоны. Без поля проверяются только края бумаги (frame_supplied=false). Неизвестная геометрия даёт complete=false/bounds_clear=false. Габариты консервативны: возможны пустоты; нужен итоговый PDF. Если композиция не помещается, увеличить допустимый формат основной надписи и пересчитать поле, а не уменьшать читаемость. Изменений нет. Нужен мост 1.1.10."""
    if (content_min_mm is None) != (content_max_mm is None):
        raise ValueError("Задайте обе границы полезного поля")
    if content_min_mm is not None:
        SheetReservedRegion(name="content", min_mm=content_min_mm, max_mm=content_max_mm)
    return client.command("check_sheet_layout", session_id, document_id=document_id, sheet_unique_id=sheet_unique_id,
                          minimum_gap_mm=minimum_gap_mm, content_min_mm=content_min_mm, content_max_mm=content_max_mm,
                          reserved_regions=[r.model_dump() for r in (reserved_regions or [])])


@mcp.tool(annotations=WRITE)
def tag_revit_linked_rooms(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                           items: Annotated[list[LinkedRoomTagCreate], Field(min_length=1, max_length=100)],
                           dry_run: bool = True, session_id: str | None = None) -> dict:
    """Пакетно маркировать помещения связи на планах этажей/потолков, без выносок. items задаёт план, экземпляр связи и тип RoomTagType. Уровень связи определяется по преобразованной отметке (10 мм), при неоднозначности задайте linked_level_unique_id. Стадия берётся из phase map связи для стадии плана. Точки помещений преобразуются GetTransform; за прямоугольной подрезкой и уже маркированные цели пропускаются. Сложная/разделённая подрезка отклоняется. До 1000 новых марок на запрос; ошибка откатывает всё. dry_run делает создание/Commit и откат; обновите revision после пробы. Фильтры/режим показа связи не воспроизводятся: проверьте вид визуально. RVT не сохраняется."""
    return documentation_write("tag_linked_rooms", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def update_revit_tags(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                      items: Annotated[list[TagUpdate], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Переместить существующие марки и изгибы выносок, не меняя цели и параметры модели. expected_head_mm — прочитанная позиция, head_mm/elbow_mm — мм внутренних координат. Проверяет владельца вида, закрепление и устаревшую позицию. dry_run выполняет Commit и откат всей группы; после пробы обновить revision. Изгиб независимой марки поддерживается для одной цели."""
    return documentation_write("update_tags", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def clone_revit_room_schedule(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                              items: Annotated[list[RoomScheduleClone], Field(min_length=1, max_length=100)],
                              dry_run: bool = True, session_id: str | None = None) -> dict:
    """Скопировать экспликацию помещений связи по образцу для фрагмента. Номера по явному префиксу, стадия по phase map; точки помещений внутри контура обрезки плана. Копирует оформление, заменяет фильтры; до 7 исключённых номеров. Сверяет номера строк с выбранными помещениями, иначе откат. Исходную спецификацию не меняет. dry_run по умолчанию; после пробы обновить revision. Для размещения использовать place_revit_views."""
    return documentation_write("clone_room_schedule", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def delete_revit_tags(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                      items: Annotated[list[TagDelete], Field(min_length=1, max_length=100)],
                      dry_run: bool = True, session_id: str | None = None) -> dict:
    """Удалить только перечисленные марки на указанных видах. Закреплённые, групповые и чужие марки отклоняются. Зависимые удаления откатывают весь пакет. dry_run выполняет Commit и откат; затем обновить revision. Элементы модели и RVT не меняются/не сохраняются."""
    return documentation_write("delete_tags", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=READ)
def get_revit_space_annotation_context(document_id: str, view_unique_id: str,
                                       radiator_unique_ids: Annotated[list[str], Field(min_length=1, max_length=100)],
                                       session_id: str | None = None) -> dict:
    """Размещённые пространства вида/стадии, границы, штатная принадлежность и точки расчёта помещения заданных радиаторов. При отсутствии принадлежности/точки требуется проверка по плану. Проверок коллизий нет. Обрезку вида проверить отдельно. Список радиаторов задаёт вызывающий, оборудование автоматически радиаторами не считается."""
    return client.command("get_space_annotation_context", session_id, document_id=document_id,
                          view_unique_id=view_unique_id, radiator_unique_ids=radiator_unique_ids)


@mcp.tool(annotations=WRITE)
def create_revit_dimensions(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                            items: Annotated[list[DimensionCreate], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Линейные размеры и цепочки по реальным Revit stable_reference. Тип и принцип привязки берутся из образца. start/end в мм внутренних координат. dry_run проверяет геометрию и откатывает группу; при сбое отменяется весь пакет."""
    return documentation_write("create_dimensions", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def create_revit_plan_views(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                            items: Annotated[list[PlanCreate], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Создать планы по выбранным проектным уровням с типом/шаблоном из образца. crop_min/max в мм внутренних координат. Конфликт масштаба с шаблоном отклоняется. dry_run откатывает создание; ID пробного запуска не использовать."""
    return documentation_write("create_plan_views", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=WRITE)
def create_revit_sections(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                          items: Annotated[list[SectionCreate], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Создать разрезы: origin во внутренних координатах, up/direction единичные, min/max в локальном базисе, длины в мм. Тип/шаблон из образца. dry_run проверяет создание/Commit и откатывает пакет."""
    return documentation_write("create_sections", document_id, expected_revision, items, dry_run, session_id)


class PlanRangeCopy(BaseModel):
    view_unique_id: str
    source_view_unique_id: str


class PlanCropCopy(BaseModel):
    model_config = ConfigDict(extra="forbid")
    view_unique_id: Annotated[str, Field(min_length=1)]
    source_view_unique_id: Annotated[str, Field(min_length=1)]

    @model_validator(mode="after")
    def different_views(self):
        if self.view_unique_id == self.source_view_unique_id:
            raise ValueError("Нужны разные планы")
        return self


@mcp.tool(annotations=WRITE)
def copy_revit_plan_crop(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                         items: Annotated[list[PlanCropCopy], Field(min_length=1, max_length=100)],
                         dry_run: bool = True, session_id: str | None = None) -> dict:
    """Перенести фактический цельный контур обрезки образца на планы других этажей. Совпадающая ориентация, без разделённых областей и scope box целевого вида. Шаблон и диапазон высот не меняются. Проверяет вершины после Regenerate; ошибка откатывает пакет. После переноса заново проверить лист и экспликацию. dry_run делает Commit и откат; обновить revision. RVT не сохраняется. Мост 1.1.11."""
    return documentation_write("copy_plan_crop", document_id, expected_revision, items, dry_run, session_id)


class PlanRangeSet(BaseModel):
    model_config = ConfigDict(extra="forbid")
    view_unique_id: Annotated[str, Field(min_length=1)]
    top_offset_mm: Annotated[float, Field(allow_inf_nan=False)]
    cut_offset_mm: Annotated[float, Field(allow_inf_nan=False)]
    bottom_offset_mm: Annotated[float, Field(allow_inf_nan=False)]
    depth_offset_mm: Annotated[float, Field(allow_inf_nan=False)]


@mcp.tool(annotations=WRITE)
def set_revit_plan_range(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                         items: Annotated[list[PlanRangeSet], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Задать четыре плоскости секущего диапазона напрямую: все смещения в мм от связанного уровня целевого плана. Проверяет допустимость в Revit. Общий шаблон не меняется; управляемый шаблоном диапазон отклоняется. dry_run выполняет Commit и откат, затем обновить revision. RVT не сохраняется."""
    return documentation_write("set_plan_range", document_id, expected_revision, items, dry_run, session_id)


@mcp.tool(annotations=READ)
def get_revit_plan_range(document_id: str, view_unique_id: str, session_id: str | None = None) -> dict:
    """Секущий диапазон плана: уровни и смещения верхней, секущей, нижней плоскостей и глубины в мм."""
    return client.command("get_plan_range", session_id, document_id=document_id, view_unique_id=view_unique_id)


@mcp.tool(annotations=READ)
def get_revit_mep_connections(document_id: str, unique_ids: Annotated[list[str], Field(min_length=1, max_length=100)], session_id: str | None = None) -> dict:
    """Физические соединители MEP и реальные соседние элементы. Точки в мм; part_type позволяет найти тройники. Направление потока не определяется."""
    return client.command("get_mep_connections", session_id, document_id=document_id, unique_ids=unique_ids)


@mcp.tool(annotations=WRITE)
def copy_revit_plan_range(document_id: str, expected_revision: Annotated[int, Field(ge=0)],
                          items: Annotated[list[PlanRangeCopy], Field(min_length=1, max_length=100)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Скопировать секущий диапазон с проверенного плана ТОГО ЖЕ уровня. Общий шаблон не меняется; управляемый шаблоном диапазон отклоняется. dry_run выполняет Commit и откат, затем обновить revision. RVT не сохраняется."""
    return documentation_write("copy_plan_range", document_id, expected_revision, items, dry_run, session_id)


if __name__ == "__main__":
    mcp.run(transport="stdio")
