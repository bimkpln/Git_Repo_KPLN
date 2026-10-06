"""Codex <-> MCP stdio <-> loopback HTTP <-> KPLN_Loader / Revit Idling."""
from __future__ import annotations

from typing import Annotated, Literal
from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations
from pydantic import BaseModel, ConfigDict, Field, StrictFloat, StrictInt, StrictStr
from bridge_client import BridgeClient

mcp = FastMCP("KPLN_RevitMcpBridge", instructions=(
    "Работа с локальным Revit через KPLN_Loader. Сначала list_revit_sessions и get_revit_context. "
    "Для чтения/записи передавайте полученный document_id, для записи также expected_revision. "
    "set_revit_parameters и set_revit_family_type_parameters по умолчанию строят план; dry_run=False применяет его. После таймаута "
    "не повторяйте запись: get_revit_operation. Правила генерации семейств храните в KPLN_RevitRules, "
    "проверки кандидатов — в KPLN_JobCandidateModelCheckRules; "
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
def print_revit_sheets_pdf(document_id: str, expected_revision: Annotated[int, Field(ge=0)], settings: BatchPrintSettings | None = None,
                           dry_run: bool = True, timeout_seconds: Annotated[int, Field(ge=5, le=60)] = 30, session_id: str | None = None) -> dict:
    """Все листы активной модели через KPLN Пакетную выдачу/PDFCreator без окон. dry_run строит план.

    Новая папка запуска внутри outputPDFFolder; модель не сохраняется, временные изменения откатываются.
    После outcome_pending читать get_revit_operation; при таймауте принтера не повторять печать автоматически.
    Нужны обновлённая KPLN_Publication и настроенный PDFCreator. Связи не печатаются.
    """
    return client.command("print_sheets_pdf", session_id, document_id=document_id, expected_revision=expected_revision,
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
def export_revit_sheets_pdf(document_id: str, expected_revision: Annotated[int, Field(ge=0)], dry_run: bool = True, session_id: str | None = None) -> dict:
    """Все листы текущего проекта в PDF штатным экспортом Revit 2023/2024. Сначала dry_run=True: перечень листов и пропусков. dry_run=False создаёт отдельный PDF каждого печатаемого листа и manifest.json в новом локальном каталоге моста; возвращает абсолютные пути, SHA256 и соответствие листам. Размер бумаги по листу, 100%, цвет. Документ не сохраняется, принтеры/настройки печати не меняются. Пропуски/ошибки дают complete=false. После таймаута читать get_revit_operation, НЕ повторять экспорт. Для визуального контроля прочитать реальные PDF, проверить страницы и размеры. Revit 2020 не поддерживается."""
    return client.command("export_sheets_pdf", session_id, document_id=document_id, expected_revision=expected_revision, dry_run=dry_run)


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
    """В открытом редакторе семейства: FamilyManager, формулы, параметры и значения типоразмеров; Double в единицах Revit. Только чтение."""
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


if __name__ == "__main__":
    mcp.run(transport="stdio")
