# Coordinator AI MCP

## Purpose

This project is moving to a single MCP-based architecture for Revit tools.

Both scenarios use the same MCP tool architecture. Every running Revit process owns
its own local MCP server:

```text
Revit -> "Work with model" WPF tab -> Internal MCP Client -> Revit MCP Server -> Revit API

Claude/Cursor/VS Code/Codex -> MCP stdio proxy -> selected Revit MCP Server -> Revit API
```

MCP is the source of truth for Revit tool discovery and execution. If the internal LLM still needs OpenAI-compatible `tools`, that format is only an adapter for the model, not a separate Revit tool architecture.

## Дополнительные инструкции отделов в WPF

В настройках, на вкладке «Внешняя модель», под общим системным промптом
находится редактор «Дополнительная инструкция отдела». Выберите АР, КР или
ИОС, введите соответствующий текст и нажмите общую кнопку сохранения настроек.
Переключение отдела сохраняет черновик в окне; запись в БД выполняется только
при сохранении. Первоначально все три инструкции пустые.

Тексты хранятся в таблице `Settings` той же БД настроек, что и
`ExternalModel.SystemPrompt` и описания команд, а не в пользовательских БД логов:

| Ключ | Назначение |
| --- | --- |
| `ExternalModel.DepartmentPrompt.AR` | АР |
| `ExternalModel.DepartmentPrompt.KR` | КР |
| `ExternalModel.DepartmentPrompt.IOS` | Все подразделения ИОС |

Отдел определяется автоматически через существующий корпоративный механизм
`CurrentUserContextService` / `SQLiteMainService`. Профили: АР (ID 2),
КР (ID 3), ИОС (ID 4, 5, 6, 7, 41, 51, 71, 72: ОВиК, ВК, ЭОМ, СС, ИТП,
ПТ, АВ, СПС). Для новых подразделений проверяется цепочка
`DependentSubDepId` до известного отдела, с защитой от циклов.
Для BIM, ALL или неопределенного отдела дополнительная инструкция не добавляется.

Перед каждым новым пользовательским вопросом инструкции перечитываются из БД
в фоновом потоке. Общий промпт дополняется только текстом выбранного отдела.
Этот снимок используется во всех промежуточных tool-раундах вопроса, без
дублирования в истории. Сохраненные изменения инструкций отделов применяются
со следующего вопроса, без повторного открытия чата. Пустое поле отключает
инструкцию соответствующего отдела; общий промпт остается без изменений.

Ожидание чтения ограничено 5 секундами и поддерживает отмену запроса.
При ошибке БД используются ранее загруженные настройки окна, если отдел удалось
определить. При превышении ожидания или смене пути БД во время чтения вопрос
выполняется без дополнительной инструкции. Зависшее сетевое чтение нельзя
принудительно остановить средствами SQLite: оно остается в фоне, но в этом окне
не создаются дополнительные параллельные чтения до его завершения.

В DiagnosticLogger событие `DEPARTMENT_PROMPT.SELECTED` показывает отдел,
профиль, признак применения и источник настроек (`database`,
`cached_settings`, `timeout` или `database_changed`). Полный текст
дополнительной инструкции в этом событии не записывается.

Механизм относится только к чату «Работа с моделью». Системные промпты внешних
MCP-хостов не изменяются. Логика вкладки «Поиск по регламентам» не меняется.
При использовании общей БД сохраненные инструкции общие для всех ее пользователей.

Автономная проверка: после сборки Debug2023 запустите
`Tests\RunDepartmentPromptTests.ps1`. Тесты используют отдельную БД в
`obj\DepartmentPromptTests`, не изменяя рабочую БД и файл выбора ее пути.
При необходимости укажите пути `-PluginAssembly` и `-NativeSqliteAssembly`
(существующая x64-библиотека `e_sqlite3.dll` из комплекта загрузчика KPLN).

## Current State

Implemented:

- shared MCP tool contracts;
- one shared MCP registry for 55 Revit tools and 3 WPF-only attachment tools;
- one shared executor for all registered read-only and UI-changing tools;
- `ExternalEvent` wrapper for safe Revit API execution;
- per-process local HTTP JSON-RPC endpoint inside Revit;
- automatic registration and discovery of all running Revit processes;
- internal MCP client used by the WPF "Work with model" tab for `tools/list` and `tools/call`;
- stdio proxy exe for external MCP hosts;
- consistent structured MCP errors and user-facing error messages;
- pagination for potentially large tool results;
- a limit of five real tool executions per model response in the WPF agent loop;
- shared `DiagnosticLogger` and grouped `ChatLogger` records with scenario markers;
- structured per-user SQLite history for the WPF "Work with model" tab;
- user-facing WPF progress statuses that do not expose model chain-of-thought.
- explicit file attachments in the WPF chat, including DOCX, with local text extraction, search and paged reading.
- official Autodesk Product Help MCP tools in the WPF "Work with model" agent.

## Autodesk Product Help MCP

The WPF "Work with model" agent combines two independent MCP sources before each
request is sent to the configured AI model:

```text
Local Revit MCP -> model and Revit API tools
Autodesk Product Help MCP -> official Autodesk documentation search
```

The official remote endpoint is:

```text
https://developer.api.autodesk.com/knowledge/public/v1/mcp
```

It uses MCP Streamable HTTP, is read-only, and currently requires neither an
Autodesk account nor an API key. Its tools are discovered dynamically and added to
the OpenAI-compatible tool list used by the WPF agent:

- `get_available_products` validates Autodesk product and release codes;
- `search_help_content` searches official Autodesk Help and Support content.

For Revit searches, the wrapper supplies safe defaults when the model omits them:

```text
product_code = RVT
release_code = current Revit major version
locale = en_US
max_results = 3 (maximum 5)
```

The system prompt tells the model to translate Russian search terms into English,
use the documentation tools for reference questions, and cite article titles and
the direct URLs returned by Autodesk. These tools do not inspect the open model;
model data must still be obtained through the local Revit MCP tools.

Only the documentation search arguments are sent to Autodesk. The plugin does not
send the Revit model, API key, attached files, complete chat history, or local MCP
request headers to the remote documentation endpoint.

If Autodesk Product Help or the internet is unavailable, its tools are omitted for
that request, `AUTODESK_HELP.TOOLS_ERROR` is written to `DiagnosticLogger`, and all
local Revit tools continue to work. Remote discovery and search calls have a
90-second upper timeout and also honor the WPF request cancellation button.
Successful discovery is logged as
`AUTODESK_HELP.TOOLS_LOADED`; tool calls use the normal `TOOL.START`, `TOOL.END`,
and `TOOL.ERROR` records with `toolArea=AutodeskHelp` and
`executionPath=autodesk_product_help_mcp`.

External agents such as Codex, Claude, or Cursor should connect to Autodesk Product
Help directly as a second MCP server. The local Revit endpoint intentionally
continues to expose only the tools owned by the selected Revit process.

Endpoint allocation:

```text
http://127.0.0.1:<first-free-port-in-48731-48799>/mcp/
```

The first Revit process normally receives port `48731`, the second receives
`48732`, and so on. The WPF client receives the endpoint directly from its own
Revit process and never falls back to another process.

## Supported MCP Methods

- `initialize`
- `ping`
- `tools/list`
- `tools/call`

## External MCP Stdio Proxy

Many MCP hosts start tools through `stdio`: the host launches an executable, sends one JSON-RPC message per stdin line, and reads JSON-RPC responses from stdout.

This repository includes a small proxy executable:

```text
ExternalMcpStdioProxy\RevitMcpStdioProxy.csproj
```

The proxy does not call the Revit API directly. By default it discovers registered
Revit processes and forwards MCP JSON-RPC messages to the selected process endpoint.

```text
%LOCALAPPDATA%\KPLN\CoordinatorAI\MCP\instances\*.json
```

Development build example. For actual external agent setup, use the distributed executable and stable local path from "External MCP Host Configuration" below:

```powershell
"&C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" `
  "C:\Users\mtarchokov\Source\Repos\Рабочие файлы KPLN_AI-coordinator\ExternalMcpStdioProxy\RevitMcpStdioProxy.csproj" `
  /p:Configuration=Debug /p:Platform=x64 /v:minimal
```

The repository build output is a development artifact. Do not use `bin\Debug` or `bin\Release` as the stable external agent command path.

The proxy marks requests as:

```text
external_mcp
```

You can override this label with the `REVIT_MCP_CLIENT_NAME` environment variable if the same proxy is used from another host.

Automatic discovery is the recommended mode. An explicit endpoint is still
supported for diagnostics and backwards compatibility:

```json
{
  "mcpServers": {
    "revit": {
      "command": "C:\\Users\\<UserName>\\AppData\\Local\\KPLN\\CoordinatorAI\\MCP\\RevitMcpStdioProxy.exe",
      "args": [
        "--endpoint",
        "http://127.0.0.1:48731/mcp/"
      ]
    }
  }
}
```

Revit must be running with the plugin loaded before the external MCP host can call Revit tools.

### Selecting a Revit Process

The stdio proxy adds three routing tools:

- `list_revit_instances` lists live registered Revit processes with version, PID,
  model name, window title, endpoint, and instance id;
- `select_revit_instance` selects the exact process for subsequent Revit tools;
- `get_selected_revit_instance` returns the current selection.

If exactly one matching Revit process is running, the proxy selects it
automatically. If several processes are running, a normal Revit tool call returns
`REVIT_INSTANCE_REQUIRED` until `select_revit_instance` is called.

The proxy never silently switches to another Revit after selection. If the
selected process closes, calls return `REVIT_INSTANCE_UNAVAILABLE`.

During staged deployment, if no instance registrations exist and no selector was
configured, the proxy tries the legacy endpoint
`http://127.0.0.1:48731/mcp/`. This keeps one-process installations working,
but exact process selection requires the updated Revit plugin and its instance
registrations.

Optional startup selectors are available when separate MCP server entries are
preferred:

```text
--instance-id <instance_id>
--process-id <pid>
--revit-version 2024
--model "Project name.rvt"
```

Selectors can be combined. They must resolve to exactly one live process before a
Revit tool is executed.

## External MCP Host Configuration

Use this section as the current connection guide. The stable proxy path for every Windows user is:

```text
%LOCALAPPDATA%\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe
```

For example:

```text
C:\Users\<UserName>\AppData\Local\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe
```

The approved company distribution copy of `RevitMcpStdioProxy.exe` must be taken from:

```text
Z:\Отдел BIM\03\_Скрипты\03\_Скрипты\1. MCP-мост для ИИ-агента
```

Expected source executable:

```text
Z:\Отдел BIM\03\_Скрипты\03\_Скрипты\1. MCP-мост для ИИ-агента\RevitMcpStdioProxy.exe
```

Copy this executable to the local stable path under `%LOCALAPPDATA%` before configuring an agent. External agents should run the local copy, not the network file and not a repository `bin\Debug` or `bin\Release` artifact.

For developers, the proxy project is:

```text
ExternalMcpStdioProxy\RevitMcpStdioProxy.csproj
```

It is configured to place a local build in the same `%LOCALAPPDATA%\KPLN\CoordinatorAI\MCP` folder on the developer's computer.

All external agents connect by the same principle:

```text
Agent -> stdio MCP -> RevitMcpStdioProxy.exe -> selected registered endpoint -> Revit
```

Before connecting an external agent:

- start Revit;
- open or create a Revit model;
- make sure the plugin is loaded;
- check that the required Revit process appears in the local instance registry.

Instance discovery check:

```powershell
$instancesPath = Join-Path $env:LOCALAPPDATA "KPLN\CoordinatorAI\MCP\instances"
Get-ChildItem -LiteralPath $instancesPath -Filter "*.json" |
  ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json } |
  Format-Table revit_version, process_id, instance_id, endpoint
```

Example:

```text
revit_version process_id instance_id                       endpoint
------------- ---------- -----------                       --------
2020          37588      1d8f...                           http://127.0.0.1:48731/mcp/
2024          3132       a24b...                           http://127.0.0.1:48732/mcp/
```

### Codex

Codex uses a TOML config file.

Typical Windows path:

```text
C:\Users\<UserName>\.codex\config.toml
```

Example:

```toml
[mcp_servers.revit]
command = 'C:\Users\<UserName>\AppData\Local\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe'
args = []
startup_timeout_sec = 30
```

Optional log marker:

```toml
[mcp_servers.revit]
command = 'C:\Users\<UserName>\AppData\Local\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe'
args = []
startup_timeout_sec = 30

[mcp_servers.revit.env]
REVIT_MCP_CLIENT_NAME = 'external_mcp'
```

Restart Codex after changing the config.

Test prompt:

```text
Use the Revit MCP server and tell me the active Revit view name.
```

### Claude Desktop

Claude Desktop uses a JSON config with an `mcpServers` object.

The config location can vary by app version. On Windows it is commonly:

```text
%APPDATA%\Claude\claude_desktop_config.json
```

Example:

```json
{
  "mcpServers": {
    "revit": {
      "command": "C:\\Users\\<UserName>\\AppData\\Local\\KPLN\\CoordinatorAI\\MCP\\RevitMcpStdioProxy.exe",
      "args": [],
      "env": {
        "REVIT_MCP_CLIENT_NAME": "claude_external_mcp"
      }
    }
  }
}
```

Restart Claude Desktop after changing the config.

Test prompt:

```text
Use the Revit MCP tool and list several model categories.
```

### Cursor

Cursor uses MCP server entries with a command, args and optional env values. Depending on the Cursor version, this can be configured through the UI or through an `mcp.json` file.

Example:

```json
{
  "mcpServers": {
    "revit": {
      "command": "C:\\Users\\<UserName>\\AppData\\Local\\KPLN\\CoordinatorAI\\MCP\\RevitMcpStdioProxy.exe",
      "args": [],
      "env": {
        "REVIT_MCP_CLIENT_NAME": "cursor_external_mcp"
      }
    }
  }
}
```

After changing the config, restart Cursor or reload MCP servers from Cursor settings.

Test prompt:

```text
Use Revit MCP and get the active view in Revit.
```

### VS Code MCP Clients

VS Code does not have one universal MCP config file for every setup. The exact file depends on the MCP extension or agent extension being used.

The values are the same:

```text
Name: revit
Transport: stdio
Command: C:\Users\<UserName>\AppData\Local\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe
Args: empty
```

If the extension supports JSON config, the entry usually looks like:

```json
{
  "mcpServers": {
    "revit": {
      "command": "C:\\Users\\<UserName>\\AppData\\Local\\KPLN\\CoordinatorAI\\MCP\\RevitMcpStdioProxy.exe",
      "args": [],
      "env": {
        "REVIT_MCP_CLIENT_NAME": "vscode_external_mcp"
      }
    }
  }
}
```

### DeepSeek Harness

DeepSeek Harness is plugin-oriented, so the exact connection place depends on the Harness version and selected plugin/config workflow.

Use these values when adding the local MCP server:

```text
Name: revit
Transport: stdio
Command: C:\Users\<UserName>\AppData\Local\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe
Args: empty
Environment:
  REVIT_MCP_CLIENT_NAME=deepseek_harness_external_mcp
```

If Harness asks for JSON-like MCP config, use:

```json
{
  "mcpServers": {
    "revit": {
      "command": "C:\\Users\\<UserName>\\AppData\\Local\\KPLN\\CoordinatorAI\\MCP\\RevitMcpStdioProxy.exe",
      "args": [],
      "env": {
        "REVIT_MCP_CLIENT_NAME": "deepseek_harness_external_mcp"
      }
    }
  }
}
```

### Perplexity

Perplexity may use local or remote MCP connectors depending on the product surface.

For this Revit integration, use a local MCP connector because Revit is running on the user's machine and the endpoint is local:

```text
127.0.0.1
```

A remote/cloud MCP connector usually cannot reach the user's local Revit process.

Use these values:

```text
Name: revit
Transport: stdio
Command: C:\Users\<UserName>\AppData\Local\KPLN\CoordinatorAI\MCP\RevitMcpStdioProxy.exe
Args: empty
Environment:
  REVIT_MCP_CLIENT_NAME=perplexity_external_mcp
```

### Troubleshooting External Hosts

If an external agent does not see Revit tools:

- check that Revit is open and the plugin has loaded;
- check the JSON registrations in
  `%LOCALAPPDATA%\KPLN\CoordinatorAI\MCP\instances`;
- check that `RevitMcpStdioProxy.exe` exists in `%LOCALAPPDATA%\KPLN\CoordinatorAI\MCP`;
- restart the external agent after changing its config;
- make sure the config points to the local installed proxy path, not to the repository `bin\Debug` path.

If rebuilding fails because `RevitMcpStdioProxy.exe` is locked, close the external agent that started it, then build again.

## Registered Tools

Currently registered tools:

- `get_active_view_in_revit`
- `get_project_view_items`
- `get_all_elements_shown_in_view`
- `get_category_by_keyword`
- `get_elements_by_category`
- `get_model_categories`
- `get_categories_from_elementids`
- `get_object_classes_from_elementids`
- `get_element_types_for_elementids`
- `get_all_elementids_for_specific_type_ids`
- `get_all_used_families_in_model`
- `get_all_used_families_of_category`
- `get_all_used_types_of_a_family`
- `get_all_elements_of_specific_families`
- `get_parameters_from_elementid`
- `get_parameter_value_for_element_ids`
- `get_all_additional_properties_from_elementid`
- `get_additional_property_for_all_elementids`
- `get_revitlookup_like_properties`
- `get_titleblock_family_parameters_description`
- `get_location_for_element_ids`
- `get_boundingboxes_for_element_ids`
- `get_boundary_lines`
- `get_room_boundary_lines`
- `get_host_id_for_element_ids`
- `get_material_layers_from_types`
- `set_view_section_box_to_elements`
- `get_model_file_info`
- `get_all_project_units`
- `get_all_warnings_in_the_model`
- `get_all_workset_information`
- `get_worksets_from_elementids`
- `get_worksharing_information_for_element_ids`
- `get_user_selection_in_revit`
- `set_user_selection_in_revit`
- `get_graphic_overrides_for_element_ids_in_view`
- `get_graphic_filters_applied_to_views`
- `get_all_parameter_filters_in_model`
- `get_graphic_overrides_view_filters`
- `get_category_visibility_overrides_in_view`
- `get_workset_visibility_in_view`
- `get_link_graphics_overrides_in_view`
- `get_detailed_link_graphics_overrides_in_view`
- `get_all_phases_in_model`
- `get_phase_visibility_settings`
- `get_if_elements_pass_filter`
- `get_viewports_and_schedules_on_sheets`
- `get_schedules_info_and_columns`
- `get_schedule_rows_and_values`
- `get_schedule_sorting_info`
- `get_journal_entries_since`
- `get_revit_links_in_model`
- `get_revit_link_elements`
- `get_revit_link_categories`
- `get_revit_link_elements_by_category`
- `get_selected_revit_link_element_id`
- `get_revit_link_element_properties`

## Endpoint Check

After Revit starts and loads the plugin:

```powershell
$instancesPath = Join-Path $env:LOCALAPPDATA "KPLN\CoordinatorAI\MCP\instances"
$instance = Get-ChildItem -LiteralPath $instancesPath -Filter "*.json" |
  ForEach-Object { Get-Content -LiteralPath $_.FullName -Raw | ConvertFrom-Json } |
  Where-Object { $_.revit_version -eq 2024 } |
  Select-Object -First 1

$endpoint = $instance.endpoint
Invoke-WebRequest $endpoint -TimeoutSec 5
```

Change `2024` to the required Revit version. If several processes of the same
version are running, select the registration by `process_id` or `instance_id`.

Expected response:

```json
{
  "name": "revit-mcp",
  "status": "running",
  "endpoint": "http://127.0.0.1:48732/mcp/",
  "instance_id": "a24b...",
  "process_id": 3132,
  "revit_version": 2024,
  "tools": 57,
  "wpf_tools": 60
}
```

`tools` is the number visible to external MCP hosts. `wpf_tools` also includes the three internal file-attachment tools.

## tools/list Check

```powershell
$body = @{
  jsonrpc = "2.0"
  id = 1
  method = "tools/list"
  params = @{}
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json" -Body $body |
  ConvertTo-Json -Depth 20
```

## tools/call Check

```powershell
$body = @{
  jsonrpc = "2.0"
  id = 2
  method = "tools/call"
  params = @{
    name = "get_active_view_in_revit"
    arguments = @{}
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json" -Body $body |
  ConvertTo-Json -Depth 20
```

## Project Views, Schedules And Sheets

`get_project_view_items` returns user-facing project views by default. It does
not include view templates or Revit internal/browser views unless requested.

Current use case - return all ordinary project views, 200 at a time:

```json
{
  "name": "get_project_view_items",
  "arguments": {
    "itemTypes": ["view"],
    "limit": 200,
    "offset": 0
  }
}
```

Future combined use case:

```json
{
  "name": "get_project_view_items",
  "arguments": {
    "itemTypes": ["view", "schedule", "sheet", "view_template"],
    "limit": 200,
    "offset": 0
  }
}
```

Optional `nameContains` performs a case-insensitive name search. Optional
`viewTypes` accepts Revit API `ViewType` names such as `FloorPlan`, `ThreeD`,
`Section`, `Elevation`, `DrawingSheet`, or `Schedule`. Set `includeInternal` to
true only when Revit internal and browser views are deliberately required.

Every item contains `id`, `name`, `itemType`, `viewType`, `isTemplate`, and
`canBePrinted`. Ordinary views and templates also report level, discipline,
phase, phase filter, and assigned view template where available. Sheets report
sheet number and placeholder status. Schedules report schedule type, category,
key-schedule status, and material-takeoff status. If an optional property cannot
be read, that item remains in the result and contains `metadataErrors`.

The result uses `items`, `total_count`, `returned_count`, `limit`, `offset`,
`has_more`, and `next_offset`. Keep all filters unchanged and continue with
`next_offset` until `has_more` is false. Ordering is deterministic: item kind,
then name, then element ID.

## Phase Visibility Tool Check

The command below reads the active view discipline, phase, assigned phase filter and all phase filters in the model. Add Revit element ids to `list_elementIds` to calculate their phase status and whether the assigned phase filter allows that status.

```powershell
$body = @{
  jsonrpc = "2.0"
  id = 3
  method = "tools/call"
  params = @{
    name = "get_phase_visibility_settings"
    arguments = @{
      list_elementIds = @(123456, 123457)
      includeAllFilters = $true
      limit = 200
      offset = 0
    }
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json; charset=utf-8" -Body $body |
  ConvertTo-Json -Depth 30
```

The `allowed_by_phase_filter` field evaluates phase rules only. Other view settings can still hide an element.

Revit API 2020/2023/2024 does not expose the phase-status graphic overrides configured on the Graphic Overrides tab of the Phasing dialog. Colors, line settings, patterns, halftone and materials must be inspected manually in Revit when they may affect the diagnosis.

## Schedule Cell Values

`get_schedule_rows_and_values` reads displayed text from the Body section of one
`ViewSchedule`, including calculated formula/percentage fields and totals. It uses
`ViewSchedule.GetCellText`, not element parameters or `TableSectionData.GetCellText`.
No transaction, view activation, or model changes are needed.

```json
{
  "jsonrpc": "2.0",
  "id": 30,
  "method": "tools/call",
  "params": {
    "name": "get_schedule_rows_and_values",
    "arguments": {
      "scheduleId": 123456,
      "limit": 200,
      "offset": 0
    }
  }
}
```

Replace `123456` with the schedule view ID, not its placement ID on a sheet.
The `columns` array describes visible Body columns. Optional `columnIndexes`
selects these physical table indices, not field IDs or field-order indices that
include hidden fields; omitted or empty means all visible columns.
Hidden fields are excluded from the mapping. If embedded schedules or another
table structure prevents a reliable mapping, `field_mapping_available` is false,
field metadata is null, and `warnings` explains this; cell text is still read by
its actual coordinates.

The `items` array contains rows of this form (example values):

```json
{
  "rowIndex": 2,
  "cells": [
    { "columnIndex": 0, "text": "Beam" },
    { "columnIndex": 1, "text": "12,50 m2" },
    { "columnIndex": 2, "text": "25%" }
  ]
}
```

Values are strings with the units, rounding and language displayed by Revit,
not unrounded numbers or formula expressions. Empty/group/total rows are
preserved: a row is not necessarily one element, and no row-to-element ID mapping
is inferred. `firstRowIndex` and `firstColumnIndex` give the actual API bounds.

Pagination uses `total_count`, `limit`, `offset`, `returned_count`, `has_more` and
`next_offset`. A page contains at most 200 complete Body rows and 204800 UTF-8
bytes of the result JSON, including column metadata. The MCP response envelope
and its text representation add transport overhead. Follow `next_offset` exactly;
a size/time limit may produce fewer rows than requested. Keep the schedule and
column selection unchanged between pages; results are not a persistent snapshot.

A failed cell has `text: null` and an `error`; other cells remain available.
`cell_error_count` counts these failures. `stop_reason` is `end`, `row_limit`,
`size_limit`, or `time_limit`. Cancellation and a five-second cooperative budget
are checked between cell reads; a single synchronous Revit API call cannot be
interrupted. If even one complete row or column metadata cannot fit, the command
returns an explicit error and asks for fewer `columnIndexes`, without silently
truncating values or advancing the offset past an unread row.

Run the pagination/error/cancellation tests without Revit:

```powershell
& ".\Tests\RunScheduleCellPageReaderTests.ps1"
```

Live verification after rebuilding and restarting Revit: read a schedule with a
formula, a percentage, a hidden field before visible fields, and enabled totals.
Compare cell text with the schedule and follow all pages using `next_offset`.
These API checks must be performed in the target Revit version; the standalone
tests exercise only the pagination reader.

## Paginated Tool Results

Tools that can return a large amount of model data support a shared pagination shape.

For element/list-based tools:

```json
{
  "items": [],
  "total_count": 0,
  "limit": 200,
  "offset": 0,
  "has_more": false,
  "next_offset": null
}
```

The default and maximum page size is `200` items. Use `next_offset` from the previous response to request the next page.

`get_viewports_and_schedules_on_sheets` is an intentional exception: it processes one sheet per page. Pagination is applied to the input sheet ids before any sheet content is collected, and the same `list_elementIds` must be sent again with `next_offset`. Elements owned by the sheet are collected with `ElementOwnerViewFilter` instead of a view-specific collector, avoiding unnecessary background initialization of every sheet. The one-sheet boundary gives Revit a safe opportunity to process WPF cancellation between heavy sheets.

Example:

```powershell
$body = @{
  jsonrpc = "2.0"
  id = 20
  method = "tools/call"
  params = @{
    name = "get_elements_by_category"
    arguments = @{
      categoryId = -2000011
      limit = 200
      offset = 0
    }
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json; charset=utf-8" -Body $body |
  ConvertTo-Json -Depth 20
```

The response also keeps the old collection name where possible, for example `element_ids`, `warnings`, `locations`, `overrides`, `sheet_contents`, or `schedules_info`. This preserves compatibility while giving agents a common `items` field.

`get_user_selection_in_revit` is also paginated. It returns only the current page in `selected_element_ids`, `items`, and `elements`, while `total_count`, `has_more`, and `next_offset` describe the whole Revit selection.

For dictionary-like results, `items` is an array of `{ "key": "...", "value": ... }`, and the original dictionary property contains the same page as an object.

For text-heavy journal results, the default and maximum page size is `204800` characters, approximately 200 KB of text.

## Journal Tool Check

`get_journal_entries_since` reads local Revit journal files from:

```text
%LOCALAPPDATA%\Autodesk\Revit\Autodesk Revit <version>\Journals\
```

Supported date formats:

```text
2026-08-28 21:17:00
2026-08-28T21:17:00
28.08.2026 21:17:00
```

Example:

```powershell
$body = @{
  jsonrpc = "2.0"
  id = 10
  method = "tools/call"
  params = @{
    name = "get_journal_entries_since"
    arguments = @{
      dateTime = "2026-08-28 21:17:00"
      endDateTime = "2026-08-28 21:19:00"
      limit = 204800
      offset = 0
    }
  }
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json; charset=utf-8" -Body $body |
  ConvertTo-Json -Depth 20
```

When analyzing journal results, focus first on user actions. Technical startup, shutdown, application loading and add-in events should be treated as supporting context unless they are the only events in the requested range.

The journal reader scans recent journal files by journal number, not only by file modification time. This keeps previous Revit session journals visible after a Revit restart. The reader is bounded by file count, file size, elapsed time and result size. If matched entries exceed one response, the tool returns a page of `entries`, also available as `items`, and sets `has_more = true` with `next_offset`. Call the same tool again with the returned `next_offset` to read the next page. `limit` is measured in characters, defaults to `204800`, and is capped at `204800`.

## MCP Tool Errors

Failed `tools/call` responses return `isError: true`. The text content starts with a short machine-readable code, and `structuredContent` contains the detailed error object.

Example:

```json
{
  "content": [
    {
      "type": "text",
      "text": "invalid_arguments: Required MCP tool argument is missing: categoryId"
    }
  ],
  "isError": true,
  "structuredContent": {
    "success": false,
    "toolName": "get_elements_by_category",
    "error": {
      "code": "invalid_arguments",
      "message": "Required MCP tool argument is missing: categoryId",
      "exceptionType": "KPLN_CoordiantorAI.ExternalAIModel.Mcp.RevitMcpToolArgumentException",
      "details": {
        "toolName": "get_elements_by_category",
        "argumentName": "categoryId",
        "expected": "required"
      }
    }
  }
}
```

Common error codes:

- `tool_not_found`
- `tool_not_implemented`
- `invalid_arguments`
- `tool_result_error`
- `active_document_missing`
- `ui_document_missing`
- `revit_busy`
- `external_event_raise_failed`
- `revit_api_context_error`
- `revit_api_error`

## WPF Tab Check

In the "Work with model" tab, ask:

```text
Call get_active_view_in_revit and show the result. Do not answer without calling the tool.
```

Then inspect:

```text
C:\Users\mtarchokov\AppData\Local\KPLN\CoordinatorAI\Diagnostics\<UserName>_diagnostic_YYYY-MM-DD.txt
```

Expected log markers:

```text
WPF chat routes tool through MCP. ToolName=get_active_view_in_revit
[MCP] tools/call begin. ToolName=get_active_view_in_revit, scenario=wpf_window
[MCP] tools/call end. ToolName=get_active_view_in_revit, Success=True, scenario=wpf_window
```

During request processing, the WPF tab shows short progress/status messages such as:

- `Анализирую запрос...`
- `Проверяю, нужны ли данные из Revit...`
- `Выполняю команду Revit...`
- `Получаю категории модели...`
- `Читаю журнал Revit...`
- `Обрабатываю результат...`
- `Формирую ответ...`

These messages are generated by the application code through `IModelProgressReporter`. They are not model chain-of-thought and do not expose raw JSON, tool arguments, HTTP details, stack traces, or MCP transport internals.

## WPF File Attachments

The paperclip button in the "Work with model" tab attaches only files explicitly selected by the user. A separate warning dialog is not displayed. The disclosure retained for documentation is:

> Будут доступны только файлы, которые вы сейчас выберете.
>
> Для анализа содержимое выбранных файлов может отправляться модели порциями до 200 КБ.
>
> Адрес модели: настроенный пользователем API endpoint или локальный сервер.
>
> Пути к файлам модели не передаются. Доступ действует до удаления файла или закрытия этого окна.

Supported files are DOCX, text, Markdown, CSV/TSV, JSON/JSONL, XML/YAML, logs, configuration files and common source-code formats. Legacy binary DOC files are not supported. A file is limited to 50 MB, and one WPF chat can hold no more than 10 attachments.

The model initially receives only the attachment `file_id`, file name, extension and size. The absolute filesystem path and file content are not placed in the initial prompt. Content is obtained later through these internal MCP tools:

- `get_attached_files` lists metadata for the current WPF attachment scope;
- `search_attached_file` searches text and returns matching line previews;
- `read_attached_file` reads at most 204800 characters per page and returns `has_more` and `next_offset`.

DOCX files are opened as read-only Open XML packages without starting Microsoft Word. Paragraphs, tables, headers, footers, footnotes and endnotes are converted to plain text. Embedded images, drawings, charts, equations and other non-text objects are not extracted or OCR-processed. Extracted Word text is limited to 10 million characters and cached only for the lifetime of the current WPF window, with a 20 million character total cache limit.

When the user asks to process a complete file, the model must continue reading pages until `has_more` is `false`. The files are opened read-only with sharing enabled. Removing an attachment or closing the WPF window revokes its in-memory `file_id` and scope.

The three attachment tools are omitted from external MCP `tools/list` responses and direct external calls are rejected. They are available only to the internal `wpf_window` scenario. File contents are treated as untrusted reference data, not as instructions for the model.

## WPF Image Attachments

The WPF chat accepts PNG, JPG/JPEG, WEBP and BMP images through the paperclip button. An image already copied to the Windows clipboard can be attached by pressing `Ctrl+V` while the message input is focused. Attached images are shown as removable thumbnails above the input.

Clipboard images are temporarily stored as PNG files under:

```text
%LOCALAPPDATA%\KPLN\CoordinatorAI\Attachments\<session-id>\
```

Only explicitly selected or pasted images are registered. Absolute paths, image bytes and Base64 data are not written to `ChatLogger` or `DiagnosticLogger`. Temporary clipboard files are deleted when the image is removed, after a successful answer, or when the WPF window closes.

One question can contain up to 5 images. An original image is limited to 15 MB. Before transmission, metadata is removed by re-encoding, the longest side is reduced to at most 2048 pixels, and the prepared image is limited to 8 MB. PNG is retained for screenshots when practical; oversized PNG data falls back to JPEG quality 85.

Images are model input rather than Revit API operations. They are sent in the current user message as OpenAI-compatible multimodal `image_url` data. Revit discovery and actions continue to use the shared MCP `tools/list` and `tools/call` architecture. After the final answer, the multimodal history entry is replaced with text so the same Base64 image is not resent in later questions.

The configured API model must support OpenAI-compatible vision input. If it rejects an image request, the WPF error points the user to check vision support. Text-only models cannot inspect attached images.

## WPF Request Cancellation

While a request is active, the `Отправить` button becomes `Отменить`. Pressing it cancels the current model HTTP request, stops subsequent model rounds and tool calls, removes an MCP tool call that is still waiting for Revit `ExternalEvent`, discards partial tool-chain messages, and restores the normal input state.

Revit API commands execute synchronously on Revit's main UI thread. Revit does not provide a safe general-purpose way to terminate arbitrary API code in the middle of such a call. If a tool is already inside synchronous Revit API execution, it finishes at the next safe return point; its result is discarded and no following tools are started. Long-running tools should therefore check cancellation cooperatively if finer-grained interruption is added later.

## Diagnostics

Check the allocated MCP ports:

```powershell
Get-NetTCPConnection -State Listen -ErrorAction SilentlyContinue |
  Where-Object { $_.LocalPort -ge 48731 -and $_.LocalPort -le 48799 }
```

Main diagnostic log:

```text
C:\Users\mtarchokov\AppData\Local\KPLN\CoordinatorAI\Diagnostics\<UserName>_diagnostic_YYYY-MM-DD.txt
```

MCP diagnostics are written to the same `DiagnosticLogger` file as the rest of the plugin diagnostics. MCP lines are marked with `[MCP]`; transport-specific lines can also contain `[MCP_TRANSPORT]`. Tool-call records include `InstanceId`, `RevitVersion`, and `ProcessId` so calls from simultaneous Revit processes can be distinguished.

This log contains useful MCP events: server start/stop, startup failures, `tools/call` begin/end, `tools/call` errors, invalid JSON, unsupported MCP methods, non-standard HTTP listener errors, and Revit API / `ExternalEvent` errors that pass through MCP.

It does not log low-level transport noise for every successful service request: accepted HTTP requests, successful `initialize`, `ping`, `tools/list`, response write sizes, response close events, request-body decoding details, or listener wait-loop iterations.

If the endpoint times out, check the last diagnostic log lines for startup failures, tool call start/end, and exceptions.

For JSON requests without an explicit `charset`, the transport first reads the body as UTF-8. This is the JSON default and matches normal MCP clients.

Windows PowerShell 5.1 can corrupt Cyrillic string request bodies before they reach the server when `charset` is omitted. Once the bytes are already `?`, the server cannot restore the original word. For Cyrillic arguments in Windows PowerShell 5.1, either specify UTF-8 in `ContentType`:

```powershell
Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json; charset=utf-8" -Body $body
```

or send explicit UTF-8 bytes:

```powershell
$bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
Invoke-RestMethod -Uri $endpoint -Method Post -ContentType "application/json; charset=utf-8" -Body $bytes
```

The transport still protects JSON request decoding with UTF-8 defaults and fallback encoding detection, but these low-level decisions are not written to the normal diagnostic log.

If the log contains:

```text
System.MissingMethodException: Method not found: JToken.ToString(Newtonsoft.Json.Formatting)
```

then Revit loaded a different `Newtonsoft.Json` version than the one used during compilation. The transport avoids `JToken.ToString(Formatting.None)` and uses `JsonConvert.SerializeObject(...)` for compatibility.

## Structured SQL History

The WPF "Work with model" tab writes structured request history to SQLite in parallel with the existing text `ChatLogger`. The text log has not been removed. External MCP hosts continue to use the existing MCP diagnostics and grouped `ChatLogger` records; they do not create `ModelChatRequests` rows.

Each Windows user receives a separate database. Its folder is configured in Settings -> External model -> Chat history database folder. If the field is empty, the default folder is:

```text
Z:\Отдел BIM\Тарчоков Мухамед\2. Плагины\1. Собственные плагины\0. Shared Plugins\Logs\BimHelper\1. Логи в форме БД\CoordinatorAI_History_<UserName>.db
```

The database and its schema are created automatically when the WPF tab opens. The implementation uses the existing `System.Data.SQLite` dependency and keeps five application tables:

### ModelChatSessions

One row represents one opened WPF chat session. It contains `SessionId`, `UserName`, `StartedAt`, `EndedAt`, `RevitVersion`, `RevitProcessId`, `PluginVersion`, `ModelName`, and `ConnectionType`.

### ModelChatRequests

One row represents one user question and its complete processing lifecycle. It contains:

- request identity and relation to the session: `RequestId`, `SessionId`;
- timing: `RequestTime`, `ResponseTime`, `DurationMs`;
- content and outcome: `UserQuestion`, `FinalAnswer`, `Status`, `ErrorMessage`;
- execution context: `Scenario`, `AiModelName`, `RevitVersion`, `RevitProcessId`, `ModelName`, `ViewName`;
- unique tools used by the request: `UsedToolNamesJson`;
- usage and calculated cost: `CacheHitTokens`, `CacheMissTokens`, `CompletionTokens`, `Cost`.

`Status` changes from `Processing` to `Completed`, `Cancelled`, or `Failed`. `Scenario` is currently `wpf_window`. `UsedToolNamesJson` is a JSON array of the MCP tools that were actually executed, in first-call order and without duplicate names. Tools skipped by the five-tools-per-model-response limit are not included.

Token values are accumulated across all model API rounds required to answer one user question. They are written only when every relevant online API response provides a complete usage breakdown. Missing usage data is stored as SQL `NULL`, not as zero. For a local model, `Cost` is zero; token fields can remain `NULL` when the local server does not report them.

For supported DeepSeek models, `Cost` is the calculated API consumption in US dollars for the whole request. The calculation separately prices cache-hit input, cache-miss input, and completion tokens, then sums all API rounds. It uses the requested billing model and the request start time, including the configured weekday peak/off-peak periods. This is an estimate based on the rates encoded in the plugin, not a provider invoice or bank charge. If the model is unknown or usage is incomplete, `Cost` remains `NULL`. Pricing changes must be synchronized with the official [DeepSeek pricing documentation](https://api-docs.deepseek.com/quick_start/pricing/).

### ModelAttachments

One row represents one file or image attached to a request. It contains `Id`, `RequestId`, `AttachmentType`, `FileName`, `FileExtension`, `ContentType`, `SizeBytes`, `Width`, and `Height`.

Only attachment metadata is stored. File contents, image bytes, Base64 data, and absolute local paths are not written to the database.

### ToolDefinitions

This reference table is synchronized from `RevitMcpToolRegistry` when the WPF tab opens and also includes the two read-only Autodesk Product Help tools under `OperationType=AutodeskHelp`. It contains `ToolId`, `ToolName`, `Description`, `OperationType`, `IsReadOnly`, `SupportsPagination`, `PaginationUnit`, and `DefaultPageSize`.

`OperationType` uses the registry area, for example `ViewContext`, `Visibility`, `Categories`, `Families`, `Parameters`, `Geometry`, `ModelInfo`, `Files`, `Worksets`, `Selection`, `Schedules`, `Journal`, or `Links`. Pagination metadata describes the unit and default page size used by each tool; unsupported values remain `NULL`.

### ModelAnswerFeedback

Each completed answer in the WPF "Work with model" tab has controls for a positive or negative rating and an optional comment. A negative rating can also include one reason: incorrect answer, incomplete answer, misunderstood question, Revit data failure, wrong Revit action, slow response, or other.

Feedback is stored in the same per-user SQLite database as the request history. `ModelAnswerFeedback` contains `FeedbackId`, `RequestId`, `Rating`, `ReasonCode`, and `Comment`. `RequestId` is unique, so changing a rating updates the existing feedback instead of creating a duplicate. Removing the selected rating deletes its feedback row.

Question text, final answer, model metadata, tool names, tokens, and cost are not duplicated in the feedback table. Reports can obtain them by joining `ModelAnswerFeedback.RequestId` to `ModelChatRequests.RequestId`.

Feedback has no `CreatedAt` or `UpdatedAt` columns. If the network database is unavailable, the plugin does not create a local feedback queue or write the feedback elsewhere. The WPF panel reports that the feedback could not be saved, while the technical exception is written to `DiagnosticLogger` as `MODEL_FEEDBACK.SAVE_ERROR` or `MODEL_FEEDBACK.DELETE_ERROR`.

### Time, Migration, and Failures

`StartedAt`, `EndedAt`, `RequestTime`, and `ResponseTime` are stored in Moscow time using the exact format `yyyy-MM-dd HH:mm:ss`. They contain no milliseconds, `T`, or `Z`. Existing UTC/ISO values are normalized automatically when the WPF tab next opens.

On database initialization, requests with missing `Cost` are recalculated when all three token fields, the model name, and request time are available. Rows without trustworthy usage data remain unchanged.

SQLite logging failures do not interrupt the user's model request. They are written to the regular `DiagnosticLogger` with events such as `MODEL_HISTORY.SESSION_START_ERROR`, `MODEL_HISTORY.SESSION_END_ERROR`, and `MODEL_HISTORY.REQUEST_SAVE_ERROR`.

## Revit API Constraint

External HTTP/MCP requests do not arrive in a valid Revit API context. Revit API calls must not run directly from the HTTP listener thread.

`RevitMcpExternalEventHandler` passes tool execution into Revit through `ExternalEvent`.

## Security

The current endpoint listens only on:

```text
127.0.0.1
```

Current WPF integration:

- the "Work with model" chat loads tool definitions through MCP `tools/list`;
- tool execution from the chat is routed through MCP `tools/call`;
- `ExternalModelWindow.xaml.cs` no longer contains a local `toolsArray` source of truth.

Log scenario markers:

- WPF tab records use `wpf_window`;
- external stdio proxy records use `external_mcp` by default;
- `DiagnosticLogger` writes the marker as `scenario=...`;
- `ChatLogger` writes the marker as `SCENARIO: ...`;
- external `ChatLogger` records are grouped: several consecutive MCP tool calls are written as one log entry with the called tool list;
- MCP diagnostic records are marked with `[MCP]` in the common `DiagnosticLogger` file.

Current risk handling:

- every registered tool has a `ReadOnly`, `UiChanging`, or `ModelChanging` classification;
- every MCP tool call is written to the common diagnostics with a scenario marker;
- the current registry contains read-only and UI-changing tools; no model-changing tool is registered;
- explicit confirmation for external UI-changing calls and a multi-request ExternalEvent queue are intentionally deferred and are not part of the current migration scope.

## Migration Status

The migration to one shared MCP architecture is complete. The WPF tab and external MCP hosts use the same registry, executor, Revit `ExternalEvent` path, result contracts, pagination rules, and diagnostics.

No mandatory MCP architecture changes remain. Before distributing a release, build the supported Revit configurations and run the endpoint, WPF, stdio proxy, pagination, UI-changing tool, restart, and diagnostics checks documented above.

## Recent Revit Operation Errors

The read-only MCP tool `get_recent_revit_failures` reads transient operation failures
(for example, "Cannot make type") captured by `FailuresProcessing`.
Use it when a user reports a failed operation or an error dialog.
`get_all_warnings_in_the_model` remains the tool for persistent model warnings.

Capture starts when the plugin loads. The observer copies detached values during
the event; it does not resolve failures, delete elements, change transaction
results, or dismiss Revit dialogs. Only failures reaching this event are observable:
this is not a universal reader for all Revit/Windows dialogs. An error that occurred
before capture started cannot be recovered by this tool.

### Session Storage and Cleanup

Local snapshots are stored in:

```text
%LOCALAPPDATA%\KPLN\CoordinatorAI\RevitFailures\
failures_<RevitVersion>_<PID>_<ProcessStartUtcTicks>_<SessionGuid>.jsonl
```

- One file belongs to one plugin capture session in one Revit process. Under normal
  operation this is one file per Revit launch, created only after the first failure.
- Each line is a JSON record. Documents in the same process share the file but have
  distinct session-local `document_id` values, even when their names are identical.
- The fast query source is memory. A background writer updates a bounded JSONL
  snapshot; this file is not a permanent error history.
- The file and memory records are removed on application closing/module shutdown.
  Writer and deletion are synchronized so a late write cannot recreate the file.
- After a crash, the next plugin launch removes orphan files. Both PID and process
  start time are checked. Files belonging to live processes, or whose ownership
  cannot be established, are preserved.
- Disk errors do not stop capture in memory. `file_write_error` explains failed
  writes; write/delete/capture problems also go to the existing DiagnosticLogger.
  No network drive or chat-history database is required.
- Retention is at most 200 records per document, 2000 records per process, and
  4 MiB of serialized records. A single event captures at most 200 failure messages.
  Identical messages within one callback are merged using `occurrences_in_event`;
  later callbacks are separate observations, not a count of distinct user actions.
- Text descriptions are limited to 8192 characters and each related-element list
  to 200 ids. Truncation is explicit in the returned record.
  `buffer_incomplete` and `dropped_count` report discarded records.

### Query Arguments and Results

Defaults: latest 10 observations for the active document, newest first.
The tool returns `failures`, `total_count`, `returned_count`, `limit`,
`offset`, `has_more`, `next_offset`, and `snapshot_sequence`.
The maximum page is 200 records and 200 KiB of compact UTF-8 result JSON
(the MCP envelope also includes the text representation).

Optional arguments:

- `document_id`: choose a document id returned in an earlier failure record.
- `all_documents: true`: query all captured documents in the selected Revit
  process. Cannot be combined with `document_id`.
- `dateTime`, `endDateTime`: inclusive time bounds; accepted forms include
  `2026-10-15 11:00:00`, `2026-10-15T11:00:00`, and ISO with `Z` or a UTC
  offset. Values without an offset mean Moscow time (UTC+03:00).
- `severity`: `Warning`, `Error`, or `DocumentCorruption`.
- `limit`: 1-200, default 10; `offset`: non-negative integer.
- `snapshot_sequence`: keep the value returned by the first page, together
  with the same scope/time filters, while following `next_offset`. This excludes
  newly arriving records from the paged selection. Buffer evictions can still
  change the retained selection; check `buffer_incomplete`.

Each record includes time in UTC and Moscow, document id/name/path, Revit version,
process/session identifiers, transaction name, failure-definition id, severity,
description and associated element ids. If an optional API property cannot be
read, the other fields remain available and `capture_incomplete`/`read_errors`
describe the missing data.

A recorded failure is historical evidence, not proof that a dialog is still open
or that the problem still exists. IDs may refer to elements removed by rollback.
An empty result does not prove there was no error: check document scope, capture
start, retention limits, and whether the message belongs to Revit's failure system.

### Modal Dialogs and Multiple Revit Processes

Reading this tool uses no Revit API calls or ExternalEvent. The HTTP listener
handles up to 32 simultaneous requests so an ordinary call waiting for Revit
does not block a separate HTTP query for captured failures. Other model-reading
tools still require Revit API context and may wait for the dialog to close.

The existing stdio bridge processes requests sequentially: if a previous request
is already waiting inside that bridge, a later request through the same bridge
will wait too. In that case query the selected process's HTTP endpoint separately
or close the dialog to finish the earlier request. The WPF window itself may also
be inaccessible while Revit is showing a modal dialog; recorded errors remain
available after it closes. This change does not alter the stdio bridge.

For multiple Revit processes, select the required instance in the external host.
A tool call reads only the process receiving it; `all_documents` does not mean
all running Revit versions. Active-document identity is cached on ViewActivated
and Idling. If it is unavailable (for example, during model opening), use
`all_documents: true`.

### Verification

1. Restart Revit after deploying the matching Debug2020/Debug2023/Debug2024 DLL.
2. Refresh/reconnect the external MCP host so its tools list includes the new tool.
3. Reproduce a known family error in a test model. Leave the dialog open to test
   independent HTTP access, or cancel the operation and query from the WPF chat:
   "Read the last Revit operation error and explain its message."
4. Query the endpoint for that Revit process (48731 is the first/default port;
   additional processes can use other ports from their instance registration):

```powershell
$endpoint = 'http://127.0.0.1:48731/mcp/'
$body = @{
    jsonrpc = '2.0'
    id = 1
    method = 'tools/call'
    params = @{
        name = 'get_recent_revit_failures'
        arguments = @{ all_documents = $true; limit = 10 }
    }
} | ConvertTo-Json -Depth 10
$bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
Invoke-RestMethod -Uri $endpoint -Method Post -ContentType 'application/json; charset=utf-8' -Body $bytes |
    ConvertTo-Json -Depth 20
```

5. Confirm the message, document and IDs match the reproduced error. Test the same
   model name in two documents/processes to check isolation.
6. Confirm the file appears only after an error and disappears on closing its
   Revit process; files of other running processes must remain.
7. Run `Tests\RunRevitFailureStoreTests.ps1` for automated retention, pagination,
   time filtering, invalid arguments, file lifecycle and HTTP concurrency checks.
   The HTTP test uses the production transport and a simulated blocked Revit
   handler. It does not replace reproducing the actual failure in Revit.

Json.NET runtime compatibility: Revit 2023/2024 can load Newtonsoft.Json 13.0.1,
although the project compiles against 13.0.4 (both have assembly version 13.0.0.0).
Failure snapshots use JsonConvert.SerializeObject instead of the newer
JToken.ToString(Formatting) overload. To test against the actual host library:

```powershell
.\Tests\RunRevitFailureStoreTests.ps1 -RuntimeJsonAssembly 'C:\Program Files\Autodesk\Revit 2024\Newtonsoft.Json.dll'
```
