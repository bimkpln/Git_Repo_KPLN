# Coordinator AI MCP

## Purpose

This project is moving to a single MCP-based architecture for Revit tools.

Both scenarios should use the same local Revit MCP server:

```text
Revit -> "Work with model" WPF tab -> Internal MCP Client -> Revit MCP Server -> Revit API

Claude/Cursor/VS Code/Codex -> MCP -> Revit MCP Server -> Revit API
```

MCP is the source of truth for Revit tool discovery and execution. If the internal LLM still needs OpenAI-compatible `tools`, that format is only an adapter for the model, not a separate Revit tool architecture.

## Current State

Implemented:

- shared MCP tool contracts;
- one shared MCP registry for all 54 Revit tools;
- one shared executor for all registered read-only and UI-changing tools;
- `ExternalEvent` wrapper for safe Revit API execution;
- local HTTP JSON-RPC endpoint inside Revit;
- internal MCP client used by the WPF "Work with model" tab for `tools/list` and `tools/call`;
- stdio proxy exe for external MCP hosts;
- consistent structured MCP errors and user-facing error messages;
- pagination for potentially large tool results;
- a limit of five real tool executions per model response in the WPF agent loop;
- shared `DiagnosticLogger` and grouped `ChatLogger` records with scenario markers;
- user-facing WPF progress statuses that do not expose model chain-of-thought.

Endpoint:

```text
http://127.0.0.1:48731/mcp/
```

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

The proxy does not call the Revit API directly. It forwards MCP JSON-RPC messages to the Revit HTTP endpoint:

```text
http://127.0.0.1:48731/mcp/
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

If the Revit MCP endpoint ever changes, pass it explicitly:

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
Agent -> stdio MCP -> RevitMcpStdioProxy.exe -> http://127.0.0.1:48731/mcp/ -> Revit
```

Before connecting an external agent:

- start Revit;
- open or create a Revit model;
- make sure the plugin is loaded;
- check that the local MCP endpoint responds.

Endpoint check:

```powershell
Invoke-WebRequest "http://127.0.0.1:48731/mcp/" -TimeoutSec 5
```

Expected response:

```json
{
  "name": "revit-mcp",
  "status": "running",
  "endpoint": "http://127.0.0.1:48731/mcp/",
  "tools": 54
}
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
- check `http://127.0.0.1:48731/mcp/`;
- check that `RevitMcpStdioProxy.exe` exists in `%LOCALAPPDATA%\KPLN\CoordinatorAI\MCP`;
- restart the external agent after changing its config;
- make sure the config points to the local installed proxy path, not to the repository `bin\Debug` path.

If rebuilding fails because `RevitMcpStdioProxy.exe` is locked, close the external agent that started it, then build again.

## Registered Tools

Currently registered tools:

- `get_active_view_in_revit`
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
- `get_phase_visibility_settings`
- `get_if_elements_pass_filter`
- `get_viewports_and_schedules_on_sheets`
- `get_schedules_info_and_columns`
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
Invoke-WebRequest "http://127.0.0.1:48731/mcp/" -TimeoutSec 5
```

Expected response:

```json
{
  "name": "revit-mcp",
  "status": "running",
  "endpoint": "http://127.0.0.1:48731/mcp/",
  "tools": 54
}
```

## tools/list Check

```powershell
$body = @{
  jsonrpc = "2.0"
  id = 1
  method = "tools/list"
  params = @{}
} | ConvertTo-Json -Depth 10

Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json" -Body $body |
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

Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json" -Body $body |
  ConvertTo-Json -Depth 20
```

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

Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json; charset=utf-8" -Body $body |
  ConvertTo-Json -Depth 30
```

The `allowed_by_phase_filter` field evaluates phase rules only. Other view settings can still hide an element.

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

Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json; charset=utf-8" -Body $body |
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

Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json; charset=utf-8" -Body $body |
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

## Diagnostics

Check the port:

```powershell
Get-NetTCPConnection -LocalPort 48731 -ErrorAction SilentlyContinue
```

Main diagnostic log:

```text
C:\Users\mtarchokov\AppData\Local\KPLN\CoordinatorAI\Diagnostics\<UserName>_diagnostic_YYYY-MM-DD.txt
```

MCP diagnostics are written to the same `DiagnosticLogger` file as the rest of the plugin diagnostics. MCP lines are marked with `[MCP]`; transport-specific lines can also contain `[MCP_TRANSPORT]`.

This log contains useful MCP events: server start/stop, startup failures, `tools/call` begin/end, `tools/call` errors, invalid JSON, unsupported MCP methods, non-standard HTTP listener errors, and Revit API / `ExternalEvent` errors that pass through MCP.

It does not log low-level transport noise for every successful service request: accepted HTTP requests, successful `initialize`, `ping`, `tools/list`, response write sizes, response close events, request-body decoding details, or listener wait-loop iterations.

If the endpoint times out, check the last diagnostic log lines for startup failures, tool call start/end, and exceptions.

For JSON requests without an explicit `charset`, the transport first reads the body as UTF-8. This is the JSON default and matches normal MCP clients.

Windows PowerShell 5.1 can corrupt Cyrillic string request bodies before they reach the server when `charset` is omitted. Once the bytes are already `?`, the server cannot restore the original word. For Cyrillic arguments in Windows PowerShell 5.1, either specify UTF-8 in `ContentType`:

```powershell
Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json; charset=utf-8" -Body $body
```

or send explicit UTF-8 bytes:

```powershell
$bytes = [System.Text.Encoding]::UTF8.GetBytes($body)
Invoke-RestMethod -Uri "http://127.0.0.1:48731/mcp/" -Method Post -ContentType "application/json; charset=utf-8" -Body $bytes
```

The transport still protects JSON request decoding with UTF-8 defaults and fallback encoding detection, but these low-level decisions are not written to the normal diagnostic log.

If the log contains:

```text
System.MissingMethodException: Method not found: JToken.ToString(Newtonsoft.Json.Formatting)
```

then Revit loaded a different `Newtonsoft.Json` version than the one used during compilation. The transport avoids `JToken.ToString(Formatting.None)` and uses `JsonConvert.SerializeObject(...)` for compatibility.

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
