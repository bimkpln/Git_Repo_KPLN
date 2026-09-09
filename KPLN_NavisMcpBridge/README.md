# KPLN_NavisMcpBridge

Плагин Navisworks Manage 2020 и Python MCP-сервер для работы с Clash Detective.
Перенесён из `C:\Users\tkutsko\source\navis2020-mcp` с переименованием под нейминг KPLN.

## Структура

- `KPLN_NavisMcpBridge.sln` — решение Visual Studio, Debug/Release, x64.
- `addin/KPLN_NavisMcpBridge.csproj` — плагин .NET Framework 4.7.2.
- `server/navis_mcp_server.py` — MCP-сервер, транспорт stdio.
- `server/clash_rules.json` — существующий реестр правил, перенесён без изменений.
- Три `Autodesk.Navisworks.*.dll` в корне — исходные ссылки API для сборки; в bundle не копируются.

## Сборка

Откройте решение в Visual Studio или выполните из папки проекта:

```powershell
dotnet msbuild .\KPLN_NavisMcpBridge.sln -restore -t:Rebuild -p:Configuration=Release -p:Platform=x64
```

DLL создаётся в `addin\bin\x64\Release\KPLN_NavisMcpBridge.dll`.
Как в исходном проекте, обычная сборка автоматически устанавливает bundle в
`%ProgramData%\Autodesk\ApplicationPlugins\KPLN_NavisMcpBridge.bundle`.
Для сборки комплекта в папке проекта без установки используйте:

```powershell
$bundlePath = Join-Path $PWD 'artifacts\KPLN_NavisMcpBridge.bundle\'
dotnet msbuild .\KPLN_NavisMcpBridge.sln -restore -t:Rebuild -p:Configuration=Release -p:Platform=x64 "-p:BundleDir=$bundlePath"
```

Bundle содержит `PackageContents.xml`, `Contents\KPLN_NavisMcpBridge.dll` и
`Contents\Newtonsoft.Json.dll`. Перед заменой установленного плагина закройте
Navisworks. При переходе на новое имя перенесите старую `NavisMcpBridge.bundle`
за пределы `ApplicationPlugins`, чтобы Navisworks загружал одну версию плагина.

## Запуск

1. Установите bundle и откройте Navisworks Manage 2020.
2. Нажмите **KPLN MCP Bridge** в надстройках. Первый клик запускает HTTP-мост
   `http://127.0.0.1:8765`, повторный останавливает его.
3. Для Python окружения MCP-клиента установите зависимости
   `python -m pip install -r server\requirements.txt`.
4. В существующей конфигурации MCP-клиента сохраните используемый Python и замените
   путь скрипта на
   `X:\BIM\5_Scripts\Git_Repo_KPLN\KPLN_NavisMcpBridge\server\navis_mcp_server.py`.

Плагин поднимает HTTP-мост внутри Navisworks. Python MCP-сервер запускается отдельно
MCP-клиентом, как в исходном проекте. Имя сервера при инициализации —
`KPLN_NavisMcpBridge`; имена и аргументы девяти MCP-инструментов сохранены.

## Объём переноса

Изменены имена решения, проекта, сборки, пространств имён, плагина и bundle,
название кнопки и MCP-сервера. Абсолютный путь вывода заменён относительным `bin\`;
путь `BundleDir` допускает переопределение при сборке.
Обработка коллизий, HTTP-маршруты, порт, алгоритмы Python, зависимости, реестр правил
и идентификаторы ProductCode/UpgradeCode сохранены.

Проверены сборка Release x64, синтаксис Python, соответствие исходному коду после
обратной подстановки имён и комплектность bundle. Запуск новой DLL внутри Navisworks
при переносе не выполнялся.
