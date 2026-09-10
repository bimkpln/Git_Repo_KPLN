# Navisworks 2020 → MCP мост (Clash Detective)

Статус: **v0.3** — теперь как проект Visual Studio (`KPLN_NavisMcpBridge.sln`).
Собран и переписан по итогам двух раундов декомпиляции ваших реальных DLL —
сначала `Autodesk.Navisworks.Api.dll` и `Autodesk.Navisworks.Interop.ComApi.dll`,
потом ещё и `Autodesk.Navisworks.ComApi.dll` (третья сборка, которую вы сами
нашли и докинули). Сети в песочнице нет (pip/npm/apt заблокированы), поэтому
готового декомпилятора поставить было нельзя — написал свой минимальный
парсер метаданных ECMA-335 на чистом Python и вытащил реальные имена
классов/методов/сигнатур прямо из ваших DLL. Ошибка сборки, которую вы
прислали (`RunTest` требует аргументы), тоже отсюда — я изначально не
досмотрел его параметры, теперь поправлено и перепроверено.

## Как открыть в Visual Studio

Просто откройте `KPLN_NavisMcpBridge.sln` — там один проект (`addin\KPLN_NavisMcpBridge.csproj`).
**F5 запускает Navisworks и сразу цепляет отладчик** (настроено в
`addin\KPLN_NavisMcpBridge.csproj.user`) — можно ставить breakpoint'ы прямо в
`ClashService.cs`. После каждой сборки (в т.ч. по F5) DLL и манифест сами
копируются в bundle-папку `%ProgramData%\Autodesk\ApplicationPlugins\KPLN_NavisMcpBridge.bundle\`
(таргет `DeployBundle` в .csproj) — руками копировать не нужно.

Если Copy в `DeployBundle` падает с "отказано в доступе" — запустите Visual
Studio от администратора (первое создание папки в `%ProgramData%\Autodesk\...`
иногда требует прав, дальше обычно уже нет).

Если F5 не запускает Navisworks — я предположил, что exe называется
`Roamer.exe` (историческое имя для Navisworks), но **это не проверено**
(это не .NET-сборка, декомпилировать нечем). Гляньте, на что ссылается ваш
ярлык "Navisworks Manage 2020" (правой кнопкой → Свойства → Объект) и
поправьте `StartProgram` в `KPLN_NavisMcpBridge.csproj.user`.

## Три DLL нужны в корне подключённой папки

Сейчас у вас там `Autodesk.Navisworks.Api.dll` и `Autodesk.Navisworks.ComApi.dll`
— не хватает третьей, **`Autodesk.Navisworks.Interop.ComApi.dll`** (её вы
раньше присылали мне в чат, 310 072 байт) — она держит все `Inw*`-интерфейсы
(`InwOpState10`, `InwOclClashTest` и т.д.), без неё сборка не найдёт эти
типы. Скопируйте её из `C:\Program Files\Autodesk\Navisworks Manage 2020\`
туда же, где лежат две другие (`X:\BIM\5_Scripts\Git_Repo_KPLN\KPLN_NavisMcpBridge\`).

## Главное открытие раунда 1: в Navisworks 2020 нет managed Clash API

`Autodesk.Navisworks.Api.Clash` (управляемый API для Clash Detective)
появился в SDK **позже** 2020 года. `Document.Clash` в вашей версии
возвращает `IDocumentClash` — пустой интерфейс-заглушку без единого метода.
Так же пусты `IDocumentTimeliner` и `IDocumentTakeoff`. Реальный Clash
Detective доступен только через **COM API** (`Interop.ComApi`).

## Главное открытие раунда 2: ComApiBridge всё-таки существует

Сначала я не нашёл публичный класс `ComApiBridge` (только внутренний,
но тоже публичный `Autodesk.Navisworks.Internal.ApiImplementation.ComApi.Bridge`
внутри `Api.dll`) и работал через него. Когда вы прислали третью DLL —
`Autodesk.Navisworks.ComApi.dll` — в ней нашёлся настоящий, "канонiчный"
`Autodesk.Navisworks.Api.ComApi.ComApiBridge` со строго типизированным
`State` (сразу `InwOpState10`, без каста) и, что особенно ценно,
`ComApiBridge.ToModelItem(InwOaPath)` — конвертирует путь клэша в обычный
managed `ModelItem`. Код теперь использует именно его вместо ручного обхода
COM-дерева узлов.

Цепочка вызовов в `addin/ClashService.cs`:

```
ComApiBridge.State                        // сразу InwOpState10, каста не нужно
  .Plugins()                              // InwPluginsColl (метод, не свойство!)
    → перебор, cast к InwOpClashElement   // корневой объект Clash Detective
      .Tests()                            // InwClashTestsColl (метод!)
        → InwOclClashTest                 // .name, .status, .RunTest(0,null), .results()
          .results()                      // InwTestResultsColl (метод!)
            → InwOclTestResult            // .name, .status, .distance, .Path1/.Path2, .Comments()
              → ComApiBridge.ToModelItem(Path1/Path2) → ModelItem.AncestorsAndSelf
```

## Важная деталь стиля этого COM API

Часть членов интерфейса — это методы БЕЗ префикса `get_`, хотя выглядят
как свойства: `Plugins()`, `Tests()`, `results()`, `Comments()` — их нужно
вызывать со скобками. Другая часть — настоящие C#-свойства: `.name`,
`.status`, `.distance`, `.Path1`, `.Path2`. Перепутать легко, в коде это
откомментировано на каждом вызове.

## Что подтверждено декомпиляцией (можно доверять)

- `InwOclClashTest.RunTest(int test_ndx, object ovProgress)` — **не** без
  параметров, как я думал сначала (отсюда была ошибка сборки CS7036).
  Смысл `test_ndx` на объекте-тесте (а не коллекции) неясен из одних
  метаданных — похоже на неиспользуемый параметр общего шаблона; вызываем
  как `RunTest(0, null)`.
- Статусы теста: `nwEClashTestStatus` = NEW / OLD / PARTIAL / OK.
- Статусы результата: `nwETestResultStatus` = NEW / ACTIVE / APPROVED /
  RESOLVED / REVIEWED / UNMERGED (у результата, не у теста).
- У тестов и результатов **нет Guid** — только `.name` (строка). MCP-тулы
  идентифицируют тест/результат по имени.
- `ComApiBridge.ToModelItem(path)` → обычный `ModelItem`, дальше стандартный
  managed API (`.AncestorsAndSelf.Select(a => a.DisplayName)`).
- Комментарии к результату: `result.Comments()` → `InwCommentsColl`, каждый
  элемент — `InwOpComment` с `.Body`, `.User`, `.Date`.
- Коллекции (`InwClashTestsColl`, `InwTestResultsColl`, `InwPluginsColl`,
  `InwCommentsColl`) реализуют `InwCollBase` → `System.Collections.IEnumerable`
  — обычный `foreach` работает.

## Что осталось непроверенным (нужен реальный запуск)

1. **Как находится сам элемент Clash Detective среди `State.Plugins()`.**
   В 2020 нет метода `GetPlugin(id)` по строковому имени — код перебирает
   все плагины документа и берёт первый, что приводится к
   `InwOpClashElement`. Должно работать независимо от точного имени
   плагина, но не проверено запуском.
2. **Добавление комментария** (`SetResultStatus` с непустым `comment`) —
   не реализовано, вернёт 501. Нужен `State.ObjectFactory(...)`, чей
   enum-параметр статически не виден.
3. **Экспорт отчёта** — не реализован, вернёт 501. Вероятный путь —
   `State.DriveIOPlugin(pluginName, filePath, options)`, но точное имя
   IO-плагина отчёта Clash Detective не определить статически.

Пункт 1 — единственное, что может помешать базовому сценарию (список
тестов / результаты / запуск теста / статус). 2 и 3 — сверх MVP.

## Что нужно от вас дальше

1. Докиньте `Autodesk.Navisworks.Interop.ComApi.dll` в подключённую папку
   (см. выше).
2. Откройте `KPLN_NavisMcpBridge.sln` в Visual Studio, соберите (или сразу F5).
3. Откройте Navisworks 2020 с моделью и тестами Clash Detective (если
   запускали не через F5 — вручную), нажмите "MCP Bridge" на вкладке
   Add-ins.
4. Запустите `server/navis_mcp_server.py` (нужен `pip install -r
   server/requirements.txt`) и позовите `list_clash_tests` — пришлите
   результат или текст ошибки, доведём до рабочего состояния оставшиеся
   пункты.

## Архитектура

```
Navisworks Manage 2020 (процесс)
  └─ addin/KPLN_NavisMcpBridge.dll   — плагин AddInPlugin, HttpListener на 127.0.0.1:8765
         │  JSON по HTTP (только localhost)
         ▼
server/navis_mcp_server.py    — MCP-сервер (stdio), тулы → HTTP-запросы к аддину
         │  MCP (stdio)
         ▼
      Codex / другой MCP-клиент
```

## MCP-тулы (server/navis_mcp_server.py)

- `list_clash_tests()` — имя, статус, число результатов по каждому тесту
- `get_clash_test_results(test_name)` — результаты теста
- `run_clash_test(test_name)` — пересчитать тест
- `set_clash_result_status(test_name, result_name, status, comment="")` —
  статус (New/Active/Approved/Resolved/Reviewed); `comment` пока не
  реализован
- `export_clash_report(test_name, format="html")` — пока не реализовано
