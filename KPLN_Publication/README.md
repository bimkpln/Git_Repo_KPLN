# KPLN_Publication

## Сборка

Откройте `KPLN_Publication.sln` в Visual Studio:

- Revit 2020/2023/2024: `Debug2020/2023/2024` или `Revit2020/2023/2024`, платформа `x64`, .NET Framework 4.8.
- Revit 2026: `Debug2026` или `Revit2026`, платформа `x64`, .NET 8. Общие настройки подключаются из `..\build\KPLN.Revit2026.props`, как в BIMTools.

Нужны Revit API соответствующего года, KPLN_Loader, KPLN_Library_Forms и KPLN_Library_DBWorker. Для 2026 Debug-сборки используют библиотеки из `bin\Debug\2026`, release — из `bin\2026`. Путь Loader 2026 можно переопределить через свойство `KPLNLoaderPath`.

Обычный Build/Rebuild автоматически восстанавливает отсутствующий `project.assets.json` для 2026, в том числе при Batch Build из Visual Studio.

Debug DLL создаётся в `KPLN_Publication\bin\Debug\2026`. Конфигурация `Revit2026` сохраняет принятый путь выпуска: `Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader_v.2\Publication\2026`.

Рядом с DLL необходимы `formats.txt`, `itextsharp.dll` и `BouncyCastle.Cryptography.dll`; сборка копирует их автоматически. Для 2026 используется вариант BouncyCastle `net6.0`; iTextSharp остаётся версии 5.5.13.4.

## Проверка PDF на .NET 8

Из папки решения:

```powershell
dotnet run --project tests\Revit2026Smoke\Revit2026Smoke.csproj --configuration Release
```

Проверка создаёт тестовый PDF с вектором и изображением, выполняет преобразование цветов, обработку границ, объединение PDF и проверяет журналирование. Файлы остаются в папке вывода теста. Запуск Revit и реальная печать через PDFCreator в этот тест не входят.
