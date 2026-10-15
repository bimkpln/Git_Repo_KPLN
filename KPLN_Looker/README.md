# KPLN_Looker — версии Revit

Revit 2020/2023/2024 используют прежние конфигурации `Debug20xx` / `Revit20xx`, x64, .NET Framework 4.8. Для Revit 2026 добавлены `Debug2026` / `Revit2026`, x64, .NET 8 по образцу KPLN_Publication, с общим `build/KPLN.Revit2026.props`.

## Сборка

Откройте `KPLN_Looker.sln`. Все проекты сохранены в решении, но явные ProjectDependencies удалены. Looker использует готовые DLL через Reference/HintPath; библиотеки обновляются отдельно. При сборке всего решения участвуют проекты с включённой галочкой Build в Configuration Manager; для сборки только Looker выбирайте его проект или соответствующие галочки.

Необходимы установленный Revit соответствующего года и готовые KPLN_Loader, KPLN_Library_Forms, KPLN_Library_DBWorker, KPLN_Library_Bitrix24Worker того же года. Для 2026 используются .NET 8 сборки: Debug — `bin\Debug\2026`, release — `bin\2026`. Путь к Loader можно переопределить свойством `KPLNLoaderPath`.

Конфигурации 2026 автоматически восстанавливают отсутствующий `project.assets.json` при Build/Rebuild, в том числе в Visual Studio Batch Build. Сохраняются явные списки исходников, XAML и ресурсов, а также существующий AssemblyInfo.

Debug Looker: `C:\KPLN_Debug\KPLN_Looker\Debug\2026`. Release Looker: `Z:\Отдел BIM\03_Скрипты\09_Модули_KPLN_Loader_v.2\Looker\2026`. Release-сборка без переопределения `OutputPath` пишет в папку выпуска. Символ `REVIT` сохранён вместе с `Revit2026`.

Tools использует SQLite 2.0.3 `netstandard2.1`, Newtonsoft.Json 13.0.3 `net6.0` и существующую RevitServerAPILib. Их предоставление в среде Loader требуется, как и для остальных общих зависимостей с `Private=False`. Копировать DLL Revit API рядом с модулем не нужно.

## Проверка в Revit

После установки согласованного комплекта DLL проверить загрузку модуля, переключение видов и автосохранение, синхронизацию с автопроверками связей/рабочих наборов/осей, отслеживание помещений, обработку загрузки семейств и закрытие видов. Эти обработчики могут сохранять модели, изменять Revit.ini и отправлять штатные уведомления в Bitrix; тестировать на согласованной тестовой модели.

Сборка вне Revit проверяет совместимость API, но не заменяет проверку этих событий в запущенном Revit 2026.
