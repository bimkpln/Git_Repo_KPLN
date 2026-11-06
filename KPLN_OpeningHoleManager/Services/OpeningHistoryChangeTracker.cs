using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using KPLN_OpeningHoleManager.Core;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_OpeningHoleManager.Services
{
    /// <summary>Индекс зависимостей карточек. В DocumentChanged не читает историю и геометрию.</summary>
    internal sealed class OpeningHistoryChangeTracker
    {
        // Создаются при первом событии в контексте Revit, а не при конструировании WPF.
        private static ElementFilter _levelFilter;
        private static ElementFilter _familyFilter;
        private static ElementFilter LevelFilter => _levelFilter ?? (_levelFilter = new ElementClassFilter(typeof(Level)));
        private static ElementFilter FamilyFilter => _familyFilter ?? (_familyFilter = new ElementClassFilter(typeof(Family)));

        private readonly Dictionary<string, OpeningHistoryRow> _rows = new Dictionary<string, OpeningHistoryRow>();
        private readonly Dictionary<string, HashSet<string>> _dependents = new Dictionary<string, HashSet<string>>();
        private readonly HashSet<string> _levels = new HashSet<string>();
        private readonly HashSet<string> _viewDependencies = new HashSet<string>();


        internal void Clear()
        {
            _rows.Clear();
            _dependents.Clear();
            _levels.Clear();
            _viewDependencies.Clear();
        }

        internal void Remove(string id)
        {
            OpeningHistoryRow old;
            if (!_rows.TryGetValue(id, out old))
                return;

            foreach (string dependency in old.DependencyIds)
            {
                HashSet<string> dependents = _dependents[dependency];
                dependents.Remove(id);

                if (dependents.Count == 0)
                    _dependents.Remove(dependency);
            }

            _rows.Remove(id);
        }

        internal void Set(OpeningHistoryRow row)
        {
            Remove(row.ElementId);
            _rows[row.ElementId] = row;

            foreach (string dependency in row.DependencyIds)
            {
                HashSet<string> dependents;
                if (!_dependents.TryGetValue(dependency, out dependents))
                {
                    dependents = new HashSet<string>();
                    _dependents.Add(dependency, dependents);
                }

                dependents.Add(row.ElementId);
            }
        }

        internal void SetLevels(IEnumerable<string> ids)
        {
            _levels.Clear();
            _levels.UnionWith(ids);
        }

        internal void SetViewDependencies(IEnumerable<ElementId> ids)
        {
            _viewDependencies.Clear();
            _viewDependencies.UnionWith(ids.Where(OpeningHistoryService.IsValidId).Select(id => id.ToString()));
        }


        internal OpeningHistoryChanges Read(Document document, DocumentChangedEventArgs args, bool onlyActiveView)
        {
            var changes = new OpeningHistoryChanges
            {
                CancelAutoCache = args.Operation != UndoOperation.TransactionCommitted,
            };

            // Для основ, типов, семейств и уровней достаточно пересечения с сохранённым индексом.
            // Чужие изменённые элементы не раскрываем через Document.GetElement.
            foreach (ElementId id in args.GetModifiedElementIds().Concat(args.GetDeletedElementIds()))
            {
                string key = id.ToString();
                if (_rows.ContainsKey(key))
                    changes.Add(id);

                HashSet<string> dependents;
                if (_dependents.TryGetValue(key, out dependents))
                {
                    foreach (string openingId in dependents)
                        changes.Add(_rows[openingId].RevitId);
                }

                changes.RefreshLevels |= _levels.Contains(key);
                changes.RefreshVisibility |= onlyActiveView && _viewDependencies.Contains(key);
            }


            // Категории и класс отсекаются внутри Revit до создания управляемых объектов.
            foreach (ElementId id in args.GetAddedElementIds(OpeningHistoryService.CandidateFilter))
            {
                FamilyInstance opening = document.GetElement(id) as FamilyInstance;
                if (!OpeningHistoryService.IsOpening(opening))
                    continue;

                changes.Add(id);
                if (!changes.CancelAutoCache)
                    changes.Created.Add(opening);
            }

            foreach (ElementId id in args.GetModifiedElementIds(OpeningHistoryService.CandidateFilter))
            {
                if (!changes.Elements.ContainsKey(id.ToString()) && OpeningHistoryService.IsOpening(document.GetElement(id)))
                    changes.Add(id);
            }

            changes.RefreshLevels |= args.GetAddedElementIds(LevelFilter).Count > 0
                || args.GetModifiedElementIds(LevelFilter).Count > 0;


            // Семейство могло получить подходящее имя: его экземпляров ещё нет в индексе.
            // Это редкий случай, в котором безопаснее заново собрать поддерживаемые отверстия.
            foreach (ElementId id in args.GetModifiedElementIds(FamilyFilter))
            {
                if (!_dependents.ContainsKey(id.ToString())
                    && OpeningHistoryService.IsOpeningFamily((document.GetElement(id) as Family)?.Name))
                {
                    changes.FullRefresh = true;
                    break;
                }
            }

            changes.RefreshVisibility |= onlyActiveView && (changes.Elements.Count > 0 || changes.FullRefresh);
            return changes;
        }
    }
}
