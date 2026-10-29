using KPLN_Library_Schedules.Common.Schedules;
using Autodesk.Revit.DB;
using KPLN_ViewsAndLists_CopyByLevels.Common;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace KPLN_ViewsAndLists_CopyByLevels.Forms.Entities
{
    public sealed class ElementChoice
    {
        public ElementId Id { get; set; }
        public string Name { get; set; }
        public override string ToString() => Name;
    }

    public sealed class LevelRow : ScheduleLevelRow { }

    public static class CopyActions
    {
        public const string NewPlan = "Новый план уровня";
        public const string Shared = "Общий вид";
        public const string Duplicate = "Независимая копия";
        public const string Replace = "Другой существующий вид";
        public const string Skip = "Не размещать";
    }

    public sealed class ContentRow : ObservableObject
    {
        private string _action;
        private ElementChoice _replacement;
        public Viewport Viewport { get; set; }
        public View View { get; set; }
        public string Name => View.Name;
        public string Kind => View.ViewType.ToString();
        public List<string> Actions { get; set; }
        public List<ElementChoice> Replacements { get; set; }
        public string Action { get => _action; set => Change(ref _action, value); }
        public ElementChoice Replacement { get => _replacement; set => Change(ref _replacement, value); }
    }

    public sealed class PlanDetailsRow : ObservableObject
    {
        private const string Copy = "Копия";
        private string _action = Copy;

        public string Name => "Независимое оформление планов";
        public string Kind => "Оформление";
        public string Description => "Тексты, линии, области и общие аннотации новых планов. Марки и размеры оформить отдельно.";
        public IReadOnlyList<string> Actions { get; } = new[] { Copy, CopyActions.Skip };
        public List<ElementChoice> Replacements { get; } = new List<ElementChoice>();
        public ElementChoice Replacement { get => null; set { } }
        public string Action
        {
            get => _action;
            set
            {
                if (value == Copy || value == CopyActions.Skip)
                    Change(ref _action, value);
            }
        }
        public bool CopyEnabled { get => Action == Copy; set => Action = value ? Copy : CopyActions.Skip; }
    }

    public sealed class ScheduleRow : ScheduleItem
    {
        public List<ScheduleSheetInstance> Instances { get; set; }
    }

    public sealed class PreviewRow : ObservableObject
    {
        private string _number;
        private string _name;
        private string _viewName;
        public LevelRow Level { get; set; }
        public int Index { get; set; }
        public string LevelName => Level.Name;
        public string Number { get => _number; set => Change(ref _number, value); }
        public string Name { get => _name; set => Change(ref _name, value); }
        public string ViewName { get => _viewName; set => Change(ref _viewName, value); }
        public string Filters { get; set; }
        public string Status { get; set; }
        public bool SkipExisting { get; set; }
    }

    public sealed class CopyByLevelsM
    {
        public Document Document { get; set; }
        public ViewSheet SourceSheet { get; set; }
        public ViewPlan MainPlan { get; set; }
        public ObservableCollection<LevelRow> Levels { get; } = new ObservableCollection<LevelRow>();
        public ObservableCollection<ContentRow> Contents { get; } = new ObservableCollection<ContentRow>();
        public ObservableCollection<ScheduleRow> Schedules { get; } = new ObservableCollection<ScheduleRow>();
        public ObservableCollection<PreviewRow> Preview { get; } = new ObservableCollection<PreviewRow>();
        public PlanDetailsRow PlanDetails { get; } = new PlanDetailsRow();
        public IEnumerable<object> ContentsForDisplay
        {
            get
            {
                foreach (ContentRow row in Contents)
                    yield return row;
                yield return PlanDetails;
            }
        }
        public bool CopyPlanDetails { get => PlanDetails.CopyEnabled; set => PlanDetails.CopyEnabled = value; }
        public bool CopySheetAnnotations { get; set; } = true;
        public bool SkipExisting { get; set; } = true;
    }
}
