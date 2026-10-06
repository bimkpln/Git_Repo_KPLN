// Синтетический API проверяет контракты адаптера, но не подменяет интеграционный прогон в Revit.
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using KPLN_RevitMcpBridge.Server;

namespace Autodesk.Revit.Exceptions
{
    public class ApplicationException : Exception { public ApplicationException(string message) : base(message) { } }
    public class InvalidOperationException : ApplicationException { public InvalidOperationException() : base("rotation unsupported") { } }
}

namespace Autodesk.Revit.DB
{
    public class ElementId : IEquatable<ElementId>
    {
        public long Value;
        public ElementId(long value) { Value = value; }
        public bool Equals(ElementId other) => other != null && other.Value == Value;
        public override bool Equals(object obj) => Equals(obj as ElementId);
        public override int GetHashCode() => Value.GetHashCode();
        public static bool operator ==(ElementId a, ElementId b) => Equals(a, b);
        public static bool operator !=(ElementId a, ElementId b) => !Equals(a, b);
    }
    public enum BuiltInCategory { OST_IOSModelGroups = -1, OST_TitleBlocks = -2 }
    public enum CategoryType { Model, Annotation }
    public class Category
    {
        public ElementId Id;
        public CategoryType CategoryType = CategoryType.Model;
        public static Category GetCategory(Document doc, ElementId id) => doc.Categories.FirstOrDefault(c => c.Id == id);
    }
    public class WorksetId { public int IntegerValue; }
    public class XYZ { public double X, Y, Z; }
    public class Location { }
    public class LocationPoint : Location
    {
        public XYZ Point = new XYZ();
        public bool Unsupported;
        public double Angle;
        public double Rotation => Unsupported ? throw new Exceptions.InvalidOperationException() : Angle;
    }
    public class LocationCurve : Location { public Curve Curve; }
    public class Curve
    {
        public bool IsBound = true;
        public double Length = 1;
        public XYZ GetEndPoint(int index) => new XYZ { X = index };
    }
    public class Element
    {
        public ElementId Id;
        public string UniqueId;
        public string Name;
        public Category Category;
        public ElementId GroupId = new ElementId(-1), OwnerViewId = new ElementId(-1), LevelId = new ElementId(-1);
        public WorksetId WorksetId = new WorksetId();
        public Location Location;
        public bool Hidden, Supported = true, ThrowIntersection;
        public HashSet<ElementId> Intersections = new HashSet<ElementId>();
        public bool IsHidden(View view) => Hidden;
    }
    public class ElementType : Element { }
    public class Group : Element
    {
        public Element GroupType = new Element { Id = new ElementId(100), UniqueId = "type" };
        public bool IsAttached;
        public ElementId AttachedParentId = new ElementId(-1);
        public List<ElementId> Members = new List<ElementId>(), AttachedTypes = new List<ElementId>();
        public IList<ElementId> GetMemberIds() => Members;
        public IList<ElementId> GetAvailableAttachedDetailGroupTypeIds() => AttachedTypes;
        public IList<ElementId> GetShownAttachedDetailGroupTypeIds(View view) => AttachedTypes;
    }
    public class Document
    {
        public bool IsFamilyDocument, IsWorkshared;
        public bool IsLinked, IsModified;
        public string Title = "Synthetic test document";
        public int ExportCalls;
        public Dictionary<long, string> ExportModes = new Dictionary<long, string>();
        public PDFExportOptions LastOptions;
        public List<Element> Elements = new List<Element>();
        public List<Category> Categories = new List<Category>();
        public Element GetElement(ElementId id) => Elements.FirstOrDefault(e => e.Id == id);
        public Element GetElement(string id) => Elements.FirstOrDefault(e => e.UniqueId == id);
        public bool Export(string folder, IList<ElementId> ids, PDFExportOptions options)
        {
            ExportCalls++;
            LastOptions = options;
            string mode;
            ExportModes.TryGetValue(ids.Single().Value, out mode);
            if (mode == "throw") throw new Exceptions.ApplicationException("synthetic export failure");
            if (mode == "false") return false;
            if (mode == "missing") return true;
            System.IO.File.WriteAllText(System.IO.Path.Combine(folder, options.FileName + ".pdf"), mode == "invalid" ? "invalid" : "%PDF-test fixture, not a real rendered PDF");
            return true;
        }
    }
    public enum ExportPaperFormat { Default }
    public enum ColorDepthType { Color }
    public class PDFExportOptions : IDisposable
    {
        public bool Combine, HideCropBoundaries, HideScopeBoxes, HideReferencePlane, HideUnreferencedViewTags, StopOnError;
        public string FileName;
        public ExportPaperFormat PaperFormat;
        public ColorDepthType ColorDepth;
        public void Dispose() { }
    }
    public enum ViewType { FloorPlan }
    public enum DetailLevel { Fine }
    public enum DisplayStyle { Wireframe }
    public class View : Element
    {
        public ViewType ViewType;
        public DetailLevel DetailLevel;
        public DisplayStyle DisplayStyle;
        public int Scale = 100;
        public bool IsTemplate, Valid = true;
        public ElementId ViewTemplateId = new ElementId(-1);
        public List<ElementId> FilterIds = new List<ElementId>();
        public HashSet<ElementId> Candidates = new HashSet<ElementId>(), HiddenCategories = new HashSet<ElementId>();
        public ICollection<ElementId> GetFilters() => FilterIds;
        public bool IsTemporaryHideIsolateActive() => false;
        public bool GetFilterVisibility(ElementId id) => false;
        public bool GetIsFilterEnabled(ElementId id) => true;
        public WorksetVisibility GetWorksetVisibility(WorksetId id) => WorksetVisibility.UseGlobalSetting;
        public bool GetCategoryHidden(ElementId id) => HiddenCategories.Contains(id);
    }
    public class ViewSheet : View
    {
        public string SheetNumber = "1";
        public bool IsPlaceholder;
        public bool CanBePrinted = true;
        public List<ElementId> Viewports = new List<ElementId>();
        public IList<ElementId> GetAllViewports() => Viewports;
    }
    public class Viewport : Element
    {
        public ElementId ViewId;
        public XYZ GetBoxCenter() => new XYZ();
    }
    public class ScheduleSheetInstance : Element { public ElementId ScheduleId; }
    public class ElementFilter : IDisposable
    {
        public virtual bool PassesFilter(Element element) => true;
        public void Dispose() { }
    }
    public class ParameterFilterElement : Element
    {
        public List<ElementId> Categories = new List<ElementId>();
        public IList<ElementId> GetCategories() => Categories;
        public ElementFilter GetElementFilter() => new ElementFilter();
    }
    public class SelectionFilterElement : Element
    {
        public List<ElementId> Members = new List<ElementId>();
        public ICollection<ElementId> GetElementIds() => Members;
    }
    public class ElementMulticategoryFilter : ElementFilter
    {
        private readonly ICollection<ElementId> _categories;
        public ElementMulticategoryFilter(ICollection<ElementId> categories) { _categories = categories; }
        public override bool PassesFilter(Element element) => element.Category != null && _categories.Contains(element.Category.Id);
    }
    public class ElementIntersectsFilter : ElementFilter
    {
        public static bool IsElementSupported(Element element) => element.Supported;
        public static bool IsCategorySupported(Element element) => element.Category != null;
    }
    public class ElementIntersectsElementFilter : ElementIntersectsFilter
    {
        private readonly Element _source;
        public ElementIntersectsElementFilter(Element source)
        {
            if (source.ThrowIntersection) throw new Exceptions.ApplicationException("synthetic API failure");
            _source = source;
        }
        public override bool PassesFilter(Element element) => _source.Id == element.Id || _source.Intersections.Contains(element.Id);
    }
    public class FilteredElementCollector : IEnumerable<Element>, IDisposable
    {
        private IEnumerable<Element> _elements;
        public FilteredElementCollector(Document doc) { _elements = doc.Elements; }
        public FilteredElementCollector(Document doc, ICollection<ElementId> ids) { _elements = doc.Elements.Where(e => ids.Contains(e.Id)); }
        public FilteredElementCollector(Document doc, ElementId view) { _elements = doc.Elements.Where(e => ((View)doc.GetElement(view)).Candidates.Contains(e.Id)); }
        public static bool IsViewValidForElementIteration(Document doc, ElementId id) => ((View)doc.GetElement(id)).Valid;
        public FilteredElementCollector WhereElementIsNotElementType() { _elements = _elements.Where(e => !(e is ElementType)); return this; }
        public FilteredElementCollector WherePasses(ElementFilter filter) { _elements = _elements.Where(filter.PassesFilter); return this; }
        public FilteredElementCollector OfCategory(BuiltInCategory cat) => OfCategoryId(new ElementId((long)cat));
        public FilteredElementCollector OfCategoryId(ElementId id) { _elements = _elements.Where(e => e.Category?.Id == id); return this; }
        public FilteredElementCollector OfClass(Type type) { _elements = _elements.Where(type.IsInstanceOfType); return this; }
        public ICollection<ElementId> ToElementIds() => _elements.Select(e => e.Id).ToArray();
        public IEnumerator<Element> GetEnumerator() => _elements.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
        public void Dispose() { }
    }
    public enum WorksetKind { UserWorkset }
    public enum WorksetVisibility { UseGlobalSetting }
    public class Workset { public WorksetId Id = new WorksetId(); public string Name; public bool IsOpen = true; }
    public class FilteredWorksetCollector
    {
        public FilteredWorksetCollector(Document doc) { }
        public FilteredWorksetCollector OfKind(WorksetKind kind) => this;
        public IList<Workset> ToWorksets() => new List<Workset>();
    }
    public class WorksetDefaultVisibilitySettings
    {
        public static WorksetDefaultVisibilitySettings GetWorksetDefaultVisibilitySettings(Document doc) => new WorksetDefaultVisibilitySettings();
        public bool IsWorksetVisible(WorksetId id) => true;
    }
    public enum SectionType { Header, Body }
    public class TableSectionData { public int NumberOfRows, NumberOfColumns, FirstRowNumber, FirstColumnNumber; }
    public class TableData
    {
        public TableSectionData Body = new TableSectionData(), Header = new TableSectionData();
        public TableSectionData GetSectionData(SectionType section) => section == SectionType.Body ? Body : Header;
    }
    public class ViewSchedule : View
    {
        public ScheduleDefinition Definition = new ScheduleDefinition();
        public TableData Table = new TableData();
        public TableData GetTableData() => Table;
        public string GetCellText(SectionType section, int row, int column)
        {
            var data = Table.GetSectionData(section);
            if (row < data.FirstRowNumber || row >= data.FirstRowNumber + data.NumberOfRows || column < data.FirstColumnNumber || column >= data.FirstColumnNumber + data.NumberOfColumns)
                throw new Exception("cell outside actual bounds");
            return section + ":" + row + ":" + column;
        }
    }
    public class ScheduleFieldId { public int IntegerValue; }
    public class ScheduleField
    {
        public ScheduleFieldId FieldId = new ScheduleFieldId();
        public ElementId ParameterId = new ElementId(-10);
        public string ColumnHeading = "Area", FieldType = "Instance", DisplayType = "Totals";
        public bool IsHidden, IsCalculatedField;
        public string GetName() => ColumnHeading;
    }
    public class ScheduleSortGroupField
    {
        public ScheduleFieldId FieldId = new ScheduleFieldId();
        public string SortOrder = "Ascending";
        public bool ShowHeader, ShowFooter, ShowFooterTitle, ShowFooterCount, ShowBlankLine;
    }
    public class ScheduleFilter
    {
        public ScheduleFieldId FieldId = new ScheduleFieldId();
        public string FilterType = "Equal";
        public object Value;
        public bool IsStringValue => Value is string;
        public bool IsIntegerValue => Value is int;
        public bool IsDoubleValue => Value is double;
        public bool IsElementIdValue => Value is ElementId;
        public string GetStringValue() => (string)Value;
        public int GetIntegerValue() => (int)Value;
        public double GetDoubleValue() => (double)Value;
        public ElementId GetElementIdValue() => (ElementId)Value;
    }
    public class ScheduleDefinition
    {
        public ElementId CategoryId = new ElementId(-10);
        public bool IsItemized, ShowTitle, ShowHeaders, ShowGrandTotal, ShowGrandTotalCount, ShowGrandTotalTitle, IncludeLinkedFiles, HasEmbeddedSchedule;
        public List<ScheduleField> Fields = new List<ScheduleField>();
        public List<ScheduleSortGroupField> Sorting = new List<ScheduleSortGroupField>();
        public List<ScheduleFilter> Filters = new List<ScheduleFilter>();
        public IList<ScheduleFieldId> GetFieldOrder() => Fields.Select(f => f.FieldId).ToArray();
        public ScheduleField GetField(ScheduleFieldId id) => Fields.Single(f => f.FieldId == id);
        public IList<ScheduleSortGroupField> GetSortGroupFields() => Sorting;
        public IList<ScheduleFilter> GetFilters() => Filters;
    }
}

namespace Autodesk.Revit.UI
{
    public enum PostableCommand { RunInterferenceCheck }
    public class RevitCommandId { public static RevitCommandId LookupPostableCommandId(PostableCommand command) => new RevitCommandId(); }
    public class UIApplication
    {
        public bool CanPost = true;
        public int Posted;
        public bool CanPostCommand(RevitCommandId command) => CanPost;
        public void PostCommand(RevitCommandId command) { Posted++; }
    }
}

namespace KPLN_RevitMcpBridge.Services
{
    internal static class RevitService
    {
        public static long IdValue(Autodesk.Revit.DB.ElementId id) => id.Value;
        public static Autodesk.Revit.DB.ElementId Id(long value) => new Autodesk.Revit.DB.ElementId(value);
        public static object ElementInfo(Autodesk.Revit.DB.Element e) => new { id = e.Id.Value.ToString(), unique_id = e.UniqueId, name = e.Name };
        public static Autodesk.Revit.DB.Element[] Resolve(Autodesk.Revit.DB.Document doc, IDictionary<string, object> input)
            => input.List("unique_ids").Cast<string>().Select(doc.GetElement).ToArray();
    }
}
