using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using KPLN_RevitMcpBridge.Server;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace KPLN_RevitMcpBridge.Services
{
    internal sealed class RevitService
    {
        private sealed class DocumentState { public readonly string Id = Guid.NewGuid().ToString("N"); public long Revision; }
        // Revit can return another managed Document wrapper for the same open
        // native document on a later Idling event. ConditionalWeakTable uses
        // reference identity and therefore produced a new document_id for every
        // request. Document overrides Equals/GetHashCode for native identity, so
        // Dictionary preserves one state across wrappers.
        private readonly Dictionary<Document, DocumentState> _documents = new Dictionary<Document, DocumentState>();

        private DocumentState GetDocumentState(Document document)
        {
            DocumentState state;
            if (!_documents.TryGetValue(document, out state))
            {
                state = new DocumentState();
                _documents.Add(document, state);
            }
            return state;
        }

        private void RemoveClosedDocuments()
        {
            foreach (var document in _documents.Keys.Where(d => !d.IsValidObject).ToArray())
                _documents.Remove(document);
        }

        public void DocumentChanged(object sender, DocumentChangedEventArgs args) { GetDocumentState(args.GetDocument()).Revision++; }

        internal static long IdValue(ElementId id)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            return id.IntegerValue;
#else
            return id.Value;
#endif
        }
        internal static ElementId Id(long value)
        {
#if Debug2020 || Revit2020 || Debug2023 || Revit2023
            if (value < int.MinValue || value > int.MaxValue) throw new BridgeException("invalid_id", "ElementId вне диапазона Revit этой версии.");
            return new ElementId((int)value);
#else
            return new ElementId(value);
#endif
        }
        private object DocumentInfo(Document doc)
        {
            var state = GetDocumentState(doc);
            return new
            {
                document_id = state.Id, revision = state.Revision,
                title = doc.Title, path = doc.PathName, is_family = doc.IsFamilyDocument, is_linked = doc.IsLinked,
                is_read_only = doc.IsReadOnly, is_workshared = doc.IsWorkshared, is_modified = doc.IsModified
            };
        }
        internal static object ElementInfo(Element e) => new
        {
            id = IdValue(e.Id).ToString(CultureInfo.InvariantCulture), unique_id = e.UniqueId, name = e.Name,
            category = e.Category?.Name, category_id = e.Category == null ? null : IdValue(e.Category.Id).ToString(CultureInfo.InvariantCulture),
            class_name = e.GetType().Name, is_type = e is ElementType,
            type_id = IdValue(e.GetTypeId()).ToString(CultureInfo.InvariantCulture)
        };
        private static object Box(Element element)
        {
            var box = element.get_BoundingBox(null);
            if (box == null) return null;
            var points = (from x in new[] { box.Min.X, box.Max.X } from y in new[] { box.Min.Y, box.Max.Y } from z in new[] { box.Min.Z, box.Max.Z }
                          select box.Transform.OfPoint(new XYZ(x, y, z))).ToArray();
            return new { coordinate_system = "document_internal", units = "mm", min = new[] { points.Min(p => p.X) * 304.8, points.Min(p => p.Y) * 304.8, points.Min(p => p.Z) * 304.8 },
                max = new[] { points.Max(p => p.X) * 304.8, points.Max(p => p.Y) * 304.8, points.Max(p => p.Z) * 304.8 } };
        }

        public object Execute(UIApplication app, Dictionary<string, object> input)
        {
            string command = input.Text("command", true);
            if (command == "get_documents")
            {
                RemoveClosedDocuments();
                return new { ok = true, result = new { documents = app.Application.Documents.Cast<Document>().Select(DocumentInfo).ToArray(), active_document_id = app.ActiveUIDocument == null ? null : GetDocumentState(app.ActiveUIDocument.Document).Id } };
            }
            var uidoc = app.ActiveUIDocument;
            if (uidoc == null) throw new BridgeException("no_document", "В Revit нет активного документа.", 409);
            var doc = uidoc.Document;
            var state = GetDocumentState(doc);
            if (command != "get_context" && input.Text("document_id", true) != state.Id)
                throw new BridgeException("wrong_document", "Активный документ изменился; получите get_context заново.", 409);
            object result;
            switch (command)
            {
                case "get_context":
                    result = new { document = DocumentInfo(doc), active_view = ElementInfo(uidoc.ActiveView), selection_count = uidoc.Selection.GetElementIds().Count, revit_version = app.Application.VersionNumber,
                        units = new { geometry = "mm", coordinates = "document_internal", double_parameters = "Revit internal units; see data_type/unit_type" } }; break;
                case "get_selection":
                    result = Page(uidoc.Selection.GetElementIds().OrderBy(IdValue).Select(doc.GetElement).Where(e => e != null), input); break;
                case "get_categories":
                    result = doc.Settings.Categories.Cast<Category>().OrderBy(c => c.Name).Select(c => new { id = IdValue(c.Id).ToString(CultureInfo.InvariantCulture), name = c.Name, category_type = c.CategoryType.ToString() }).ToArray(); break;
                case "find_elements":
                    using (var collector = new FilteredElementCollector(doc))
                    {
                        if (input.Flag("include_types")) collector.WhereElementIsElementType(); else collector.WhereElementIsNotElementType();
                        if (input.Get("category_id") != null) collector.OfCategoryId(Id(Json.Integer(input.Get("category_id"), "category_id")));
                        var name = input.Text("name_contains");
                        result = Page(collector.Where(e => string.IsNullOrEmpty(name) || (e.Name ?? "").IndexOf(name, StringComparison.OrdinalIgnoreCase) >= 0).OrderBy(e => IdValue(e.Id)), input);
                    }
                    break;
                case "get_elements":
                    result = Resolve(doc, input).Select(e => new { element = ElementInfo(e), bounding_box = Box(e), parameters = ParameterService.Read(e), location = ElementLocationService.Read(e) }).ToArray(); break;
                case "get_group_members": result = ModelInspectionService.GroupMembers(doc, input); break;
                case "get_sheet_contents": result = ModelInspectionService.SheetContents(doc, input); break;
                case "get_view_elements": result = ModelInspectionService.ViewElements(doc, input); break;
                case "get_view_visibility": result = ModelInspectionService.ViewVisibility(doc, input); break;
                case "get_schedule_data": result = ModelInspectionService.ScheduleData(doc, input); break;
                case "check_intersections": result = InterferenceService.Check(doc, input); break;
                case "open_interference_check": result = InterferenceService.OpenNative(app); break;
                case "export_sheets_pdf":
                    if (Json.Integer(input.Get("expected_revision"), "expected_revision") != state.Revision)
                        throw new BridgeException("stale_document", "Модель изменилась; получите контекст заново перед экспортом.", 409);

                    result = SheetPdfExportService.Export(doc, input);
                    break;
                case "print_sheets_pdf":
                    if (Json.Integer(input.Get("expected_revision"), "expected_revision") != state.Revision)
                        throw new BridgeException("stale_document", "Модель изменилась; получите контекст заново перед печатью.", 409);

                    result = PublicationPrintService.Print(doc, input);
                    break;
                case "get_views":
                    using (var collector = new FilteredElementCollector(doc))
                        result = Page(collector.OfClass(typeof(View)).Cast<View>().Where(v => !v.IsTemplate).OrderBy(v => IdValue(v.Id)), input);
                    break;
                case "get_families":
                    using (var collector = new FilteredElementCollector(doc))
                        result = Page(collector.OfClass(typeof(Family)).OrderBy(e => IdValue(e.Id)), input);
                    break;
                case "get_family_types":
                    var family = doc.GetElement(input.Text("family_unique_id", true)) as Family;
                    if (family == null) throw new BridgeException("not_found", "Семейство не найдено.", 404);
                    result = Page(family.GetFamilySymbolIds().OrderBy(IdValue).Select(doc.GetElement), input); break;
                case "get_family_document": result = FamilyDocument(doc, input); break;
                case "get_links":
                    using (var collector = new FilteredElementCollector(doc))
                        result = collector.OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().Select(l => new { element = ElementInfo(l), loaded = l.GetLinkDocument() != null,
                            linked_document = l.GetLinkDocument() == null ? null : DocumentInfo(l.GetLinkDocument()), transform = TransformInfo(l.GetTotalTransform()) }).ToArray();
                    break;
                case "select_elements":
                    var elements = Resolve(doc, input);
                    uidoc.Selection.SetElementIds(elements.Select(e => e.Id).ToList());
                    if (input.Flag("show")) uidoc.ShowElements(elements.Select(e => e.Id).ToList());
                    result = new { selected = elements.Select(ElementInfo).ToArray() }; break;
                case "set_parameters":
                    if (Json.Integer(input.Get("expected_revision"), "expected_revision") != state.Revision)
                        throw new BridgeException("stale_document", "Модель изменилась. Перечитайте параметры и повторите планирование.", 409);
                    result = ParameterService.Set(doc, input); break;
                case "create_family_types":
                    if (Json.Integer(input.Get("expected_revision"), "expected_revision") != state.Revision)
                        throw new BridgeException("stale_document", "Семейство изменилось. Перечитайте FamilyManager и повторите планирование.", 409);
                    result = FamilyTypeService.Create(doc, input); break;
                case "set_family_type_parameters":
                    if (Json.Integer(input.Get("expected_revision"), "expected_revision") != state.Revision)
                        throw new BridgeException("stale_document", "Семейство изменилось. Перечитайте FamilyManager и повторите планирование.", 409);
                    result = FamilyTypeService.Update(doc, input); break;
                default: throw new BridgeException("unknown_command", "Неизвестная команда: " + command, 404);
            }
            return new { ok = true, document_id = state.Id, revision = state.Revision, result };
        }

        private static object Page(IEnumerable<Element> source, IDictionary<string, object> input)
        {
            int offset = input.Range("offset", 0, 0, 10000000), limit = input.Range("limit", 100, 1, 200);
            var values = source.Skip(offset).Take(limit + 1).ToArray();
            return new { items = values.Take(limit).Select(ElementInfo).ToArray(), offset, next_offset = values.Length > limit ? (int?)(offset + limit) : null };
        }
        internal static Element[] Resolve(Document doc, IDictionary<string, object> input)
        {
            var ids = input.List("unique_ids");
            var seen = new HashSet<string>(); var result = new List<Element>();
            foreach (var value in ids)
            {
                var id = value as string;
                if (string.IsNullOrEmpty(id) || !seen.Add(id)) throw new BridgeException("invalid_id", "Ожидаются неповторяющиеся UniqueId элементов.");
                var e = doc.GetElement(id);
                if (e == null) throw new BridgeException("not_found", "Элемент не найден: " + id, 404);
                result.Add(e);
            }
            return result.ToArray();
        }
        private static double[] Point(XYZ p) => new[] { p.X * 304.8, p.Y * 304.8, p.Z * 304.8 };
        private static double[] Vector(XYZ p) => new[] { p.X, p.Y, p.Z };
        private static object TransformInfo(Transform t) => new { origin_mm = Point(t.Origin), basis_x = Vector(t.BasisX), basis_y = Vector(t.BasisY), basis_z = Vector(t.BasisZ) };
        private static object FamilyDocument(Document doc, IDictionary<string, object> input)
        {
            if (!doc.IsFamilyDocument) throw new BridgeException("not_family_document", "Команда требует открытого редактора семейства.");
            int offset = input.Range("offset", 0, 0, 1000000), limit = input.Range("limit", 50, 1, 100);
            var manager = doc.FamilyManager;
            var parameters = manager.Parameters.Cast<FamilyParameter>().ToArray();
            var types = manager.Types.Cast<FamilyType>().Skip(offset).Take(limit + 1).ToArray();
            return new { family = ElementInfo(doc.OwnerFamily), current_type = manager.CurrentType?.Name,
                parameters = parameters.Select(p => new { id = IdValue(p.Id).ToString(CultureInfo.InvariantCulture), name = p.Definition.Name, storage_type = p.StorageType.ToString(), is_instance = p.IsInstance, is_read_only = p.IsReadOnly, formula = p.Formula }).ToArray(),
                types = types.Take(limit).Select(t => new { name = t.Name, values = parameters.Select(p => new { parameter_id = IdValue(p.Id).ToString(CultureInfo.InvariantCulture), value = FamilyValue(t, p) }).ToArray() }).ToArray(),
                next_offset = types.Length > limit ? (int?)(offset + limit) : null };
        }
        private static object FamilyValue(FamilyType t, FamilyParameter p)
        {
            if (!t.HasValue(p)) return null;
            switch (p.StorageType)
            {
                case StorageType.String: return t.AsString(p);
                case StorageType.Integer: return t.AsInteger(p);
                case StorageType.Double: return t.AsDouble(p);
                case StorageType.ElementId: return IdValue(t.AsElementId(p)).ToString(CultureInfo.InvariantCulture);
                default: return null;
            }
        }
    }
}
