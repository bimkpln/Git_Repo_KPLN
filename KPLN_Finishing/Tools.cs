using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using KPLN_Library_Forms.UI.HtmlWindow;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Finishing
{
    public static class Tools
    {
        public static List<WPFParameter> GetLocalParameters(Document doc)
        {
            List<WPFParameter> builtInParameters = new List<WPFParameter>();
            Element element = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().FirstElement();
            foreach (Parameter parameter in element.Parameters)
            {
                if (parameter.Definition.Name.ToString().StartsWith("О_"))
                { continue; }
                if (parameter.IsShared)
                {
                    builtInParameters.Add(new WPFLocalParameter(parameter.Definition.Name, parameter.StorageType, parameter.GUID));
                }
            }
            List<WPFParameter> builtInParametersSorted = builtInParameters.OrderBy(o => o.Name).ToList();
            return builtInParametersSorted;
        }
        
        public static Element GetTypeElement(Element element)
        {
            switch (element.Category.Id.IntegerValue)
            {
                case -2000011://Walls
                    try
                    {
                        Wall wall = element as Wall;
                        WallType type = wall.WallType as WallType;
                        if (type == null || type.Id.IntegerValue == -1) { throw new Exception(); }
                        else { return type; }
                    }
                    catch (Exception) { }
                    break;
                case -2000032://Floors
                    try
                    {
                        Floor floor = element as Floor;
                        FloorType type = floor.FloorType as FloorType;
                        if (type == null || type.Id.IntegerValue == -1) { throw new Exception(); }
                        else { return type; }
                    }
                    catch (Exception) { }
                    break;
                default:
                    try
                    {
                        ElementType type = element.Document.GetElement(element.GetTypeId()) as ElementType;
                        if (type == null || type.Id.IntegerValue == -1) { throw new Exception(); }
                        else { return type; }
                    }
                    catch (Exception) { }
                    break;
            }
            return null;
        }
        
        public static string GetGroupParameter(Document doc, Element element)
        {
            try
            {
                if (element.GetType() == typeof(Wall))
                {
                    Wall wall = element as Wall;
                    WallType type = wall.WallType;
                    string value = type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                    if (value == null) { return string.Empty; }
                    else { return value.ToLower(); }
                }
                if (element.GetType() == typeof(Floor))
                {
                    Floor floor = element as Floor;
                    FloorType type = floor.FloorType;
                    string value = type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                    if (value == null) { return string.Empty; }
                    else { return value.ToLower(); }
                }
                if (element.GetType() == typeof(Ceiling))
                {
                    Ceiling ceiling = element as Ceiling;
                    CeilingType type = doc.GetElement(ceiling.GetTypeId()) as CeilingType;
                    string value = type.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                    if (value == null) { return string.Empty; }
                    else { return value.ToLower(); }
                }
            }
            catch (Exception) { }
            return string.Empty;
        }
        
        public static void LoadElementParameters(Document doc)
        {
            List<ScriptParameter> parametersElements = new List<ScriptParameter>() 
            { 
                new ScriptParameter("О_Id помещения", typeof(InstanceBinding), doc),
                new ScriptParameter("О_Имя помещения", typeof(InstanceBinding), doc),
                new ScriptParameter("О_Номер помещения", typeof(InstanceBinding), doc),
                new ScriptParameter("О_Группа", typeof(InstanceBinding), doc),
                new ScriptParameter("О_Описание", typeof(TypeBinding), doc),
                new ScriptParameter("О_Плинтус", typeof(TypeBinding), doc),
                new ScriptParameter("О_Плинтус_Высота", typeof(TypeBinding), doc)
            };
            LoadParameters(doc, new BuiltInCategory[] { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Ceilings }, parametersElements);
        }
        
        public static void LoadRoomParameters(Document doc)
        {
            List<ScriptParameter> parametersRooms = new List<ScriptParameter>()
            { 
                new ScriptParameter("О_ПОМ_ГОСТ_Описание стен", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Описание плинтусов", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Описание полов", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Описание потолков", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Площадь стен_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Площадь потолков_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Площадь полов_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Длина плинтусов_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_Группа", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_Ведомость", typeof(InstanceBinding), doc)
            };
            LoadParameters(doc, new BuiltInCategory[] { BuiltInCategory.OST_Rooms }, parametersRooms);
        }
        
        private static bool IsParameterExist(Element element, ScriptParameter parameter)
        {
            foreach (Parameter p in element.Parameters)
            {
                if (p.IsShared && p.Definition.Name == parameter.Name)
                {
                    return true;
                }
            }
            return false;
        }
        
        public static bool AllParametersExist(Element room, Document doc)
        {
            List<ScriptParameter> parametersRooms = new List<ScriptParameter>()
            { 
                new ScriptParameter("О_ПОМ_ГОСТ_Описание стен", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Описание плинтусов", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Описание полов", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Описание потолков", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Площадь стен_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Площадь потолков_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Площадь полов_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_ГОСТ_Длина плинтусов_Текст", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_Группа", typeof(InstanceBinding), doc),
                new ScriptParameter("О_ПОМ_Ведомость", typeof(InstanceBinding), doc)
            };

            foreach (ScriptParameter parameter in parametersRooms)
            {
                if (!IsParameterExist(room, parameter))
                { return false; }
            }
            return true;
        }
        private static void LoadParameters(Document doc, BuiltInCategory[] requiredCategories, List<ScriptParameter> parameters)
        {
            using (Transaction t = new Transaction(doc, Names.assembly))
            {
                try
                {
                    t.Start();
                    foreach (ScriptParameter parameter in parameters)
                    {
                        ReinsertParameter(parameter, parameter.Categories, doc, requiredCategories);
                    }
                    t.Commit();
                }
                catch (Exception e)
                {
                    HtmlOutput.PrintError(e);
                }
            }
        }
        private static void ReinsertParameter(ScriptParameter parameter, CategorySet categories, Document doc, BuiltInCategory[] requiredCategories)
        {
            CategorySet correctedCategories = new CategorySet();
            if (categories != null)
            {
                foreach (Category cat in categories)
                {
                    correctedCategories.Insert(cat);
                }
            }
            foreach (BuiltInCategory builtInCategory in requiredCategories)
            {
                Category category = Category.GetCategory(doc, builtInCategory);
                if (!correctedCategories.Contains(category))
                {
                    correctedCategories.Insert(category);
                }
            }
            Autodesk.Revit.ApplicationServices.Application app = doc.Application;
            app.SharedParametersFilename = Names.shared_Parameters_File_Path();
            DefinitionFile SharedParametersFile = app.OpenSharedParameterFile();
            //
            foreach (DefinitionGroup dg in SharedParametersFile.Groups)
            {
                if (dg.Name == Names.shared_Parameters_File_Group)
                {
                    Definition externalDefinition = dg.Definitions.get_Item(parameter.Name);
                    if (parameter.TypeBinding == typeof(TypeBinding))
                    {
                        TypeBinding newIB = app.Create.NewTypeBinding(correctedCategories);
                        try
                        {
                            doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA);
                            HtmlOutput.Print($"ДОБАВЛЕН: «{externalDefinition.Name}»", MessageType.Regular);
                        }
                        catch (Exception e)
                        {
                            { HtmlOutput.PrintError(e); }
                            try
                            {
                                doc.ParameterBindings.ReInsert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA);
                                HtmlOutput.Print($"ОБНОВЛЕН: «{externalDefinition.Name}»", MessageType.Regular);
                            }
                            catch (Exception e2)
                            {
                                { HtmlOutput.PrintError(e2); }
                            }

                        }

                    }
                    else
                    {
                        InstanceBinding newIB = app.Create.NewInstanceBinding(correctedCategories);
                        try
                        {
                            doc.ParameterBindings.Insert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA);
                            HtmlOutput.Print(string.Format("ДОБАВЛЕН: «{0}»", externalDefinition.Name), MessageType.Regular);
                        }
                        catch (Exception e)
                        {
                            { HtmlOutput.PrintError(e); }
                            try
                            {
                                doc.ParameterBindings.ReInsert(externalDefinition, newIB, BuiltInParameterGroup.PG_DATA);
                                HtmlOutput.Print(string.Format("ОБНОВЛЕН: «{0}»", externalDefinition.Name), MessageType.Regular);
                            }
                            catch (Exception e2)
                            {
                                HtmlOutput.PrintError(e2);
                            }
                        }
                    }
                }
            }
            DefinitionBindingMapIterator bindingMapIterator = doc.ParameterBindings.ForwardIterator();
            try
            {
                while (bindingMapIterator.MoveNext())
                {
                    if (parameter.Name == bindingMapIterator.Key.Name)
                    {
                        (bindingMapIterator.Key as InternalDefinition).SetAllowVaryBetweenGroups(doc, true);
                    }
                }
            }
            catch (Exception) { }
        }
        public static bool IntersectsBoundingBox(BoundingBoxXYZ boxA, BoundingBoxXYZ boxB)
        {
            if (boxA.Min.X > boxB.Max.X || boxB.Min.X > boxA.Max.X)
            {
                return false;
            }
            if (boxA.Min.Y > boxB.Max.Y || boxB.Min.Y > boxA.Max.Y)
            {
                return false;
            }
            if (boxA.Min.Z > boxB.Max.Z || boxB.Min.Z > boxA.Max.Z)
            {
                return false;
            }
            return true;
        }
        public static bool IntersectsOrNearBoundingBox(BoundingBoxXYZ boxA, BoundingBoxXYZ boxB)
        {
            double max = 10000 / 304.8;
            if ((boxA.Min.X > boxB.Max.X && boxA.Min.X - boxB.Max.X > max) || (boxB.Min.X > boxA.Max.X && boxB.Min.X - boxA.Max.X > max))
            {
                return false;
            }
            if ((boxA.Min.Y > boxB.Max.Y && boxA.Min.Y - boxB.Max.Y > max) || (boxB.Min.Y > boxA.Max.Y && boxB.Min.Y - boxA.Max.Y > max))
            {
                return false;
            }
            if ((boxA.Min.Z > boxB.Max.Z && boxA.Min.Z - boxB.Max.Z > max) || (boxB.Min.Z > boxA.Max.Z && boxB.Min.Z - boxA.Max.Z > max))
            {
                return false;
            }
            return true;
        }
        public static bool IntersectsSolid(Solid solidA, Solid solidB)
        {
            if (IntersectsCurve(solidA, solidB)) { return true; }
            return false;

        }
        public static bool IntersectsCurve(Solid solidA, Solid solidB)
        {
            SolidCurveIntersectionOptions options = new SolidCurveIntersectionOptions();
            options.ResultType = SolidCurveIntersectionMode.CurveSegmentsInside;
            try
            {
                foreach (Edge edge in solidA.Edges)
                {
                    try
                    {
                        SolidCurveIntersection intersection = solidB.IntersectWithCurve(edge.AsCurve(), options);
                        if (intersection.SegmentCount > 0)
                        {
                            return true;
                        }
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }
            return false;
        }
        public static double GetDistance(XYZ pt1, XYZ pt2)
        {
            return Math.Sqrt(Math.Pow(pt2.X - pt1.X, 2) + Math.Pow(pt2.Y - pt1.Y, 2) + Math.Pow(pt2.Z - pt1.Z, 2));
        }
        public static BoundingBoxXYZ Squarify(BoundingBoxXYZ box)
        {
            double x_length = box.Max.X - box.Min.X;
            double y_length = box.Max.Y - box.Min.Y;
            double z_length = box.Max.Z - box.Min.Z;
            double length = Math.Ceiling(Math.Max(x_length, y_length) / 12) * 12;
            length = Math.Max(length, z_length);
            XYZ c = new XYZ(Math.Round((box.Max.X + box.Min.X) / 2 / 12) * 12, Math.Round((box.Max.Y + box.Min.Y) / 2 / 12) * 12, Math.Round((box.Max.Z + box.Min.Z) / 2 / 12) * 12);
            BoundingBoxXYZ bBox = new BoundingBoxXYZ();
            bBox.Min = new XYZ(c.X - length, c.Y - length, c.Z - length);
            bBox.Max = new XYZ(c.X + length, c.Y + length, c.Z + length);
            return bBox;
        }
        public static void ApplyRoom(MatrixElement element, MatrixElement roomElement)
        {
            try
            {
                Room room = roomElement.Element as Room;
                element.Element.LookupParameter(Names.parameter_Room_Id).Set(room.Id.IntegerValue.ToString());
                element.Element.LookupParameter(Names.parameter_Room_Name).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                element.Element.LookupParameter(Names.parameter_Room_Number).Set(room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString());
                element.Element.LookupParameter(Names.parameter_Room_Department).Set(room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString());
            }
            catch (Exception) { }
        }
        public static void ApplyRoom(Element element, Element roomElement)
        {
            try
            {
                Room room = roomElement as Room;
                element.LookupParameter(Names.parameter_Room_Id).Set(room.Id.IntegerValue.ToString());
                element.LookupParameter(Names.parameter_Room_Name).Set(room.get_Parameter(BuiltInParameter.ROOM_NAME).AsString());
                element.LookupParameter(Names.parameter_Room_Number).Set(room.get_Parameter(BuiltInParameter.ROOM_NUMBER).AsString());
                element.LookupParameter(Names.parameter_Room_Department).Set(room.get_Parameter(BuiltInParameter.ROOM_DEPARTMENT).AsString());
            }
            catch (Exception) { }

        }
        public static void CopyFrom(Document doc, MatrixElement element, MatrixElement fromElement, bool applyToFrom = true)
        {
            try
            {
                ElementId roomId = new ElementId(int.Parse(fromElement.Element.LookupParameter(Names.parameter_Room_Id).AsString(), System.Globalization.NumberStyles.Integer));
                Element room = doc.GetElement(roomId);
                ApplyRoom(element.Element, room);
                if (applyToFrom)
                {
                    ApplyRoom(fromElement.Element, room);
                }
            }
            catch (Exception) { }
        }
        public static void CopyFrom(Document doc, Element element, Element fromElement, bool applyToFrom = true)
        {
            try
            {
                ElementId roomId = new ElementId(int.Parse(fromElement.LookupParameter(Names.parameter_Room_Id).AsString(), System.Globalization.NumberStyles.Integer));
                Element room = doc.GetElement(roomId);
                ApplyRoom(element, room);
                if (applyToFrom)
                {
                    ApplyRoom(fromElement, room);
                }
            }
            catch (Exception) { }
        }
        public static void Reset(Element element)
        {
            try
            {
                element.LookupParameter(Names.parameter_Room_Id).Set("");
                element.LookupParameter(Names.parameter_Room_Name).Set("");
                element.LookupParameter(Names.parameter_Room_Number).Set("");
                element.LookupParameter(Names.parameter_Room_Department).Set("");
            }
            catch (Exception) { }
        }
        public static void Print(Exception exception)
        {
            TaskDialog TD = new TaskDialog(Names.assembly);
            TD.MainContent = exception.Message;
            TD.MainInstruction = exception.StackTrace;
            TD.Show();
        }
        public static void Print(Exception exception, string header)
        {
            TaskDialog TD = new TaskDialog(header);
            TD.MainContent = exception.Message;
            TD.MainInstruction = exception.StackTrace;
            TD.Show();
        }
        public static void Update<D>(ExternalCommandData commandData, D definitions)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            List<FRoom> rooms = new List<FRoom>();
            foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements())
            {
                Room room = element as Room;
                if (room.Area > 0 && room.LevelId.IntegerValue != -1)
                {
                    FRoom fRoom = new FRoom(room, definitions as FilterRule);
                    fRoom.GetElements();
                    rooms.Add(fRoom);
                }
            }
            FRoomCollection collection = new FRoomCollection();
            foreach (FRoom room in rooms)
            {
                collection.Add(room);
            }
        }
    }
    public class WPFParameter
    {
        public virtual bool IsLocal { get; }
        public virtual BuiltInParameter BuiltInParameter { get; }
        public virtual StorageType StorageType { get { return StorageType.None; } }
        public virtual string StorageTypeString { get; }
        public virtual string Name { get { return "Нет"; } }
        public virtual string GetStringValue(Element element) { return null; }
    }
    public class WPFBuiltInParameter : WPFParameter
    {
        public override bool IsLocal { get { return false; } }
        public override BuiltInParameter BuiltInParameter { get; }
        public override StorageType StorageType { get; }
        public override string StorageTypeString { get; }
        public override string Name { get; }
        public WPFBuiltInParameter(BuiltInParameter builtInParameter, string name, StorageType storageType)
        {
            BuiltInParameter = builtInParameter;
            Name = string.Format("<{0}>", name);
            StorageType = storageType;
            StorageTypeString = storageType.ToString();
        }
        public override string GetStringValue(Element element)
        {
            string value = "";
            switch (StorageType)
            {
                case StorageType.String:
                    try
                    {
                        value = element.get_Parameter(BuiltInParameter).AsString();
                    }
                    catch (Exception)
                    {
                        try
                        {
                            value = element.get_Parameter(BuiltInParameter).AsValueString();
                        }
                        catch (Exception) { }
                    }
                    break;
                case StorageType.Integer:
                    try
                    {
                        value = element.get_Parameter(BuiltInParameter).AsInteger().ToString();
                    }
                    catch (Exception) { }
                    break;
                case StorageType.Double:
                    try
                    {
                        value = element.get_Parameter(BuiltInParameter).AsDouble().ToString(); ;
                    }
                    catch (Exception) { }
                    break;
                default:
                    break;
            }
            return value;
        }
    }
    public class WPFLocalParameter : WPFParameter
    {
        public override bool IsLocal { get { return true; } }
        public override BuiltInParameter BuiltInParameter { get; }
        public override StorageType StorageType { get; }
        public override string StorageTypeString { get; }
        public override string Name { get; }
        public Guid Guid { get; }
        public WPFLocalParameter(string localParameter, StorageType storageType, Guid guid)
        {
            Guid = guid;
            Name = localParameter;
            StorageType = storageType;
            StorageTypeString = storageType.ToString();
        }
        public override string GetStringValue(Element element)
        {
            string value = "";
            switch (StorageType)
            {
                case StorageType.String:
                    try
                    {
                        value = element.get_Parameter(Guid).AsString();
                    }
                    catch (Exception)
                    {
                        try
                        {
                            value = element.get_Parameter(Guid).AsValueString();
                        }
                        catch (Exception) { }
                    }
                    break;
                case StorageType.Integer:
                    try
                    {
                        value = element.get_Parameter(Guid).AsInteger().ToString();
                    }
                    catch (Exception) { }
                    break;
                case StorageType.Double:
                    try
                    {
                        value = element.get_Parameter(Guid).AsDouble().ToString(); ;
                    }
                    catch (Exception) { }
                    break;
                default:
                    break;
            }
            return value;
        }
    }
    public class ScriptParameter
    {
        public string Name { get; }
        public CategorySet Categories { get; }
        public Type TypeBinding { get; }
        public ScriptParameter(string name, Type typebinding, Document doc)
        {
            Name = name;
            TypeBinding = typebinding;
            //
            BindingMap bindingMap = doc.ParameterBindings;
            DefinitionBindingMapIterator iterator = bindingMap.ForwardIterator();
            iterator.Reset();
            Categories = new CategorySet();
            while (iterator.MoveNext())
            {
                try
                {
                    InternalDefinition definition = iterator.Key as InternalDefinition;
                    string definitionName = iterator.Key.Name;
                    if (definitionName == name)
                    {
                        ElementBinding binding = iterator.Current as ElementBinding;
                        Categories = binding.Categories;
                        break;
                    }
                }
                catch { }
            }
        }
    }
    public class FilterRule
    {
        public virtual string GetValue(Element element)
        {
            return null;
        }
    }
    public class FilterLookupRule : FilterRule
    {
        private string Parameter;
        public FilterLookupRule(string parameter)
        {
            this.Parameter = parameter;
        }
        public override string GetValue(Element element)
        {
            try
            {
                return element.LookupParameter(this.Parameter).AsValueString();
            }
            catch (Exception)
            {
                return element.LookupParameter(this.Parameter).AsString();
            }
        }
    }
    public class FilterBuiltInRule : FilterRule
    {
        private BuiltInParameter Parameter;
        public FilterBuiltInRule(BuiltInParameter parameter)
        {
            this.Parameter = parameter;
        }
        public override string GetValue(Element element)
        {
            try
            {
                return element.get_Parameter(Parameter).AsValueString();
            }
            catch (Exception)
            {
                return element.get_Parameter(Parameter).AsString();
            }
        }
    }
    public class FRoomCollection
    {
        public List<string> Keys = new List<string>();
        public List<List<FRoom>> Rooms = new List<List<FRoom>>();
        public void Add(FRoom room)
        {
            for (int i = 0; i < this.Keys.Count; i++)
            {
                if (this.Keys[i] == room.Key)
                {
                    Rooms[i].Add(room);
                    return;
                }
            }
            this.Keys.Add(room.Key);
            this.Rooms.Add(new List<FRoom>() { room });
        }
    }
    public class FRoom
    {
        public Room Room { get; }
        public List<Element> Elements = new List<Element>();
        public List<Element> Types = new List<Element>();
        private List<FilterRule> Rules = new List<FilterRule>();
        public string Key
        {
            get
            {
                List<string> keys = new List<string>();
                foreach (Element type in this.Types)
                {
                    keys.Add(type.Id.IntegerValue.ToString());
                }
                keys.Sort();
                if (this.Rules.Count != 0)
                {
                    foreach (FilterRule rule in this.Rules)
                    {
                        keys.Add(rule.GetValue(this.Room));
                    }
                }
                return string.Join(":", keys);
            }
        }
        public void GetElements()
        {
            foreach (BuiltInCategory cat in new BuiltInCategory[] { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Floors, BuiltInCategory.OST_Ceilings })
            {
                foreach (Element element in new FilteredElementCollector(this.Room.Document).OfCategory(cat).WhereElementIsNotElementType().ToElements())
                {
                    if (int.Parse(element.LookupParameter("О_Id помещения").AsString(), System.Globalization.NumberStyles.Integer) == this.Room.Id.IntegerValue)
                    {
                        Element type = Tools.GetTypeElement(element);
                        if (!ElementInList(element, this.Elements))
                        {
                            this.Elements.Add(type);
                        }
                        if (!ElementInList(type, this.Types))
                        {
                            this.Types.Add(type);
                        }
                    }
                }
            }
        }
        public FRoom(Room room)
        {
            this.Room = room;
        }
        public FRoom(Room room, FilterRule rule)
        {
            this.Room = room;
            this.Rules.Add(rule);
        }
        public FRoom(Room room, List<FilterRule> rules)
        {
            this.Room = room;
            foreach (FilterRule rule in rules)
            {
                this.Rules.Add(rule);
            }
        }
        public void Add(Element element)
        {
            Element type = Tools.GetTypeElement(element);
            if (!ElementInList(element, this.Elements))
            {
                this.Elements.Add(type);
            }
            if (!ElementInList(type, this.Types))
            {
                this.Types.Add(type);
            }
        }
        public void Add(List<Element> elements)
        {
            foreach (Element element in elements)
            {
                Element type = Tools.GetTypeElement(element);
                if (!ElementInList(element, this.Elements))
                {
                    this.Elements.Add(type);
                }
                if (!ElementInList(type, this.Types))
                {
                    this.Types.Add(type);
                }
            }
        }
        public FRoom(Room room, List<Element> elements)
        {
            this.Room = room;
            foreach (Element element in elements)
            {
                Element type = Tools.GetTypeElement(element);
                if (!ElementInList(element, this.Elements))
                {
                    this.Elements.Add(type);
                }
                if (!ElementInList(type, this.Types))
                {
                    this.Types.Add(type);
                }
            }
        }
        private bool ElementInList(Element element, List<Element> elements)
        {
            foreach (Element e in elements)
            {
                if (e.Id.IntegerValue == element.Id.IntegerValue)
                {
                    return true;
                }
            }
            return false;
        }
    }
}
