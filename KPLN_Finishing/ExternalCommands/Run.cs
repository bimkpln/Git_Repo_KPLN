using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using static KPLN_Finishing.Tools;

namespace KPLN_Finishing.ExternalCommands
{
    [TransactionAttribute(TransactionMode.Manual)]
    [RegenerationAttribute(RegenerationOption.Manual)]
    class Run : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            //
            TaskDialog td = new TaskDialog("Рассчет элементов");
            td.TitleAutoPrefix = false;
            td.MainContent = "Запустить расчет?";
            td.FooterText = Names.task_dialog_hint;
            td.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Close;
            TaskDialogResult result = td.Show();
            if (result != TaskDialogResult.Yes)
            { return Result.Cancelled; }
            //
            int n = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToElements().Count;
            int n_l = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElementIds().Count;
            string s = "{0} из " + n.ToString() + " помещений обработано";
            using (ProgressForm pf = new ProgressForm("Рассчет отделки", s, n, n_l))
            {
                foreach (ElementId PickedLevel in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElementIds())
                {
                    ElementLevelFilter levelFilter = new ElementLevelFilter(PickedLevel);
                    n = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().WherePasses(levelFilter).ToElements().Count;
                    pf._format = "{0} из " + n.ToString() + " помещений обработано";
                    pf.ResetMax(n);
                    pf.SetInfoLevel(doc.GetElement(PickedLevel).get_Parameter(BuiltInParameter.DATUM_TEXT).AsString());
                    List<MatrixElement> unidentifiedElements = new List<MatrixElement>();
                    Matrix matrix = new Matrix();
                    pf.SetInfoStrip(Names.message_Matrix_Reset);
                    matrix.Clear();
                    List<MatrixElement> matrixRooms = new List<MatrixElement>();
                    foreach (Element element in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().WherePasses(levelFilter).ToElements())
                    {
                        pf.Increment();
                        try
                        {
                            if (element.LevelId.IntegerValue == PickedLevel.IntegerValue && element.LevelId.IntegerValue != -1)
                            {
                                Room room = element as Room;
                                if (room.Area > 0.0000001)
                                {
                                    Solid solid = null;
                                    foreach (GeometryObject obj in room.ClosedShell)
                                    {
                                        if (obj.GetType() == typeof(Solid))
                                        {
                                            solid = obj as Solid;
                                            break;
                                        }
                                    }
                                    if (solid == null)
                                    {
                                        pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Element_Geometry_Filter));
                                        continue;
                                    }
                                    XYZ centroid = solid.ComputeCentroid();
                                    BoundingBoxXYZ boundingBox = solid.GetBoundingBox();
                                    BoundingBoxXYZ normalizedBoundingBox = new BoundingBoxXYZ();
                                    normalizedBoundingBox.Min = new XYZ(boundingBox.Min.X + centroid.X, boundingBox.Min.Y + centroid.Y, boundingBox.Min.Z + centroid.Z);
                                    normalizedBoundingBox.Max = new XYZ(boundingBox.Max.X + centroid.X, boundingBox.Max.Y + centroid.Y, boundingBox.Max.Z + centroid.Z);
                                    MatrixElement newMatrixElement = new MatrixElement(element, solid, centroid, normalizedBoundingBox);
                                    matrix.AppendToRooms(newMatrixElement);
                                    matrixRooms.Add(newMatrixElement);
                                    pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Matrix_Adding_Element));
                                }
                                else
                                {
                                    pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Element_Filter));
                                }
                            }
                            else
                            {
                                pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Element_Filter));
                            }
                        }
                        catch (Exception) { }
                    }
                    if (matrixRooms.Count == 0)
                    {
                        continue;
                    }
                    pf.SetInfoStrip(Names.message_Matrix_Squarifying);
                    matrix.BoundingBox = Squarify(matrix.BoundingBox);
                    pf.SetInfoStrip(Names.message_Matrix_Preparing_Conainers);
                    MatrixContainer matrixContainer = matrix.PrepareContainers();
                    foreach (MatrixElement matrixRoom in matrixRooms)
                    {
                        pf.SetInfoStrip(string.Format("{0} - {1}", matrixRoom.Element.Id.ToString(), Names.message_Matrix_Adding_Element));
                        matrixContainer.InsertItem(matrixRoom);
                    }
                    List<MatrixElement> matrixElements = new List<MatrixElement>();
                    BuiltInCategory[] categories = new BuiltInCategory[] { BuiltInCategory.OST_Walls, BuiltInCategory.OST_Ceilings, BuiltInCategory.OST_Floors };
                    using (Transaction t = new Transaction(doc, Names.assembly))
                    {
                        t.Start();
                        foreach (BuiltInCategory category in categories)
                        {
                            n = new FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().WherePasses(levelFilter).ToElements().Count;
                            pf._format = "{0} из " + n.ToString() + " элементов обработано <" + category.ToString() + ">";
                            pf.ResetMax(n);
                            foreach (Element element in new FilteredElementCollector(doc).OfCategory(category).WhereElementIsNotElementType().WherePasses(levelFilter).ToElements())
                            {
                                pf.Increment();
                                
                                string elemGroupModelParamData = GetTypeElement(element).get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString();
                                if (string.IsNullOrEmpty(elemGroupModelParamData)) continue;
                                
                                if (elemGroupModelParamData.ToLower() == Names.value_All_Model_Model && element.LevelId.IntegerValue == PickedLevel.IntegerValue)
                                {
                                    Options options = new Options();
                                    options.IncludeNonVisibleObjects = false;
                                    GeometryElement geomEl = element.get_Geometry(options);
                                    Solid solid = null;
                                    foreach (GeometryObject geomObj in geomEl)
                                    {
                                        try
                                        {
                                            if (geomObj.GetType() == typeof(Solid))
                                            {
                                                Solid enumerateSolid = geomObj as Solid;
                                                if (solid == null)
                                                {
                                                    solid = enumerateSolid;
                                                }
                                                else
                                                {
                                                    if (enumerateSolid.Volume > solid.Volume || solid == null)
                                                    {
                                                        solid = enumerateSolid;
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception) { }
                                    }
                                    if (solid != null)
                                    {
                                        Reset(element);
                                        XYZ centroid = solid.ComputeCentroid();
                                        BoundingBoxXYZ boundingBox = solid.GetBoundingBox();
                                        BoundingBoxXYZ normalizedBoundingBox = new BoundingBoxXYZ();
                                        normalizedBoundingBox.Min = new XYZ(boundingBox.Min.X + centroid.X, boundingBox.Min.Y + centroid.Y, boundingBox.Min.Z + centroid.Z);
                                        normalizedBoundingBox.Max = new XYZ(boundingBox.Max.X + centroid.X, boundingBox.Max.Y + centroid.Y, boundingBox.Max.Z + centroid.Z);
                                        MatrixElement newMatrixElement = new MatrixElement(element, solid, centroid, normalizedBoundingBox);
                                        matrixElements.Add(newMatrixElement);
                                        pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Matrix_Adding_Element));
                                    }
                                    else
                                    {
                                        pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Element_Geometry_Filter));
                                    }
                                }
                                else
                                {
                                    pf.SetInfoStrip(string.Format("{0} - {1}", element.Id.ToString(), Names.message_Element_Filter));
                                }
                                
                            }
                        }
                        t.Commit();
                    }
                    if (matrixElements.Count == 0)
                    {
                        continue;
                    }
                    pf.SetInfoStrip(Names.message_Matrix_Optimising);
                    matrixContainer.Optimize();
                    n = matrixElements.Count;
                    pf._format = "{0} из " + n.ToString() + " элементов определено";
                    pf.ResetMax(n);
                    using (Transaction t = new Transaction(doc, Names.assembly))
                    {
                        t.Start();
                        foreach (MatrixElement element in matrixElements)
                        {
                            try
                            {
                                List<MatrixElement> context = matrixContainer.GetContext(element);
                                if (context.Count == 0)
                                {
                                    unidentifiedElements.Add(element);
                                }
                                else
                                {
                                    List<MatrixElement> Intersection = new List<MatrixElement>();
                                    foreach (MatrixElement contextRoom in context)
                                    {
                                        if (IntersectsSolid(element.Solid, contextRoom.Solid))
                                        {
                                            Intersection.Add(contextRoom);
                                        }
                                    }
                                    if (Intersection.Count == 1)
                                    {
                                        pf.SetInfoStrip(string.Format("<{0}> - {1}", element.Element.Id.ToString(), Names.message_Element_Calculated_Single));
                                        pf.Increment();
                                        ApplyRoom(element, Intersection[0]);
                                    }
                                    else
                                    {
                                        if (Intersection.Count > 1)
                                        {
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", element.Element.Id.ToString(), Names.message_Element_Calculated_Nearest));
                                            pf.Increment();
                                            ApplyRoom(element, GetClosestGeometry(element, Intersection));
                                        }
                                        else
                                        {
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", element.Element.Id.ToString(), Names.message_Element_NotCalculated));
                                            unidentifiedElements.Add(element);
                                        }
                                    }
                                }
                            }
                            catch (Exception e) { Print(e); }
                        }
                        t.Commit();
                    }
                    if (unidentifiedElements.Count == 0)
                    {
                        continue;
                    }
                    int loop = 0;
                    while (loop < 5)
                    {
                        loop++;
                        using (Transaction t = new Transaction(doc, Names.assembly))
                        {
                            List<MatrixElement> unidentifiedElementsUpdated = new List<MatrixElement>();
                            t.Start();
                            foreach (MatrixElement unidentifiedElement in unidentifiedElements)
                            {
                                try
                                {
                                    if (unidentifiedElement.Element.Category.Id.IntegerValue == -2000011)
                                    {
                                        List<string> values = new List<string>();
                                        Wall wall = unidentifiedElement.Element as Wall;
                                        List<Element> joinElements = GetElementsAtJoin(wall);
                                        if (joinElements.Count != 0)
                                        {
                                            foreach (Element joinElement in joinElements)
                                            {
                                                values.Add(joinElement.LookupParameter(Names.parameter_Room_Id).AsString());
                                            }
                                            wall.LookupParameter(Names.parameter_Room_Id).Set(values.Max());
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_Calculated_Chain));
                                            pf.Increment();
                                        }
                                        else
                                        {
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_NotCalculated));
                                            unidentifiedElementsUpdated.Add(unidentifiedElement);
                                        }
                                    }
                                    else
                                    { unidentifiedElementsUpdated.Add(unidentifiedElement); }
                                }
                                catch (Exception) { }
                            }
                            unidentifiedElements.Clear();
                            foreach (MatrixElement newUnidentifiedElement in unidentifiedElementsUpdated)
                            {
                                unidentifiedElements.Add(newUnidentifiedElement);
                            }
                            t.Commit();
                        }
                    }
                    pf.SetInfoStrip(Names.message_Matrix_Reset);
                    matrix.Clear();
                    foreach (MatrixElement matrixElement in matrixElements)
                    {
                        pf.SetInfoStrip(string.Format("<{0}> - {1}", matrixElement.Element.Id.ToString(), Names.message_Matrix_Adding_Element));
                        matrix.AppendToElements(matrixElement);
                    }
                    pf.SetInfoStrip(Names.message_Matrix_Squarifying);
                    matrix.BoundingBox = Squarify(matrix.BoundingBox);
                    pf.SetInfoStrip(Names.message_Matrix_Preparing_Conainers);
                    matrixContainer = matrix.PrepareContainers();
                    foreach (MatrixElement matrixElement in matrixElements)
                    {
                        pf.SetInfoStrip(string.Format("<{0}> - {1}", matrixElement.Element.Id.ToString(), Names.message_Matrix_Adding_Element));
                        matrixContainer.InsertItem(matrixElement);
                    }
                    pf.SetInfoStrip(Names.message_Matrix_Optimising);
                    matrixContainer.Optimize();
                    loop = 0;
                    while (loop < 5)
                    {
                        loop++;
                        using (Transaction t = new Transaction(doc, Names.assembly))
                        {
                            t.Start();
                            foreach (MatrixElement unidentifiedElement in unidentifiedElements)
                            {
                                try
                                {
                                    if (unidentifiedElement.Element.LookupParameter(Names.parameter_Room_Id).AsString() != "")
                                    {
                                        continue;
                                    }
                                    List<MatrixElement> context = matrixContainer.GetContext(unidentifiedElement);
                                    if (context.Count == 0)
                                    {
                                        continue;
                                    }
                                    else
                                    {
                                        List<MatrixElement> Intersection = new List<MatrixElement>();
                                        foreach (MatrixElement contextElement in context)
                                        {
                                            if (unidentifiedElement.Element.Id.IntegerValue == contextElement.Element.Id.IntegerValue || contextElement.Element.LookupParameter(Names.parameter_Room_Id).AsString() == "") { continue; }
                                            if (IntersectsSolid(unidentifiedElement.Solid, contextElement.Solid)) { Intersection.Add(contextElement); }
                                        }
                                        if (Intersection.Count == 1)
                                        {
                                            pf.Increment();
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_Calculated_Single));
                                            CopyFrom(doc, unidentifiedElement, Intersection[0], false);
                                            continue;
                                        }
                                        if (Intersection.Count > 1)
                                        {
                                            pf.Increment();
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_Calculated_Nearest));
                                            CopyFrom(doc, unidentifiedElement, GetClosestGeometry(unidentifiedElement, Intersection), false);
                                            continue;
                                        }
                                        else
                                        {
                                            pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_NotCalculated));
                                            continue;
                                        }
                                    }
                                }
                                catch (Exception) { }
                            }
                            t.Commit();
                        }
                    }
                    loop = 0;
                    while (loop < 5)
                    {
                        loop++;
                        using (Transaction t = new Transaction(doc, Names.assembly))
                        {
                            t.Start();
                            foreach (MatrixElement unidentifiedElement in unidentifiedElements)
                            {
                                try
                                {
                                    if (unidentifiedElement.Element.LookupParameter(Names.parameter_Room_Id).AsString() == "")
                                    {
                                        List<MatrixElement> context = matrixContainer.GetContext(unidentifiedElement);
                                        if (context.Count == 0)
                                        {
                                            continue;
                                        }
                                        else
                                        {
                                            List<MatrixElement> bBox = new List<MatrixElement>();
                                            foreach (MatrixElement contextElement in context)
                                            {
                                                if (unidentifiedElement.Element.Id.IntegerValue == contextElement.Element.Id.IntegerValue || contextElement.Element.LookupParameter(Names.parameter_Room_Id).AsString() == "")
                                                { continue; }
                                                if (IntersectsBoundingBox(unidentifiedElement.BoundingBox, contextElement.BoundingBox))
                                                { bBox.Add(contextElement); }
                                            }
                                            if (bBox.Count == 1)
                                            {
                                                pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_Calculated_Single));
                                                pf.Increment();
                                                CopyFrom(doc, unidentifiedElement, bBox[0], false);
                                                continue;
                                            }
                                            if (bBox.Count > 1)
                                            {
                                                if (IsUniform(bBox))
                                                {
                                                    pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_Calculated_Nearest));
                                                    pf.Increment();
                                                    CopyFrom(doc, unidentifiedElement, bBox[0], false);
                                                }
                                                continue;
                                            }
                                            else
                                            {
                                                pf.SetInfoStrip(string.Format("<{0}> - {1}", unidentifiedElement.Element.Id.ToString(), Names.message_Element_NotCalculated));
                                                continue;
                                            }
                                        }
                                    }
                                }
                                catch (Exception) { }
                            }
                            t.Commit();
                        }
                    }
                }
            }
            return Result.Succeeded;
        }
        private static bool IsUniform(List<MatrixElement> list)
        {
            HashSet<string> values = new HashSet<string>();
            foreach (MatrixElement Element in list)
            {
                values.Add(Element.Element.LookupParameter(Names.parameter_Room_Id).AsString());
            }
            if (values.Count == 1) { return true; }
            return false;
        }
        private static List<Element> GetElementsAtJoin(Wall wall)
        {
            List<Element> elements = new List<Element>();
            LocationCurve location = wall.Location as LocationCurve;
            foreach (int i in new int[] { 0, 1 })
            {
                ElementArray elementArray = location.get_ElementsAtJoin(i);
                if (!elementArray.IsEmpty)
                {
                    foreach (Element element in elementArray)
                    {
                        if (element.Category.Id.IntegerValue == -2000011 && element.LookupParameter(Names.parameter_Room_Id).AsString() != "" && element.LookupParameter(Names.parameter_Room_Id).AsString() != null)
                        {
                            elements.Add(element);
                        }
                    }
                }
            }
            return elements;
        }
        private static MatrixElement GetClosestGeometry(MatrixElement element, List<MatrixElement> intersection)
        {
            MatrixElement closestElement = null;
            double minDistance = 999999;
            foreach (MatrixElement contextRoom in intersection)
            {
                try
                {
                    double distance = GetDistance(element, contextRoom);
                    if (closestElement == null)
                    {
                        closestElement = contextRoom;
                        minDistance = distance;
                    }
                    else
                    {
                        if (minDistance > distance)
                        {
                            closestElement = contextRoom;
                            minDistance = distance;
                        }
                    }
                }
                catch (Exception e)
                {
                    Print(e);
                    return intersection[0];
                }
            }
            return closestElement;
        }
        private static double GetDistance(MatrixElement element, MatrixElement room)
        {
            double minimum = 999999;
            try
            {
                foreach (Edge edgeA in element.Solid.Edges)
                {
                    foreach (Edge edgeB in room.Solid.Edges)
                    {
                        try
                        {
                            Curve curveA = edgeA.AsCurve();
                            Curve curveB = edgeB.AsCurve();
                            foreach (int a in new int[] { 0, 1 })
                            {
                                foreach (int b in new int[] { 0, 1 })
                                {
                                    double v = Tools.GetDistance(curveA.GetEndPoint(a), curveB.GetEndPoint(b));
                                    if (minimum > v || minimum == 999999)
                                    { minimum = v; }
                                }
                            }
                        }
                        catch (Exception) { }
                    }
                }
                foreach (Edge edgeA in element.Solid.Edges)
                {
                    foreach (Face faceB in room.Solid.Faces)
                    {
                        foreach (int b in new int[] { 0, 1 })
                        {
                            try
                            {
                                XYZ pp = faceB.Project(edgeA.AsCurve().GetEndPoint(b)).XYZPoint;
                                UV up = faceB.Project(edgeA.AsCurve().GetEndPoint(b)).UVPoint;
                                if (faceB.IsInside(up))
                                {
                                    foreach (int a in new int[] { 0, 1 })
                                    {
                                        double v = Tools.GetDistance(edgeA.AsCurve().GetEndPoint(a), pp);
                                        if (minimum > v || minimum == 999999)
                                        { minimum = v; }
                                    }
                                }
                            }
                            catch (Exception) { }
                        }
                    }
                }
                foreach (Edge edgeA in room.Solid.Edges)
                {
                    foreach (Face faceB in element.Solid.Faces)
                    {
                        foreach (int b in new int[] { 0, 1 })
                        {
                            try
                            {
                                XYZ pp = faceB.Project(edgeA.AsCurve().GetEndPoint(b)).XYZPoint;
                                UV up = faceB.Project(edgeA.AsCurve().GetEndPoint(b)).UVPoint;
                                if (faceB.IsInside(up))
                                {
                                    foreach (int a in new int[] { 0, 1 })
                                    {
                                        double v = Tools.GetDistance(edgeA.AsCurve().GetEndPoint(a), pp);
                                        if (minimum > v || minimum == 999999)
                                        { minimum = v; }
                                    }
                                }
                            }
                            catch (Exception) { }
                        }
                    }
                }
            }
            catch (Exception e)
            { Print(e); }
            return minimum;
        }
    }
}
