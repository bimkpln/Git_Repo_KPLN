using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using KPLN_Loader.Common;
using KPLN_Clashes_Ribbon.Forms;
using KPLN_Clashes_Ribbon.Tools;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KPLN_Loader.Output.Output;

namespace KPLN_Clashes_Ribbon.Commands
{
    public class CommandPlaceFamily : IExecutableCommand
    {
        private ReportWindow ReportWindow { get; }
        public CommandPlaceFamily(XYZ point, int id1, string info1, int id2, string info2, ReportWindow window)
        {
            ReportWindow = window;
            Id1 = id1;
            ElementInfo1 = info1;
            Id2 = id2;
            ElementInfo2 = info2;
            Point = new XYZ(point.X * 3.28084, point.Y * 3.28084, point.Z * 3.28084);
        }
        private XYZ Point { get; set; }
        private int Id1 { get; set; }
        private string ElementInfo1 { get; set; }
        private int Id2 { get; set; }
        private string ElementInfo2 { get; set; }
        public Result Execute(UIApplication app)
        {
            try
            {
                if (app.ActiveUIDocument != null)
                {
                    Document doc = app.ActiveUIDocument.Document;
                    Element element1 = doc.GetElement(new ElementId(Id1));
                    Element element2 = doc.GetElement(new ElementId(Id2));
                    if (element1 == null && element2 == null)
                    {
                        return Result.Cancelled;
                    }
                    bool el1 = true;
                    bool el2 = true;
                    if (element1 != null)
                    {
                        if (!ElementInfo1.Contains(element1.Name) && !ElementInfo2.Contains(element1.Name))
                        {
                            el1 = false;
                        }
                    }
                    if (element2 != null)
                    {
                        if (!ElementInfo1.Contains(element2.Name) && !ElementInfo2.Contains(element2.Name))
                        {
                            el2 = false;
                        }
                    }
                    if (!el1 && !el2)
                    {
                        return Result.Cancelled;
                    }
                    Transaction t = new Transaction(doc, "Указатель");
                    XYZ transformed_location = doc.ActiveProjectLocation.GetTotalTransform().OfPoint(Point);
                    t.Start();

                    foreach (FamilyInstance instance in new FilteredElementCollector(doc).OfClass(typeof(FamilyInstance)).WhereElementIsNotElementType().ToElements())
                    {
                        try
                        {
                            if (instance.Symbol.FamilyName == "ClashPoint")
                            {
                                doc.Delete(instance.Id);
                            }
                        }
                        catch (Exception) { }
                    }
                    FamilyInstance createdinstance = CreateFamilyInstance(doc, transformed_location);
                    if (createdinstance != null) { ReportWindow.OnClosingActions.Add(new CommandRemoveInstance(doc, createdinstance)); }
                    ZoomTools.ZoomElement(createdinstance.get_BoundingBox(null), app.ActiveUIDocument);
                    t.Commit();
                }
                return Result.Succeeded;
            }
            catch (Exception e)
            {
                PrintError(e);
                return Result.Failed;
            }
        }
        public static List<Level> GetLevels(Document doc)
        {
            List<Level> instances = new List<Level>();
            foreach (Element e in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Levels).WhereElementIsNotElementType().ToElements())
            {
                try
                {
                    instances.Add(e as Level);
                }
                catch (Exception) { }
            }
            return instances;
        }
        private static Level GetNearestLevel(Document doc, double elevation)
        {
            Level l = null;
            double difference = 999999;
            foreach (Element lvl in GetLevels(doc))
            {
                if (l == null)
                {
                    l = lvl as Level;
                    difference = Math.Abs(l.Elevation - elevation);
                    continue;
                }
                if (Math.Abs((lvl as Level).Elevation - elevation) < difference)
                {
                    l = lvl as Level;
                    difference = Math.Abs(l.Elevation - elevation);
                }
            }
            return l;
        }
        
        private static FamilyInstance CreateFamilyInstance(Document doc, XYZ position)
        {
            GetFamilySymbol(doc);
            Level level = GetNearestLevel(doc, position.Z);
            if (level == null)
            {
                throw new Exception("В проекте отсутсвуют уровни!");
            }
            FamilyInstance instance = doc.Create.NewFamilyInstance(position, GetFamilySymbol(doc), level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            doc.Regenerate();
            instance.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).Set(position.Z - level.Elevation);
            doc.Regenerate();
            return instance;
        }
        
        private static FamilySymbol GetFamilySymbol(Document doc)
        {
            string familyName = "ClashPoint";
            foreach (Element element in new FilteredElementCollector(doc).OfClass(typeof(FamilySymbol)).OfCategory(BuiltInCategory.OST_GenericModel))
            {
                FamilySymbol searchSymbol = element as FamilySymbol;
                if (searchSymbol.FamilyName == familyName)
                {
                    searchSymbol.Activate();
                    return searchSymbol;
                }
            }
            try
            {
                if (!doc.LoadFamily(($@"{Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location)}\Source\RevitData\{ModuleData.RevitVersion}\{familyName}.rfa")))
                {
                    throw new Exception("Семейство для метки не найдено!");
                }
                doc.Regenerate();
            }
            catch (Exception e) { PrintError(e); }
            return null;
        }

    }
}
