#region License
/*Данный код опубликован под лицензией Creative Commons Attribution-ShareAlike.
Разрешено использовать, распространять, изменять и брать данный код за основу для производных в коммерческих и
некоммерческих целях, при условии указания авторства и если производные лицензируются на тех же условиях.
Код поставляется "как есть". Автор не несет ответственности за возможные последствия использования.
Зуев Александр, 2020, все права защищены.
This code is listed under the Creative Commons Attribution-ShareAlike license.
You may use, redistribute, remix, tweak, and build upon this work non-commercially and commercially,
as long as you credit the author by linking back and license your new creations under the same terms.
This code is provided 'as is'. Author disclaims any implied warranty.
Zuev Aleksandr, 2020, all rigths reserved.*/
#endregion

using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KPLN_Publication
{
    public static class EntitySupport
    {
        /// <summary>
        /// Получает листы из всех открытых документов
        /// </summary>
        /// <param name="commandData"></param>
        /// <returns></returns>
        public static Dictionary<string, List<MainEntity>> GetAllEntities(ExternalCommandData commandData, View[] selViews)
        {
            Dictionary<string, List<MainEntity>> data = new Dictionary<string, List<MainEntity>>();
            Document mainDoc = commandData.Application.ActiveUIDocument.Document;
            string mainDocTitle = GetDocTitleWithoutRvt(mainDoc.Title);

            List<MainEntity> mainSheets = GetEntitiesFromDocument(mainDoc, selViews);
            data.Add(mainDocTitle, mainSheets);

            List<RevitLinkInstance> links = new FilteredElementCollector(mainDoc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();
            foreach (RevitLinkInstance rli in links)
            {
                Document linkDoc = rli.GetLinkDocument();
                if (linkDoc == null) continue;
                string linkDocTitle = GetDocTitleWithoutRvt(linkDoc.Title);
                if (data.ContainsKey(linkDocTitle)) continue;

                RevitLinkType rlt = mainDoc.GetElement(rli.GetTypeId()) as RevitLinkType;
                List<MainEntity> curSheets = GetEntitiesFromDocument(linkDoc, selViews);

                data.Add(linkDocTitle, curSheets);
            }

            return data;
        }


        public static string GetDocTitleWithoutRvt(string docTitle)
        {
            string result = docTitle;
            if (docTitle.EndsWith(".rvt")) result = docTitle.Substring(0, docTitle.Length - 4);
            return result;
        }


        private static List<MainEntity> GetEntitiesFromDocument(Document doc, View[] selViews)
        {
            List <MainEntity> result = new List <MainEntity>();
            if (selViews.All(sv => sv is ViewSheet))
            {
                result = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(ViewSheet))
                    .Cast<ViewSheet>()
                    .Select(i => new MainEntity(i))
                    .ToList();
            }
            else
            {
                result = new FilteredElementCollector(doc)
                    .WhereElementIsNotElementType()
                    .OfClass(typeof(View))
                    .Cast<View>()
                    .Select(i => new MainEntity(i))
                    .ToList();
            }

            result.Sort();
            return result;
        }

        public static string CheckTitleblocSizeCorrects(ViewSheet sheet, FamilyInstance titleBlock, Logger logger)
        {
            string message = "";

            logger.Write("    ID основной надписи " + titleBlock.Id.IntegerValue.ToString());

            double widthMm = titleBlock.get_Parameter(BuiltInParameter.SHEET_WIDTH).AsDouble() * 304.8;
            widthMm = Math.Round(widthMm);
            double heigthMm = titleBlock.get_Parameter(BuiltInParameter.SHEET_HEIGHT).AsDouble() * 304.8;
            heigthMm = Math.Round(heigthMm);


            double widthMmCheck = -1, heigthMmCheck = -1;

            Parameter checkWidthParam = titleBlock.LookupParameter("Ширина");
            if (checkWidthParam != null)
            {

                widthMmCheck = checkWidthParam.AsDouble() * 304.8;
                widthMmCheck = Math.Round(widthMmCheck);
                logger.Write("    Есть параметр экземпляра Ширина = " + widthMmCheck.ToString("F3"));
            }

            Parameter checkHeigthParam = titleBlock.LookupParameter("Высота");
            if (checkHeigthParam != null)
            {
                heigthMmCheck = checkHeigthParam.AsDouble() * 304.8;
                heigthMmCheck = Math.Round(heigthMmCheck);
                logger.Write("    Есть параметр экземпляра Высота = " + heigthMmCheck.ToString("F3"));
            }

            if (widthMmCheck == -1 || heigthMmCheck == -1)
            {
                logger.Write("    Семейство основной надписи не Weandrevit, проверить размеры не удалось");
                return string.Empty;
            }

            double epsilon = 2.0;
            bool widthEquals = DoubleEquals(widthMm, widthMmCheck, epsilon);
            bool heightEquals = DoubleEquals(heigthMm, heigthMmCheck, epsilon);

            if (!widthEquals || !heightEquals)
            {
                logger.Write("    Проблема с размерами листа! Не совпадает ширина или высота более чем на " + epsilon.ToString("F3") + "мм");
                message += "Лист '" + sheet.SheetNumber + " : " + sheet.Name;
                message += "'. Не удалось определить размеры основной надписи.\n";
                if (widthMm != widthMmCheck)
                {
                    message += "Параметр 'Ширина': " + widthMmCheck.ToString("F0") + "мм, 'Ширина листа': " + widthMm.ToString("F0") + "мм.\n";
                }
                if (heigthMm != heigthMmCheck)
                {
                    message += "Параметр 'Высота': " + heigthMmCheck.ToString("F0") + "мм, 'Высота листа': " + heigthMm.ToString("F0") + "мм.\n";
                }

                message += "Проверьте семейство основной надписи на элементы, выступающие за край листа, или обновите семейство.";
            }
            else
            {
                logger.Write("    Размеры основной надписи корректны");
            }

            return message;
        }

        private static bool DoubleEquals(double d1, double d2, double epsilon)
        {
            double c = Math.Abs(d1 - d2);
            if (c <= epsilon) return true;
            return false;
        }
    }
}
