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
using KPLN_Publication.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;

namespace KPLN_Publication
{
    /// <summary>
    /// Вспомогательный класс, более удобно хранящий информацию о параметрах печати листа
    /// </summary>
    public class YayPrintSettings
    {
        #region Настройки экспорта в PDF
        public string printerName = new System.Drawing.Printing.PrinterSettings().PrinterName;
        public string outputPDFFolder = @"C:\PDF_Print";
        public string pdfNameConstructor = "<Номер листа>_<Имя листа>.pdf";
        public HiddenLineViewsType hiddenLineProcessing = HiddenLineViewsType.VectorProcessing;
        public ColorType colorsType = ColorType.Monochrome;
        public RasterQualityType rasterQuality = RasterQualityType.High;

        public bool isPDFExport = true;
        public bool isMergePdfs = false;
        public bool isPrintToPaper = false;
        //public bool colorStamp;
        public bool isUseOrientation = false;
        public bool isRefreshSchedules = true;
        public bool isExcludeBorders = true;

        public List<PdfColor> excludeColors = new List<PdfColor>();
        public List<PdfColor> excludeBorderColors = new List<PdfColor>();
        #endregion

        #region Настройки экспорта в DWG
        public bool isDWGExport = false;
        public ExportDWGSettingsShell dwgExportSettingShell;
        public string outputDWGFolder = @"C:\DWG_Print";
        public string dwgNameConstructor = "<Номер листа>_<Имя листа>.dwg";
        #endregion

        /// <summary>
        /// Беспараметрический конструктор для сериализатора
        /// </summary>
        public YayPrintSettings()
        {
        }

        /// <summary>
        /// Получение параметров печати
        /// </summary>
        public static YayPrintSettings GetSavedPrintSettings()
        {
            string xmlpath = ActivateFolder();

            YayPrintSettings ps;
            XmlSerializer serializer = new XmlSerializer(typeof(YayPrintSettings));

            if (File.Exists(xmlpath))
            {
                using (StreamReader reader = new StreamReader(xmlpath))
                {
                    ps = (YayPrintSettings)serializer.Deserialize(reader);
                    if (ps == null)
                    {
                        TaskDialog.Show("Внимание", "Не удалось получить сохраненные настройки печати");
                        ps = new YayPrintSettings();
                    }
                }
            }
            else
            {
                ps = new YayPrintSettings
                {
                    excludeColors = new List<PdfColor>
                    {
                        new PdfColor(System.Drawing.Color.FromArgb(0,0,255)),
                        new PdfColor(System.Drawing.Color.FromArgb(192,192,192)),
                        new PdfColor(System.Drawing.Color.FromArgb(242,242,242))
                    }
                };
            }

            ps.excludeBorderColors = new List<PdfColor>
            {
                // Цвет для исключения печати рамки штампа
                new PdfColor(System.Drawing.Color.FromArgb(3,2,51))
            };

            return ps;
        }

        public static bool SaveSettings(YayPrintSettings yps)
        {
            string xmlpath = ActivateFolder();
            if (File.Exists(xmlpath)) 
                File.Delete(xmlpath);
            
            XmlSerializer serializer = new XmlSerializer(typeof(YayPrintSettings));
            using (FileStream writer = new FileStream(xmlpath, FileMode.OpenOrCreate))
            {
                serializer.Serialize(writer, yps);
            }
            
            return true;
        }

        private static string ActivateFolder()
        {
            string appdataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
            
            string bspath = Path.Combine(appdataPath, "bim-starter");
            if (!Directory.Exists(bspath)) 
                Directory.CreateDirectory(bspath);
            
            string localFolder = Path.Combine(bspath, "KPLN_Publication");
            if (!Directory.Exists(localFolder)) 
                Directory.CreateDirectory(localFolder);
            
            string xmlpath = Path.Combine(localFolder, "yayPrintSettings.xml");
            
            return xmlpath;
        }
    }
}
