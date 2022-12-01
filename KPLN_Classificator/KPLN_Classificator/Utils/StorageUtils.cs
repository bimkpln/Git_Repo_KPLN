using Autodesk.Revit.DB;
using KPLN_Classificator.Data;
using System;
using System.Collections.Generic;
using System.IO;
using static KPLN_Classificator.ApplicationConfig;

namespace KPLN_Classificator
{
    public class StorageUtils
    {
        public string xmlFilePath;

        public StorageUtils()
        {
            this.xmlFilePath = null;
        }
        public InfosStorage getInfoStorage()
        {
            System.Xml.Serialization.XmlSerializer storageSerializer =
                new System.Xml.Serialization.XmlSerializer(typeof(UtilsStorage));

            UtilsStorage utilsStorage = null;
            try
            {
                using (StreamReader r = new StreamReader(string.Format("C:\\TEMP\\ccsettings{0}.xml", RevitVersion)))
                {
                    utilsStorage = (UtilsStorage)storageSerializer.Deserialize(r);
                }
                System.Windows.Forms.OpenFileDialog storageDialog = new System.Windows.Forms.OpenFileDialog();
                storageDialog.InitialDirectory = utilsStorage.path;
                storageDialog.Multiselect = false;
                storageDialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
                if (storageDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    xmlFilePath = storageDialog.FileName;
                }
                else return null;
            }
            catch (Exception) { }

            if (utilsStorage == null)
            {
                string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string folder = System.IO.Path.GetDirectoryName(dllPath);
                System.Windows.Forms.OpenFileDialog dialog = new System.Windows.Forms.OpenFileDialog();
                dialog.InitialDirectory = folder;
                dialog.Multiselect = false;
                dialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
                if (dialog.ShowDialog() != System.Windows.Forms.DialogResult.OK) return null;
                xmlFilePath = dialog.FileName;
                utilsStorage = new UtilsStorage();
            }
            utilsStorage.path = Path.GetDirectoryName(xmlFilePath);

            InfosStorage storage = new InfosStorage();
            System.Xml.Serialization.XmlSerializer serializer =
                new System.Xml.Serialization.XmlSerializer(typeof(InfosStorage));

            using (StreamReader r = new StreamReader(xmlFilePath))
            {
                storage = (InfosStorage)serializer.Deserialize(r);
            }

            if (!Directory.Exists("C:\\TEMP"))
            {
                Directory.CreateDirectory("C:\\TEMP");
            }

            using (StreamWriter r = new StreamWriter(string.Format("C:\\TEMP\\ccsettings{0}.xml", RevitVersion)))
            {
                storageSerializer.Serialize(r, utilsStorage);
            }

            return storage;
        }

        public InfosStorage getStorageFromFilePath(string xmlFilePath)
        {
            this.xmlFilePath = xmlFilePath;
            InfosStorage storage = new InfosStorage();
            System.Xml.Serialization.XmlSerializer serializer =
                new System.Xml.Serialization.XmlSerializer(typeof(InfosStorage));

            using (StreamReader r = new StreamReader(xmlFilePath))
            {
                storage = (InfosStorage)serializer.Deserialize(r);
            }

            UtilsStorage utilsStorage = new UtilsStorage();
            utilsStorage.path = Path.GetDirectoryName(xmlFilePath);

            if (!Directory.Exists("C:\\TEMP"))
            {
                Directory.CreateDirectory("C:\\TEMP");
            }

            System.Xml.Serialization.XmlSerializer storageSerializer =
                 new System.Xml.Serialization.XmlSerializer(typeof(UtilsStorage));

            using (StreamWriter r = new StreamWriter(string.Format("C:\\TEMP\\ccsettings{0}.xml", RevitVersion)))
            {
                storageSerializer.Serialize(r, utilsStorage);
            }

            return storage;
        }

        public void saveInfoStorage(InfosStorage storage)
        {
            System.Xml.Serialization.XmlSerializer utilsSerializer =
                  new System.Xml.Serialization.XmlSerializer(typeof(UtilsStorage));
            System.Xml.Serialization.XmlSerializer infoStorageSerializer =
                 new System.Xml.Serialization.XmlSerializer(typeof(InfosStorage));

            UtilsStorage utilsStorage = null;
            try
            {
                using (StreamReader r = new StreamReader(string.Format("C:\\TEMP\\ccsettings{0}.xml", RevitVersion)))
                {
                    utilsStorage = (UtilsStorage)utilsSerializer.Deserialize(r);
                }
            }
            catch (Exception) { }

            if (utilsStorage == null)
            {
                string dllPath = System.Reflection.Assembly.GetExecutingAssembly().Location;
                string folder = System.IO.Path.GetDirectoryName(dllPath);
                utilsStorage = new UtilsStorage();
                utilsStorage.path = folder;
            }

            System.Windows.Forms.SaveFileDialog storageDialog = new System.Windows.Forms.SaveFileDialog();
            storageDialog.InitialDirectory = utilsStorage.path;
            storageDialog.FileName = xmlFilePath != null ? Path.GetFileName(xmlFilePath) : null;
            storageDialog.Filter = "xml files (*.xml)|*.xml|All files (*.*)|*.*";
            if (storageDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                xmlFilePath = storageDialog.FileName;
            }

            using (StreamWriter r = new StreamWriter(xmlFilePath))
            {
                infoStorageSerializer.Serialize(r, storage);
            }

            if (!Directory.Exists("C:\\TEMP"))
            {
                Directory.CreateDirectory("C:\\TEMP");
            }

            using (StreamWriter r = new StreamWriter(string.Format("C:\\TEMP\\ccsettings{0}.xml", RevitVersion)))
            {
                utilsSerializer.Serialize(r, utilsStorage);
            }
        }
    }
}
