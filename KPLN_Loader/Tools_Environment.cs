using System;
using System.IO;
using System.Linq;
using System.Threading;
using static KPLN_Loader.Output.Output;
using static KPLN_Loader.Preferences;

namespace KPLN_Loader
{
    public class Tools_Environment
    {
        public string RevitVersion { get; }
        private readonly DirectoryInfo UserLocation = new DirectoryInfo(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), @"AppData\Local"));
        public DirectoryInfo ApplicationLocation { get; }
        private DirectoryInfo SessionLocation { get; }
        public DirectoryInfo ModulesLocation { get; }
        public Tools_Environment(string revitVersion)
        {
            RevitVersion = revitVersion;
            ApplicationLocation = new DirectoryInfo(Path.Combine(UserLocation.FullName, "KPLN_Loader"));
            SessionLocation = new DirectoryInfo(Path.Combine(ApplicationLocation.FullName, string.Format(@"{0}_{1}", RevitVersion, Guid.NewGuid().ToString())));
            ModulesLocation = new DirectoryInfo(Path.Combine(SessionLocation.FullName, "Modules"));
        }
        public bool PrepareLocalDirectory()
        {
            try
            {
                PrepareLocation(ApplicationLocation, true);
                PrepareLocation(SessionLocation, false);
                PrepareLocation(ModulesLocation, false);
                return true;
            }
            catch (Exception e)
            {
                PrintError(e);
                return false;
            }
        }
        public void ClearPreviousLog()
        {
            string outputPath = string.Format(@"{0}\log_{1}.html", Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), Guid.NewGuid());
            DirectoryInfo userdoc = new DirectoryInfo(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments));
            foreach (FileInfo file in userdoc.EnumerateFiles())
            {
                if (file.CreationTime.Date.Day != DateTime.Now.Day && file.Name.StartsWith("log_") && file.Name.EndsWith(".html"))
                {
                    if (!FileIsBusy(file))
                    {
                        try
                        {
                            file.Delete();
                        }
                        catch (Exception)
                        { }
                    }
                }
            }
        }
        private bool PrepareLocation(DirectoryInfo loc, bool clearFreeElements)
        {
            if (IsDirectoryExist(loc))
            {
                if (clearFreeElements)
                {
                    foreach (DirectoryInfo subLoc in loc.GetDirectories())
                    {
                        if (!BusyFilesInDirectory(subLoc))
                        { 
                            ClearDirectory(subLoc);
                            subLoc.Delete();
                        }
                    }
                }
                return true;
            }
            else
            {
                CreateDirectory(loc);
                return true;
            }
        }
        public DirectoryInfo CopyModuleFromPath(DirectoryInfo path, string version, string name)
        {
            DirectoryInfo moduleDirectory = Directory.CreateDirectory(Path.Combine(ModulesLocation.FullName, path.Name));
            DirectoryCopy(path.FullName, Path.Combine(ModulesLocation.FullName, path.Name), true);
            Print(string.Format("Инфо: Модуль «{1}» получен и готов к активации [версия модуля: {0}]", version, name), MessageType.System_Regular);
            return new DirectoryInfo(Path.Combine(ModulesLocation.FullName, path.Name));
        }
        public bool IsDirectoryExist(DirectoryInfo path)
        {
            if (Directory.Exists(ApplicationLocation.FullName))
            {
                return true;
            }
            return false;
        }
        private void CreateDirectory(DirectoryInfo path)
        {
            try
            {
                if (!IsDirectoryExist(path))
                {
                    DirectoryInfo dir = Directory.CreateDirectory(path.FullName);
                }
            }
            catch (Exception e)
            {
                PrintError(e);
            }
        }
        private bool BusyFilesInDirectory(DirectoryInfo dirPath)
        {
            foreach (FileInfo file in dirPath.GetFiles())
            { 
                if (FileIsBusy(file)) { return true; }
            }
            foreach (DirectoryInfo folder in dirPath.GetDirectories())
            {
                if (BusyFilesInDirectory(folder)) { return true; }
            }
            return false;
        }
        private void ClearModulesDirectory(int loop = 0)
        {
            try
            {
                ClearDirectory(ModulesLocation);
            }
            catch (Exception)
            {
                if (loop < 5)
                {
                    Thread.Sleep(5000);
                    ClearModulesDirectory(loop++);
                }
                else
                { }
            }
        }
        private void ClearDirectory(DirectoryInfo dirPath)
        {
            try
            {
                foreach (FileInfo file in dirPath.GetFiles())
                {
                    try
                    {
                        file.Delete();
                    }
                    catch (Exception) { }
                }
                foreach (DirectoryInfo folder in dirPath.GetDirectories())
                {
                    try
                    {
                        if (folder.GetFiles().Count() == 0 && folder.GetDirectories().Count() == 0) { folder.Delete(); }
                        else { ClearDirectory(folder); }
                        folder.Delete();
                    }
                    catch (Exception) { }
                }
            }
            catch (Exception) { }
        }
        private void DirectoryCopy(string sourceDirName, string destDirName, bool copySubDirs)
        {
            DirectoryInfo dir = new DirectoryInfo(sourceDirName);

            if (!dir.Exists)
            {
                throw new DirectoryNotFoundException();
            }

            DirectoryInfo[] dirs = dir.GetDirectories();
            if (!Directory.Exists(destDirName))
            {
                Directory.CreateDirectory(destDirName);
                Print(string.Format("Создание: {0}", destDirName), MessageType.System_Regular);
            }
            FileInfo[] files = dir.GetFiles();
            foreach (FileInfo file in files)
            {
                string temppath = Path.Combine(destDirName, file.Name);
                file.CopyTo(temppath, false);
                Print(string.Format("Копирование: {0}", temppath), MessageType.System_Regular);
            }
            if (copySubDirs)
            {
                foreach (DirectoryInfo subdir in dirs)
                {
                    string temppath = Path.Combine(destDirName, subdir.Name);
                    DirectoryCopy(subdir.FullName, temppath, copySubDirs);
                }
            }
        }
        private bool FileIsBusy(FileInfo fileName)
        {
            try
            {
                using (FileStream stream = fileName.Open(FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    stream.Close();
                }
            }
            catch (Exception){ return true; }
            return false;
        }
    }
}
