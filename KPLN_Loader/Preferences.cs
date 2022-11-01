using KPLN_Loader.Common;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Windows;

namespace KPLN_Loader
{
    public static class Preferences
    {
#if Revit2020 || Revit2022
        public static Window RevitWindow { get; set; }
#endif
        public static System.IntPtr MainWindowHandle { get; set; }
        
        public static Tools_Environment Tools { get; set; }
        
        public static Queue<IExecutableCommand> ExecutableCommands = new Queue<IExecutableCommand>();
        
        public static List<IExternalModule> LoadedModules = new List<IExternalModule>();
        
        public static readonly Queue<IExecutableCommand> CommandQueue = new Queue<IExecutableCommand>();
        
        /// <summary>
        /// Путь к сборке
        /// </summary>
        public static string AssemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString();

        public static string HTML_Output_DocType = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.01 Transitional//EN' 'http://www.w3.org/TR/html4/loose.dtd'>";
        
        public static string HTML_Output_Head 
        { 
            get 
            { 
                return string.Format(@"<head><meta http-equiv='X-UA-Compatible' content='IE=9'><meta http-equiv='content-type' content='text/html; charset=utf-8'><meta name='appversion' content='0.2.0.0'><link href='file:///{0}\Output\Styles\outputstyles.css' rel='stylesheet'></head>", Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString()); 
            } 
        }
        
        public static SQLUserInfo User;
        
        public static List<SQLUserInfo> Users = new List<SQLUserInfo>();
        
        public static List<SQLProjectInfo> User_Projects = new List<SQLProjectInfo>();
        
        /// <summary>
        /// Логирование ошибок
        /// </summary>
        public static NLog.Logger Logger = NLog.LogManager.LoadConfiguration(AssemblyPath + "\\nlog.config").GetCurrentClassLogger();
        
        public enum MessageType 
        { 
            Header, 
            Regular, 
            Error, 
            Warning, 
            Critical, 
            Code, 
            Success, 
            System_OK, 
            System_Regular 
        }
        
        public enum MessageDialogType 
        { 
            Regular, 
            Warning, 
            Important, 
            Fun, 
            System 
        }
        
        public enum MessageDialogResult 
        { 
            Read, 
            Ok, 
            Thx, 
            Close, 
            None, 
            Pending, 
            Shown 
        }
    }
}
