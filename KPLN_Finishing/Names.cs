using System.IO;

namespace KPLN_Finishing
{
    public static class Names
    {
        public static string assembly { get; set; }
        public static string assembly_Path { get; set; }
        //
        public static string parameter_Room_Id = "О_Id помещения";
        public static string parameter_Room_Name = "О_Имя помещения";
        public static string parameter_Room_Number = "О_Номер помещения";
        public static string parameter_Room_Section = "О_Номер секции";
        public static string parameter_Room_Department = "О_Назначение помещения";
        //
        public static string value_All_Model_Model = "отделка";
        //
        public static string message_Matrix_Reset = "matrix reset";
        public static string message_Matrix_Optimising = "matrix optimizing";
        public static string message_Matrix_Adding_Element = "added to marix";
        public static string message_Matrix_Preparing_Conainers = "preparing containers";
        public static string message_Matrix_Squarifying = "squrifying bounds";
        public static string message_Element_Filter = "does not match any condition";
        public static string message_Element_Geometry_Filter = "geometry error";
        public static string message_Element_Calculated_Single = "calculeted by single";
        public static string message_Element_Calculated_Nearest = "calculated by multiple";
        public static string message_Element_Calculated_Chain = "calculated from chain";
        public static string message_Element_NotCalculated = "not calculated";
        //
        public static string task_dialog_hint = "см. Moodle | Плaгины и инструменты автоматизации Rеvit | KPLN | АР. Отделка";
        //
        public static string shared_Parameters_File_Path()
        {
            return assembly_Path + @"\Source\ФОП2019_КП_АР.txt";

        }
        public static string shared_Parameters_File_Group = "АРХИТЕКТУРА - Отделка";
    }
}
