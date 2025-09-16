using System;

namespace KPLN_Quantificator.Forms
{
    internal static class Output
    {
        public static OutputForm OutputForm
        {
            get => OutputForm.GetInstance();
        }

        public static void PrintError(Exception e)
        {
            string message = string.Format("Report: {0}:\n{1};\n---", e.Message.ToString(), e.StackTrace.ToString());
            OutputForm.Show();
            OutputForm.AddErrorTextBlock(message);

        }

        public static void PrintHeader(string message)
        {

            OutputForm.Show();
            OutputForm.AddHeaderTextBlock(message);

        }

        public static void Print(string message)
        {
            OutputForm.Show();
            OutputForm.AddTextBlock(message);
        }
        public static void PrintAlert(string message)
        {
            OutputForm.Show();
            OutputForm.AddErrorTextBlock(message);
        }

        public static void PrintSuccess(string message, string boldKey = null, int? boldTotalGroups = null, int? boldIosGroups = null)
        {
            OutputForm.Show();
            OutputForm.AddSuccessTextBlock(message, boldKey, boldTotalGroups, boldIosGroups);
        }
    }
}
