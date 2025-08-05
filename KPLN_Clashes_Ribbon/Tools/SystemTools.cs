using System.Drawing;
using System.IO;

namespace KPLN_Clashes_Ribbon.Tools
{
    internal static class SystemTools
    {
        public static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[64 * 1024];
            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }

        public static Image ByteArrayToImage(byte[] byteArrayIn)
        {
            Image result = null;
            try
            {
                MemoryStream ms = new MemoryStream(byteArrayIn, 0, byteArrayIn.Length);
                ms.Write(byteArrayIn, 0, byteArrayIn.Length);
                result = Image.FromStream(ms, true);
            }
            catch { }
            return result;
        }
    }
}
