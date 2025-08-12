using System.Drawing;
using System.IO;

namespace HDR_EMEA.Common
{
    public static class Helpers
    {
        public static byte[] ImageToBytes(Bitmap image)
        {
            using (var stream = new MemoryStream())
            {
                image.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return stream.ToArray();
            }
        }
    }
}
