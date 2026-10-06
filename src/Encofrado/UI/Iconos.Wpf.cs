using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Imaging;
using Encofrado.Core;
using ImageSource = System.Windows.Media.ImageSource;
using Comun;

namespace Encofrado.UI
{
    /// <summary>Conversión de los íconos a imágenes WPF para la cinta de Revit.</summary>
    internal static partial class Iconos
    {
        public static ImageSource Ribbon(Pieza kind, int size) => ToImageSource(Bitmap(size, g => Draw(kind, g, size)));
        public static ImageSource Ribbon(Extra kind, int size) => ToImageSource(Bitmap(size, g => Draw(kind, g, size)));

        public static ImageSource Tooltip(Pieza kind) => ToImageSource(Bitmap(160, g =>
        {
            g.Clear(Color.White);
            Draw(kind, g, 160);
        }));

        private static ImageSource ToImageSource(Bitmap bmp)
        {
            using (bmp)
            using (var ms = new MemoryStream())
            {
                bmp.Save(ms, ImageFormat.Png);
                ms.Position = 0;
                var img = new BitmapImage();
                img.BeginInit();
                img.CacheOption = BitmapCacheOption.OnLoad;
                img.StreamSource = ms;
                img.EndInit();
                img.Freeze();
                return img;
            }
        }
    }
}
