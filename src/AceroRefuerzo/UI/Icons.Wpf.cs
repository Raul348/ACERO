using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Media.Imaging;
using AceroRefuerzo.Core;
using ImageSource = System.Windows.Media.ImageSource;

namespace AceroRefuerzo.UI
{
    /// <summary>Conversión de los íconos a imágenes WPF para la cinta de Revit.</summary>
    internal static partial class Icons
    {
        public static ImageSource Ribbon(ElementKind kind, int size) => ToImageSource(Bitmap(size, g => Draw(kind, g, size)));
        public static ImageSource Ribbon(Extra kind, int size) => ToImageSource(Bitmap(size, g => Draw(kind, g, size)));

        public static ImageSource Tooltip(ElementKind kind) => ToImageSource(Bitmap(160, g =>
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
