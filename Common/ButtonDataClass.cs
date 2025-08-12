using DocumentFormat.OpenXml.Presentation;
using System.Reflection;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Drawing;
using System.IO;
using System.Windows;
using WpfToolTip = System.Windows.Controls.ToolTip;

namespace HDR_EMEA.Common
{
    public class ButtonDataClass
    {
        public PushButtonData Data { get; set; }
        public WpfToolTip CustomToolTip { get; private set; }

        public ButtonDataClass(string name, string text, string className,
            byte[] largeImage,
            byte[] smallImage,
            string toolTip)
        {
            var status = ButtonStatus.Standard;
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;
            CustomToolTip = CreateToolTipContent(toolTip, status);

            Data.LargeImage = ConvertToImageSource(largeImage, status);
            Data.Image = ConvertToImageSource(smallImage, status);

            SetCommandAvailability();
        }

        public ButtonDataClass(string name, string text, string className,
            byte[] largeImage,
            byte[] smallImage,
            byte[] largeImageDark,
            byte[] smallImageDark,
            string toolTip)
        {
            var status = ButtonStatus.Standard;
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;
            CustomToolTip = CreateToolTipContent(toolTip, status);

            UITheme theme = UIThemeManager.CurrentTheme;
            if (theme == UITheme.Dark)
            {
                Data.LargeImage = ConvertToImageSource(largeImageDark, status);
                Data.Image = ConvertToImageSource(smallImageDark, status);
            }
            else
            {
                Data.LargeImage = ConvertToImageSource(largeImage, status);
                Data.Image = ConvertToImageSource(smallImage, status);
            }

            SetCommandAvailability();
        }

        public ButtonDataClass(string name, string text, string className,
            byte[] largeImage,
            byte[] smallImage,
            ButtonStatus status,
            string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;
            CustomToolTip = CreateToolTipContent(toolTip, status);

            Data.LargeImage = ConvertToImageSource(largeImage, status);
            Data.Image = ConvertToImageSource(smallImage, status);

            SetCommandAvailability();
        }

        private void SetCommandAvailability()
        {
            string nameSpace = this.GetType().Namespace;
            Data.AvailabilityClassName = $"{nameSpace}.CommandAvailability";
        }

        public static Assembly GetAssembly()
        {
            return Assembly.GetExecutingAssembly();
        }

        public static string GetAssemblyName()
        {
            return GetAssembly().Location;
        }

        public enum ButtonStatus
        {
            Standard,
            New,
            Update
        }

        public static BitmapImage ConvertToImageSource(byte[] imageData, ButtonStatus status = ButtonStatus.Standard)
        {
            using (MemoryStream mem = new MemoryStream(imageData))
            using (Bitmap baseImage = new Bitmap(mem))
            {
                if (status != ButtonStatus.Standard)
                {
                    using (Graphics g = Graphics.FromImage(baseImage))
                    {
                        int dotSize = 9;
                        int margin = 1;
                        System.Drawing.Color dotColour = status == ButtonStatus.New
                            ? System.Drawing.Color.LimeGreen
                            : System.Drawing.Color.Orange;

                        using (System.Drawing.Brush brush = new SolidBrush(dotColour))
                        {
                            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                            g.FillEllipse(brush,
                                baseImage.Width - dotSize - margin,
                                margin,
                                dotSize,
                                dotSize);
                        }
                    }
                }

                using (MemoryStream ms = new MemoryStream())
                {
                    baseImage.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                    ms.Position = 0;

                    BitmapImage bmi = new BitmapImage();
                    bmi.BeginInit();
                    bmi.StreamSource = ms;
                    bmi.CacheOption = BitmapCacheOption.OnLoad;
                    bmi.EndInit();
                    bmi.Freeze();

                    return bmi;
                }
            }
        }

        private static WpfToolTip CreateToolTipContent(string toolTipText, ButtonStatus status)
        {
            var tooltip = new WpfToolTip();
            var stackPanel = new StackPanel();

            if (status != ButtonStatus.Standard)
            {
                var label = new TextBlock
                {
                    Text = status == ButtonStatus.New ? "NEW" : "UPDATED",
                    Background = new SolidColorBrush(status == ButtonStatus.New ? Colors.LimeGreen : Colors.Orange),
                    Foreground = System.Windows.Media.Brushes.White,
                    FontWeight = FontWeights.Bold,
                    Padding = new Thickness(6, 2, 6, 2),
                    Margin = new Thickness(0, 0, 0, 6),
                    HorizontalAlignment = System.Windows.HorizontalAlignment.Left
                };
                stackPanel.Children.Add(label);
            }

            var body = new TextBlock
            {
                Text = toolTipText,
                TextWrapping = TextWrapping.Wrap,
                MaxWidth = 300
            };

            stackPanel.Children.Add(body);
            tooltip.Content = stackPanel;

            return tooltip;
        }
    }
}
