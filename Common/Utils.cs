using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System.Linq;
using System.IO;

namespace HDR_EMEA.Common
{
    internal static class Utils
    {
        internal static RibbonPanel CreateRibbonPanel(UIControlledApplication app, string tabName, string panelName)
        {
            RibbonPanel curPanel;

            if (GetRibbonPanelByName(app, tabName, panelName) == null)
                curPanel = app.CreateRibbonPanel(tabName, panelName);
            else
                curPanel = GetRibbonPanelByName(app, tabName, panelName);

            return curPanel;
        }

        internal static RibbonPanel GetRibbonPanelByName(UIControlledApplication app, string tabName, string panelName)
        {
            foreach (RibbonPanel tmpPanel in app.GetRibbonPanels(tabName))
            {
                if (tmpPanel.Name == panelName)
                    return tmpPanel;
            }

            return null;
        }

        internal static string GetParameterValue(Element element, string paramName)
        {
            Parameter parm = element.LookupParameter(paramName);
            return parm != null ? parm.AsValueString() : "N/A";
        }

        internal static bool SetParameter(Element element, string paramName, double valueInMillimeters)
        {
            Parameter param = element.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly && param.StorageType == StorageType.Double)
            {
                double internalValue = UnitUtils.ConvertToInternalUnits(valueInMillimeters, UnitTypeId.Millimeters);
                if (!MathComparisonUtils.IsAlmostEqual(param.AsDouble(), internalValue))
                {
                    return param.Set(internalValue);
                }
            }
            return false;
        }

        internal static bool SetParameter(Element element, string paramName, string value)
        {
            Parameter param = element.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly && param.StorageType == StorageType.String)
            {
                if (param.AsString() != value)
                {
                    return param.Set(value);
                }
            }
            return false;
        }

        private static byte[] ToByteArray(Stream stream)
        {
            using (var memoryStream = new MemoryStream())
            {
                stream.CopyTo(memoryStream);
                return memoryStream.ToArray();
            }
        }

        internal static System.Windows.Media.ImageSource LoadPngImage(string embeddedPath)
        {
            var assembly = System.Reflection.Assembly.GetExecutingAssembly();
            using var stream = assembly.GetManifestResourceStream(embeddedPath);
            if (stream == null) return null;

            var decoder = new System.Windows.Media.Imaging.PngBitmapDecoder(
                stream,
                System.Windows.Media.Imaging.BitmapCreateOptions.PreservePixelFormat,
                System.Windows.Media.Imaging.BitmapCacheOption.OnLoad
            );

            return decoder.Frames[0];
        }

        internal static byte[] ConvertBitmapToByteArray(System.Drawing.Bitmap bitmap)
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, System.Drawing.Imaging.ImageFormat.Png);
                return stream.ToArray();
            }
        }
    }
}
