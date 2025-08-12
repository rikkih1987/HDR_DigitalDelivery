using Autodesk.Revit.UI;
using System.Reflection;

namespace HDR_EMEA.Tools_Core
{
    public static class About
    {
        public static PushButtonData GetButtonData()
        {
            string assemblyPath = Assembly.GetExecutingAssembly().Location;

            return new PushButtonData("About", "About", assemblyPath, typeof(AboutCommand).FullName)
            {
                ToolTip = "About HDR EMEA Add-In"
            };
        }
    }
}
