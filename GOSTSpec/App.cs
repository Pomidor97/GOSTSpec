#region Namespaces
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Media.Imaging;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
#endregion

namespace GOSTSpec
{
    class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication a)
        {
            string path = Assembly.GetExecutingAssembly().Location;
            var tabName = "KAZGOR";
            // create custom ribbon tab:            
            try
            {
                a.CreateRibbonTab(tabName);
            }
            catch (Exception) { }

            RibbonPanel panel = a.GetRibbonPanels().SingleOrDefault((pl) => pl.Name == "Оформление");
            if (panel == null)
            {
                panel = a.CreateRibbonPanel(tabName, "Оформление");
            }
            PushButton pushButton1 = panel.AddItem(new PushButtonData("copyParameterValues_btn",
        "Копирование\nзначение параметров", path, "GOSTSpec.Command")) as PushButton;
            PushButton pushButton2 = panel.AddItem(new PushButtonData("autoNumbering_btn",
        "Автонумерация\nпозиций", path, "GOSTSpec.Numbering")) as PushButton;            
            PushButton pushButton3 = panel.AddItem(new PushButtonData("autoScheduling_btn",
        "Авто\nспецификация", path, "GOSTSpec.AutoSchedule")) as PushButton;

            // Set ToolTip and contextual help
            pushButton1.ToolTip = "Копирование значение параметров";
            pushButton2.ToolTip = "Автонумерация позиций";
            pushButton3.ToolTip = "Автоматическое создание спецификации";

            Stream stream1 = Assembly.GetExecutingAssembly().GetManifestResourceStream("GOSTSpec.Resources.icons8_copy_32.png");
            var decoder1 = new PngBitmapDecoder(stream1, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            Stream stream2 = Assembly.GetExecutingAssembly().GetManifestResourceStream("GOSTSpec.Resources.icons8_counter_32.png");
            var decoder2 = new PngBitmapDecoder(stream2, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);

            Stream stream3 = Assembly.GetExecutingAssembly().GetManifestResourceStream("GOSTSpec.Resources.icons8_schedule_32.png");
            var decoder3 = new PngBitmapDecoder(stream3, BitmapCreateOptions.PreservePixelFormat, BitmapCacheOption.Default);



            pushButton1.LargeImage = decoder1.Frames[0];
            pushButton2.LargeImage = decoder2.Frames[0];
            pushButton3.LargeImage = decoder3.Frames[0];


            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication a)
        {
            return Result.Succeeded;
        }
    }
}
