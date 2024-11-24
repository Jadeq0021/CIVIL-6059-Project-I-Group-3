using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Media.Imaging;
using CreateModule.Properties;

/*******************************************************************************
* Function  :  加载Revit界面上方的按钮                                                                              *
* Author    :                                                                *
* Version   :  1.0                                                           *
* Date      :  2022年8月21日                                               *
*                                                                              *
*******************************************************************************/

namespace CodeAgent
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]
    [Autodesk.Revit.Attributes.Journaling(Autodesk.Revit.Attributes.JournalingMode.NoCommandData)]
    class CodeAgentRibbon : IExternalApplication
    {
        static string AddInPath = typeof(CodeAgentRibbon).Assembly.Location;//自动寻找dll文件
        public Result OnShutdown(UIControlledApplication application)//关闭Revit
        {
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)//启动Revit
        {

            string AddInPath = this.GetType().Assembly.Location;
            application.CreateRibbonTab("ModuleAgent");
            RibbonPanel panel = application.CreateRibbonPanel("ModuleAgent", "ModuleAgent");//增加一个新的面板
            PushButtonData Recher0 = new PushButtonData("CreateModule", "CreateModule", AddInPath, "CodeAgent.openWindow");

            System.Drawing.Bitmap bmp0 = Resources.ShadowButton;
            IntPtr hBitmap0 = bmp0.GetHbitmap();
            System.Windows.Media.ImageSource WpfBitmap0 = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap
                (hBitmap0, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
            Recher0.LargeImage = WpfBitmap0;


            panel.AddItem(Recher0);
          

            return Result.Succeeded;

        }
        public class WPFBitmapConverter : System.Windows.Data.IValueConverter
        {
            #region IValueConverter Members

            public object Convert(object value, Type targetType, object parameter,
                System.Globalization.CultureInfo culture)
            {
                MemoryStream ms = new MemoryStream();
                ((System.Drawing.Bitmap)value).Save(ms, System.Drawing.Imaging.ImageFormat.Bmp);
                BitmapImage image = new BitmapImage();
                image.BeginInit();
                ms.Seek(0, SeekOrigin.Begin);
                image.StreamSource = ms;
                image.EndInit();
                return image;
            }

            public object ConvertBack(object value, Type targetType, object parameter,
                System.Globalization.CultureInfo culture)
            {
                throw new NotImplementedException();
            }

            #endregion
        }
    }
}