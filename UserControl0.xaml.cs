using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using CreateModule.Properties;
using System.IO;


namespace CodeAgent
{
    public partial class UserControl : Window
    {
        public string name = "初期支护";
        Document document = null;
        View view = null;
        //柱檩条创建事件
        ExternalEvent externalEventQuzheng = null;
        MainClassEvent openWindow_BeamEvent = new MainClassEvent();

        public double area = 0;
        public double weight = 0;
        //========================================================================
        PreviewControl previewControl = null;
        public UserControl(Document document, View view)
        {
            //显示revit界面
            this.document = document;
            this.view = view;

            InitializeComponent();
            System.Drawing.Bitmap bmp1 = CreateModule.Properties.Resources.A1;
            System.Drawing.Bitmap bmp2 = CreateModule.Properties.Resources.A2;
            System.Drawing.Bitmap bmp3 = CreateModule.Properties.Resources.A3;
            System.Drawing.Bitmap bmp4 = CreateModule.Properties.Resources.A4;
            System.Drawing.Bitmap bmp5 = CreateModule.Properties.Resources.A5;
            System.Drawing.Bitmap bmp6 = CreateModule.Properties.Resources.A6;
            System.Drawing.Bitmap bmp7 = CreateModule.Properties.Resources.All;


            externalEventQuzheng = ExternalEvent.Create(openWindow_BeamEvent);
            Image1.Source = imageget(bmp1);
            Image2.Source = imageget(bmp2);
            Image3.Source = imageget(bmp3);
            Image4.Source = imageget(bmp4);
            Image5.Source = imageget(bmp5);
            Image6.Source = imageget(bmp6);
            Image7.Source = imageget(bmp7);



        }

        public System.Windows.Media.ImageSource imageget(System.Drawing.Bitmap bmp5)
        {
            IntPtr hBitmap5 = bmp5.GetHbitmap();
            return System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(hBitmap5, IntPtr.Zero, System.Windows.Int32Rect.Empty, BitmapSizeOptions.FromEmptyOptions());
        }



        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao.Text == "请输入，平米为单位")
            {
                    openWindow_BeamEvent.area = 40;
                    openWindow_BeamEvent.wallThickness = 0.2;
                    externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.model = 0;
                openWindow_BeamEvent.area = Convert.ToDouble(this.zhanghao.Text);
                openWindow_BeamEvent.wallThickness = Convert.ToDouble(this.mima.Text);
                externalEventQuzheng.Raise();

            }

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
        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // 当Slider的值发生变化时，更新TextBlock的文本
            this.zhanghao.Text = Slider1.Value.ToString();
        }
        private void Slider_ValueChanged_1(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            // 当Slider的值发生变化时，更新TextBlock的文本
            this.mima.Text = Slider2.Value.ToString();
        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
                openWindow_BeamEvent.model = 10;
                externalEventQuzheng.Raise();
        }

        private void Button_Click_a1(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_1.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_1.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_1.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_1.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_1.Text);
                openWindow_BeamEvent.model = 1;
                externalEventQuzheng.Raise();
            }
        }

        private void Button_Click_a2(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_2.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_2.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_2.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_2.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_2.Text);
                openWindow_BeamEvent.model = 2;
                externalEventQuzheng.Raise();
            }
        }

        private void Button_Click_a3(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_1.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_3.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_3.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_3.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_3.Text);
                openWindow_BeamEvent.model = 3;
                externalEventQuzheng.Raise();
            }
        }

        private void Button_Click_a4(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_1.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_4.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_4.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_4.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_4.Text);
                openWindow_BeamEvent.model = 4;
                externalEventQuzheng.Raise();
            }
        }

        private void Button_Click_a5(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_1.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_5.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_5.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_5.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_5.Text);
                openWindow_BeamEvent.model = 5;
                externalEventQuzheng.Raise();
            }
        }

        private void Button_Click_a6(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_1.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_6.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_6.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_6.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_6.Text);
                openWindow_BeamEvent.model = 6;
                externalEventQuzheng.Raise();
            }
        }

        private void Button_Click_c(object sender, RoutedEventArgs e)
        {
            if (this.zhanghao复制__C_1.Text == "请输入，X轴宽")
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_c.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_c.Text);
                externalEventQuzheng.Raise();

            }
            else
            {
                openWindow_BeamEvent.wide = Convert.ToDouble(this.zhanghao复制__C_c.Text);
                openWindow_BeamEvent.height = Convert.ToDouble(this.mima复制__C_c.Text);
                openWindow_BeamEvent.model = 7;
                externalEventQuzheng.Raise();
            }
        }

    
    }
}
