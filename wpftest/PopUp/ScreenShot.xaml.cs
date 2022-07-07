using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace WizMes_WooJung.PopUp
{
    /// <summary>
    /// ScreenShot.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class ScreenShot : Window
    {
        public ScreenShot()
        {
            InitializeComponent();
        }

        private void ScreenShot_Loaded(object sender, RoutedEventArgs e)
        {

            //품명별 불량분석 화면에서 보낸 image소스를 받는다.
            if (MainWindow.ScreenCapture != null && MainWindow.ScreenCapture.Count > 0)
            {
                //받아서 IMAGEDATA에 넣었는데 과연 나올까??
                ImageData.Source = MainWindow.ScreenCapture[0].Source;
            }
        }        
    }
}
