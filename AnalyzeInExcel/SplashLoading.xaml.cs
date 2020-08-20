using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace AnalyzeInExcel
{
    /// <summary>
    /// Interaction logic for SplashLoading.xaml
    /// </summary>
    public partial class SplashLoading : Window
    {
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                this.X = x;
                this.Y = y;
            }
        }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetCursorPos(out POINT pPoint);

        public static int SM_CXSCREEN = 0;  // GetSystemMetrics index.
        [DllImport("USER32.DLL", SetLastError = true)]
        public static extern int GetSystemMetrics(int nIndex);

        public SplashLoading()
        {
            InitializeComponent();

            Point ptMouse = GetCursorPosition();

            Left = ptMouse.X;
            Top = ptMouse.Y;
        }

        private static Point GetCursorPosition()
        {
            GetCursorPos(out POINT cursorScreenPosition);

            double widthInDevicePixels = SplashLoading.GetSystemMetrics(SplashLoading.SM_CXSCREEN);
            double widthInDIP = SystemParameters.WorkArea.Right; // Device independent pixels.
            double scalingFactor = widthInDIP / widthInDevicePixels;

            var ptMouse = new System.Windows.Point(cursorScreenPosition.X * scalingFactor, cursorScreenPosition.Y * scalingFactor);
            return ptMouse;
        }
    }
}
