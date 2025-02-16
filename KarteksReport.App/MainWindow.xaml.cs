using DocumentFormat.OpenXml.Wordprocessing;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Services;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace KarteksReport.App
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int GWL_STYLE = -16;
        private const int WS_SYSMENU = 0x80000;
        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);




        private readonly StartApp _startApp;

        public MainWindow()
        {
            _startApp = new StartApp();

            InitializeComponent();

            //Start();
        }

        private async void Start(object sender, RoutedEventArgs e)
        {
            var hwnd = new System.Windows.Interop.WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) & ~WS_SYSMENU);

            txtBlock.Text = "Подождите. Идет формирование отчетов!";

            await _startApp.StartExamined(pBar);

            await _startApp.StartWorkload(pBar);

            txtBlock.Text = StatusReport.Status;
            btnClose.IsEnabled = true;
        }

        private void Bar()
        {
            while (true)
            {
                if (pBar.Value == 10)
                {
                    txtBlock.Text = StatusReport.Status + pBar.Value;
                    btnClose.IsEnabled = true;

                    return;
                }
            }
        }

        private void btn_Close(object sender, RoutedEventArgs e)
        {
            _= ReportFile.WriteReportsAsync();
            this.Close();
        }

        protected override void OnClosing(CancelEventArgs e)
        {
            _ = ReportFile.WriteReportsAsync();
            base.OnClosing(e);
        }
    }
}