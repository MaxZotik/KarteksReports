using DocumentFormat.OpenXml.Wordprocessing;
using ReportsLibraryOpenXmlSdk.Entities;
using ReportsLibraryOpenXmlSdk.Services;
using System.ComponentModel;
using System.Runtime.CompilerServices;
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
        private readonly StartApp _startApp;

        public MainWindow()
        {
            _startApp = new StartApp();

            InitializeComponent();

            _= Start();
        }

        private async Task Start()
        {
            txtBlock.Text = StatusReport.Status;

            await _startApp.StartExamined(pBar);

            await _startApp.StartWorkload(pBar);

            txtBlock.Text = StatusReport.Status;
            btnClose.IsEnabled = true;
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