using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace ToolsApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static readonly string _tmpWorkbookDir = @"C:\ProgramData\CatalogPrinter\exports";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Thumb_OnDragDelta(object sender, DragDeltaEventArgs e)
        {
            Left = Left + e.HorizontalChange;
            Top = Top + e.VerticalChange;
        }

        private void ButtonClose_OnClick(object sender, RoutedEventArgs e)
        {
            Application.Current.Shutdown();
        }

        private void KillExcel_OnClick(object sender, RoutedEventArgs e)
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Equals("EXCEL"))
                {
                    clsProcess.Kill();
                    break;
                }
            }
        }

        private void DeleteExports_OnClick(object sender, RoutedEventArgs e)
        {
            var dir = new System.IO.DirectoryInfo(_tmpWorkbookDir);
            var files = dir.GetFiles();
            foreach(var file in files)
            {
                if (file.Name.StartsWith("export") && file.Name.EndsWith("xlsx"))
                    File.Delete(file.FullName);
            }
        }
    }
}
