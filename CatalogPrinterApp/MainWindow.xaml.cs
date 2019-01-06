using Encrypter;
using ExcelUtil;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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

namespace CatalogPrinterApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string _configPath = @"C:\ProgramData\CatalogPrinter\CatalogPrinterApp.config";

        public MainWindow()
        {
            InitializeComponent();
        }

        private void Minimize_OnClick(object sender, RoutedEventArgs e)
        {
            WindowState = WindowState.Minimized;
        }

        private void Settings_OnClick(object sender, RoutedEventArgs e)
        {
            if (!File.Exists(_configPath))
                throw new Exception($"Config file " + _configPath + " not found!");
            ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
            configMap.ExeConfigFilename = _configPath;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
            var appSettings = config.GetSection("appSettings") as AppSettingsSection;
            // get config value
            string pathEncryptorApp = appSettings.Settings["pathEncryptorApp"].Value;

            ProcessStartInfo startInfo = new ProcessStartInfo();
            startInfo.UseShellExecute = true;
            startInfo.WorkingDirectory = Environment.CurrentDirectory;
            startInfo.FileName = pathEncryptorApp;
            startInfo.Verb = "runas";

            try
            {
                Process p = Process.Start(startInfo);
            }
            catch (System.ComponentModel.Win32Exception ex)
            {
                return;
            }
        }

        private void Thumb_OnDragDelta(object sender, DragDeltaEventArgs e)
        {
            Left = Left + e.HorizontalChange;
            Top = Top + e.VerticalChange;
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            ExcelUtility.CloseExcel();
            Application.Current.Shutdown();
        }

        private AppParameters GetConfigParameters()
        {
            var settings = new AppParameters();

            // open config
            if (!File.Exists(_configPath))
                throw new Exception($"Config file " + _configPath + " not found!");
            ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
            configMap.ExeConfigFilename = _configPath;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
            var appSettings = config.GetSection("appSettings") as AppSettingsSection;
            // get config values
            settings.hash = appSettings.Settings["key"]?.Value;
            settings.masterCatalog = appSettings.Settings["masterCatalog"]?.Value;
            settings.outputPath = appSettings.Settings["outputPath"]?.Value;

            var ranges = new InputRanges();
            ranges.catalogType = appSettings.Settings["cellCatalogType"]?.Value;
            ranges.btw = appSettings.Settings["cellBtw"]?.Value;
            ranges.footerRight = appSettings.Settings["cellFooterRight"]?.Value;
            ranges.footerLeft = appSettings.Settings["cellFooterLeft"]?.Value;
            ranges.footerMidFirst = appSettings.Settings["cellFooterMidFirst"]?.Value;
            ranges.footerMidSecond = appSettings.Settings["cellFooterMidSecond"]?.Value;
            ranges.headerRight = appSettings.Settings["cellHeaderRight"]?.Value;
            ranges.headerLeft = appSettings.Settings["cellHeaderLeft"]?.Value;
            ranges.headerMid = appSettings.Settings["cellHeaderMid"]?.Value;
            ranges.printArea = appSettings.Settings["printArea"]?.Value;

            settings.ranges = ranges;

            settings.sheetSummaryName = appSettings.Settings["sheetSummaryName"]?.Value;

            if (settings.hash == null)
                throw new Exception($"Could not find 'password' key in " + _configPath + "!");
            if (settings.masterCatalog == null)
                throw new Exception($"Could not find 'masterCatalog' key in " + _configPath + "!");
            if (settings.outputPath == null)
                throw new Exception($"Could not find 'outputPath' key in " + _configPath + "!");

            if (ranges.catalogType == null)
                throw new Exception($"Could not find 'cellCatalogType' key in " + _configPath + "!");
            if (ranges.btw == null)
                throw new Exception($"Could not find 'cellBtw' key in " + _configPath + "!");
            if (ranges.footerRight == null)
                throw new Exception($"Could not find 'cellFooterRight' key in " + _configPath + "!");
            if (ranges.footerLeft == null)
                throw new Exception($"Could not find 'cellFooterLeft' key in " + _configPath + "!");
            if (ranges.footerMidFirst == null)
                throw new Exception($"Could not find 'cellFooterMidFirst' key in " + _configPath + "!");
            if (ranges.footerMidSecond == null)
                throw new Exception($"Could not find 'cellFooterMidSecond' key in " + _configPath + "!");
            if (ranges.headerRight == null)
                throw new Exception($"Could not find 'cellHeaderRight' key in " + _configPath + "!");
            if (ranges.headerLeft == null)
                throw new Exception($"Could not find 'cellHeaderLeft' key in " + _configPath + "!");
            if (ranges.headerMid == null)
                throw new Exception($"Could not find 'cellHeaderMid' key in " + _configPath + "!");
            if (ranges.printArea == null)
                throw new Exception($"Could not find 'printArea' key in " + _configPath + "!");

            if (settings.sheetSummaryName == null)
                throw new Exception($"Could not find 'sheetSummaryName' key in " + _configPath + "!");

            return settings;
        }

        private async void PrintAllButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // progress
                var progress = new Progress<int>(value => ProgressBar.Value = value);

                // get catalog type
                var catalogType = ((ComboBoxItem)InputCatalogType.SelectedItem).Content.ToString();

                // btw
                bool inclBtw = InputBTW.IsChecked ?? false;

                // print
                await Task.Run(() => PrintCatalog(null, progress, catalogType, inclBtw));
            }
            catch (Exception ex)
            {
                ProgressBar.Value = 0;
                MessageBox.Show(ex.Message);
            }
        }

        private async void PrintSelectionButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // sheets to print
                string sheetInput = InputTarief.Text;
                if (InputTarief.Text.Length < 1)
                    throw new Exception($"Please provide Tarieven print range.");

                // get sheet order
                var sheetOrder = ExcelUtility.MultipleRange2List(sheetInput);

                // progress
                var progress = new Progress<int>(value => ProgressBar.Value = value);

                // get catalog type
                var catalogType = ((ComboBoxItem)InputCatalogType.SelectedItem).Content.ToString();

                // btw
                bool inclBtw = InputBTW.IsChecked ?? false;

                // print
                await Task.Run(() => PrintCatalog(sheetOrder, progress, catalogType, inclBtw));
            }
            catch (Exception ex)
            {
                ProgressBar.Value = 0;
                MessageBox.Show(ex.Message);
            }
        }

        private void PrintCatalog(List<string> sheetOrder, IProgress<int> progress, string catalogType, bool inclBtw)
        {
            try
            {                
                // get appconfig parameters
                var parameters = GetConfigParameters();                

                // check if files exists
                if (!File.Exists(parameters.masterCatalog))
                    throw new Exception($"Workbook " + parameters.masterCatalog + " not found!");

                // decrypt password
                string password = HashUtil.Decrypt(parameters.hash);

                // progress update
                progress.Report(5);                

                // export to pdf
                ExcelUtility.ExportWorkbook2Pdf(progress, parameters, password, catalogType,
                    sheetOrder, inclBtw);

                // progress update
                progress.Report(0);
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private static readonly Regex _regex = new Regex("[^0-9;-]+");
        private static bool IsTextInputAllowed(string text)
        {
            return _regex.IsMatch(text);
        }
        private void NumberValidationTextBox(object sender, TextCompositionEventArgs e)
        {
            e.Handled = IsTextInputAllowed(e.Text);
        }
    }
}
