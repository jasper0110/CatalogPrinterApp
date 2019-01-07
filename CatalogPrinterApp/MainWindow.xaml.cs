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
            string pathEncryptorApp = GetConfigValue("pathEncryptorApp", appSettings);

            // start new process
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
                MessageBox.Show(ex.Message);
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

        private string GetConfigValue(string value, AppSettingsSection appSettings)
        {
            var result = appSettings.Settings[value]?.Value;
            if (result == null)
                throw new Exception($"Could not find '{value}' key in " + _configPath + "!");
            return result;
        }

        private AppParameters GetConfigParameters()
        {
            var parameters = new AppParameters();

            // open config
            if (!File.Exists(_configPath))
                throw new Exception($"Config file " + _configPath + " not found!");
            ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
            configMap.ExeConfigFilename = _configPath;
            Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
            var appSettings = config.GetSection("appSettings") as AppSettingsSection;
            // get config values
            parameters.hash = GetConfigValue("key", appSettings);
            parameters.masterCatalog = GetConfigValue("masterCatalog", appSettings);
            parameters.outputPath = GetConfigValue("outputPath", appSettings);

            var ranges = new InputRanges();
            ranges.catalogType = GetConfigValue("cellCatalogType", appSettings);
            ranges.btw = GetConfigValue("cellBtw", appSettings);
            ranges.footerRight = GetConfigValue("cellFooterRight", appSettings);
            ranges.footerLeft = GetConfigValue("cellFooterLeft", appSettings);
            ranges.footerMidFirst = GetConfigValue("cellFooterMidFirst", appSettings);
            ranges.footerMidSecond = GetConfigValue("cellFooterMidSecond", appSettings);
            ranges.headerRight = GetConfigValue("cellHeaderRight", appSettings);
            ranges.headerLeft = GetConfigValue("cellHeaderLeft", appSettings);
            ranges.headerMid = GetConfigValue("cellHeaderMid", appSettings);
            ranges.printArea = GetConfigValue("printArea", appSettings);
            parameters.ranges = ranges;

            parameters.sheetSummaryName = GetConfigValue("sheetSummaryName", appSettings);           

            return parameters;
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
                var sheetOrder = ConverterUtility.MultipleRange2List(sheetInput);

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
