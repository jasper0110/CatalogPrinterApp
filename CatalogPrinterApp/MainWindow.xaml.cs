using Encrypter;
using ExcelUtil;
using System;
using System.Collections.Generic;
using System.Configuration;
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

namespace CatalogPrinterApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly string _configPath = @"C:\ProgramData\CatalogPrinterApp\CatalogPrinterApp.config";

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

        private List<string> GetSheetOrder()
        {
            var sheets2Print = new List<string>();

            var sheets = InputTarief.Text.Split(';').ToList();
            foreach (var sheet in sheets)
            {
                if (sheet.Contains("-"))
                {
                    var first = sheet.Substring(0, sheet.IndexOf("-"));
                    var second = sheet.Substring(sheet.IndexOf("-") + 1);

                    var firstInt = 0;
                    var secondInt = 0;
                    if (!Int32.TryParse(first, out firstInt))
                        throw new Exception($"Parse exception for sheet input " + first + " from range " + sheet + "! Please check sheet input.");
                    if (!Int32.TryParse(second, out secondInt))
                        throw new Exception($"Parse exception for sheet input " + second + " from range " + sheet + "! Please check sheet input.");

                    for (int i = firstInt; i <= secondInt; ++i)
                        sheets2Print.Add(i.ToString());
                }
                else
                {
                    if(sheet.Length > 0)
                        sheets2Print.Add(sheet);
                }
            }
            return sheets2Print;
        }

        private async void PrintButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // open config
                if (!File.Exists(_configPath))
                    throw new Exception($"Config file " + _configPath + " not found!");
                ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
                configMap.ExeConfigFilename = _configPath;
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
                var appSettings = config.GetSection("appSettings") as AppSettingsSection;
                // get config values
                string hash = appSettings.Settings["key"]?.Value;
                string masterCatalog = appSettings.Settings["masterCatalog"]?.Value;
                string outputPath = appSettings.Settings["outputPath"]?.Value;

                var ranges = new InputRanges();
                ranges.catalogType = appSettings.Settings["cellCatalogType"]?.Value;
                ranges.btw = appSettings.Settings["cellBtw"]?.Value;
                ranges.footerRight = appSettings.Settings["cellFooterRight"]?.Value;
                ranges.footerLeft = appSettings.Settings["cellFooterLeft"]?.Value;
                ranges.footerMid = appSettings.Settings["cellFooterMid"]?.Value;
                ranges.headerRight = appSettings.Settings["cellHeaderRight"]?.Value;
                ranges.headerLeft = appSettings.Settings["cellHeaderLeft"]?.Value;
                ranges.headerMid = appSettings.Settings["cellHeaderMid"]?.Value;
                ranges.printArea = appSettings.Settings["printArea"]?.Value;                

                if (hash == null)
                    throw new Exception($"Could not find 'password' key in " + _configPath + "!");
                if (masterCatalog == null)
                    throw new Exception($"Could not find 'masterCatalog' key in " + _configPath + "!");
                if (outputPath == null)
                    throw new Exception($"Could not find 'outputPath' key in " + _configPath + "!");

                if (ranges.catalogType == null)
                    throw new Exception($"Could not find 'cellCatalogType' key in " + _configPath + "!");
                if (ranges.btw == null)
                    throw new Exception($"Could not find 'cellBtw' key in " + _configPath + "!");
                if (ranges.footerRight == null)
                    throw new Exception($"Could not find 'cellFooterRight' key in " + _configPath + "!");
                if (ranges.footerLeft == null)
                    throw new Exception($"Could not find 'cellFooterLeft' key in " + _configPath + "!");
                if (ranges.footerMid == null)
                    throw new Exception($"Could not find 'cellFooterMid' key in " + _configPath + "!");
                if (ranges.headerRight == null)
                    throw new Exception($"Could not find 'cellHeaderRight' key in " + _configPath + "!");
                if (ranges.headerLeft == null)
                    throw new Exception($"Could not find 'cellHeaderLeft' key in " + _configPath + "!");
                if (ranges.headerMid == null)
                    throw new Exception($"Could not find 'cellHeaderMid' key in " + _configPath + "!");
                if (ranges.printArea == null)
                    throw new Exception($"Could not find 'printArea' key in " + _configPath + "!");


                // check if files exists
                if (!File.Exists(masterCatalog))
                    throw new Exception($"Workbook " + masterCatalog + " not found!");

                // decrypt password
                string password = HashUtil.Decrypt(hash);

                // get catalog type
                string catalogType = InputCatalogType.SelectedItem.ToString();

                // sheets to print
                var sheets2Print = GetSheetOrder();

                await Task.Run(() => ExcelUtility.ExportWorkbook2Pdf(masterCatalog, password, catalogType, outputPath, sheets2Print, ranges));
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
    }
}
