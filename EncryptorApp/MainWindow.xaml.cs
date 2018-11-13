using System;
using System.Collections.Generic;
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
using System.Configuration;
using Encrypter;
using ExcelUtil;

namespace EncryptorApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
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

        private void ChangePassword_OnClick(object sender, RoutedEventArgs e)
        {
            string firstPassword = NewPasswordTextBox.Password;
            string secondPassword = ConfirmNewPasswordTextBox.Password;
            if (firstPassword != secondPassword)
            {
                MessageBox.Show("Passwords do not match!");
                return;
            }

            try
            {
                // encrypt new password
                string encryptedPassword = HashUtil.Encrypt(firstPassword);

                // open config
                string configPath = @"C:\ProgramData\CatalogPrinterApp\CatalogPrinterApp.config";
                if (!File.Exists(configPath))
                    throw new Exception($"Config file " + configPath + " not found!");
                ExeConfigurationFileMap configMap = new ExeConfigurationFileMap();
                configMap.ExeConfigFilename = configPath;
                Configuration config = ConfigurationManager.OpenMappedExeConfiguration(configMap, ConfigurationUserLevel.None);
                var appSettings = config.GetSection("appSettings") as AppSettingsSection;
                // get config values
                string oldHash = appSettings.Settings["key"].Value;
                string masterCatalog = appSettings.Settings["masterCatalog"].Value;

                // try opening catalog with old password and change the password
                if (!File.Exists(masterCatalog))
                    throw new Exception($"Workbook " + masterCatalog + " not found!");
                string oldPassword = HashUtil.Decrypt(oldHash);
                ExcelUtility.ChangePassword(masterCatalog, oldPassword, firstPassword);

                // write new encrypted password to config
                appSettings.Settings["key"].Value = encryptedPassword;
                config.Save(ConfigurationSaveMode.Modified);

                MessageBox.Show("Password changed!");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                MessageBox.Show("Error! Could not change the password!");
            }

        }
    }
}
