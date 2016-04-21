//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Configuration;
using System.Windows;
using System.Windows.Media;
using FirstFloor.ModernUI.Presentation;
using FirstFloor.ModernUI.Windows.Controls;
using MTMIntegration;
using MTMLiveReporting.Content;

namespace MTMLiveReporting
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : ModernWindow
    {
        public MainWindow()
        {
            InitializeComponent();
            DataGetter.FirstRun = bool.Parse(ConfigurationManager.AppSettings["FirstRun"]);
         
            var concol = new Color();
            concol.R = byte.Parse(ConfigurationManager.AppSettings["R"]);
            concol.A = byte.Parse(ConfigurationManager.AppSettings["A"]);
            concol.G = byte.Parse(ConfigurationManager.AppSettings["G"]);
            concol.B = byte.Parse(ConfigurationManager.AppSettings["B"]);
            AppearanceManager.Current.AccentColor = concol;
            var fontsize = ConfigurationManager.AppSettings["FontSize"];
            if (fontsize.Equals("large", StringComparison.InvariantCultureIgnoreCase))
            {
                AppearanceManager.Current.FontSize = FirstFloor.ModernUI.Presentation.FontSize.Large;
            }
            else
            {
                AppearanceManager.Current.FontSize = FirstFloor.ModernUI.Presentation.FontSize.Small;
            }
            var Theme = ConfigurationManager.AppSettings["Theme"];
            if (Theme.Equals("dark", StringComparison.InvariantCultureIgnoreCase))
            {
                AppearanceManager.Current.ThemeSource = AppearanceManager.DarkThemeSource;
            }
            else
            {
                AppearanceManager.Current.ThemeSource = AppearanceManager.LightThemeSource;
            }
            DataContext = new SettingsAppearanceViewModel();
            if (!DataGetter.FirstRun)

            {
                try
                {

                    MTMInteraction.initializeVSTFUpdate(new Uri(ConfigurationManager.AppSettings["TFSUrl"]),
                        ConfigurationManager.AppSettings["TeamProject"],
                        int.Parse(ConfigurationManager.AppSettings["TestPlanID"]),
                        ConfigurationManager.AppSettings["BuildNumber"]);
                    DataGetter.diagnostic.AppendLine("TFS URL: " + ConfigurationManager.AppSettings["TFSUrl"]);
                    DataGetter.diagnostic.AppendLine("Team Project: " + ConfigurationManager.AppSettings["TeamProject"]);
                    DataGetter.diagnostic.AppendLine("---------------------------------------------------");
                   
                }
                catch (Exception exp)
                {
                    MessageBox.Show(
                         "Unable to connect to VSTF.Please check your connectivity and configuration." +
                        Environment.NewLine + "Exception Details: " + exp.Message, "OOPS!");
                    
                }
            }
        }
    }
}