//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Configuration;
using System.Windows;
using System.Windows.Media;
using FirstFloor.ModernUI.Presentation;
using MTMIntegration;
using MTMLiveReporting.Content;

namespace MTMLiveReporting
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow 
    {
        public MainWindow()
        {
            InitializeComponent();
            DataGetter.FirstRun = bool.Parse(ConfigurationManager.AppSettings["FirstRun"]);

            var concol = new Color
            {
                R = byte.Parse(ConfigurationManager.AppSettings["R"]),
                A = byte.Parse(ConfigurationManager.AppSettings["A"]),
                G = byte.Parse(ConfigurationManager.AppSettings["G"]),
                B = byte.Parse(ConfigurationManager.AppSettings["B"])
            };
            AppearanceManager.Current.AccentColor = concol;
            var fontsize = ConfigurationManager.AppSettings["FontSize"];
            AppearanceManager.Current.FontSize = fontsize.Equals("large", StringComparison.OrdinalIgnoreCase) ? FirstFloor.ModernUI.Presentation.FontSize.Large : FirstFloor.ModernUI.Presentation.FontSize.Small;
            var theme = ConfigurationManager.AppSettings["Theme"];
            AppearanceManager.Current.ThemeSource = theme.Equals("dark", StringComparison.OrdinalIgnoreCase) ? AppearanceManager.DarkThemeSource : AppearanceManager.LightThemeSource;
            DataContext = new SettingsAppearanceViewModel();
            if (!DataGetter.FirstRun)

            {
                try
                {

                    MtmInteraction.InitializeVstfConnection(new Uri(ConfigurationManager.AppSettings["TFSUrl"]),
                        ConfigurationManager.AppSettings["TeamProject"],
                        int.Parse(ConfigurationManager.AppSettings["TestPlanID"]));
                    DataGetter.Diagnostic.AppendLine("TFS URL: " + ConfigurationManager.AppSettings["TFSUrl"]);
                    DataGetter.Diagnostic.AppendLine("Team Project: " + ConfigurationManager.AppSettings["TeamProject"]);
                    DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                   
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