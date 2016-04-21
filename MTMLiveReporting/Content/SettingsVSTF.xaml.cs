//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System;
using System.Configuration;
using System.Linq;
using System.Security.Principal;
using System.Windows;
using System.Windows.Forms;
using MTMIntegration;

using MessageBox = System.Windows.MessageBox;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using UserControl = System.Windows.Controls.UserControl;

namespace MTMLiveReporting
{
    /// <summary>
    ///     Interaction logic for SettingsVSTF.xaml
    /// </summary>
    [System.Diagnostics.CodeAnalysis.SuppressMessage("Microsoft.Naming", "CA1709:IdentifiersShouldBeCasedCorrectly", MessageId = "VSTF")]
    public partial class SettingsVstf : UserControl
    {
        public SettingsVstf()
        {
            InitializeComponent();
            TxtTfsUrl.Text = ConfigurationManager.AppSettings["TFSUrl"];
            TxtTeamProject.Text = ConfigurationManager.AppSettings["TeamProject"];
        
            TxtTestPlanId.Text = ConfigurationManager.AppSettings["TestPlanID"];
     
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
           
            
            

            ConfigurationManager.AppSettings["TFSUrl"] = TxtTfsUrl.Text;
            ConfigurationManager.AppSettings["TeamProject"] = TxtTeamProject.Text;
        
            ConfigurationManager.AppSettings["TestPlanId"] = TxtTestPlanId.Text;
       
            try
            {
                MtmInteraction.Initialize_VSTF(new Uri(ConfigurationManager.AppSettings["TFSUrl"]),
                    ConfigurationManager.AppSettings["TeamProject"],
                    int.Parse(ConfigurationManager.AppSettings["TestPlanID"]));
                DataGetter.SaveConfig("TFSUrl", TxtTfsUrl.Text);
                DataGetter.SaveConfig("TeamProject", TxtTeamProject.Text);
                DataGetter.SaveConfig("TestPlanID", TxtTestPlanId.Text);

                ConfigurationManager.AppSettings["FirstRun"] = false.ToString();
                DataGetter.SaveConfig("FirstRun", false.ToString());



                DataGetter.Diagnostic.AppendLine("Config changed:");
                DataGetter.Diagnostic.AppendLine("TFS URL: " + TxtTfsUrl.Text);
                DataGetter.Diagnostic.AppendLine("Team Project: " + TxtTeamProject.Text);
               

                DataGetter.Diagnostic.AppendLine("---------------------------------------------------");
                MessageBox.Show("Voila! All changes have been saved.");
            }


            catch (Exception exp)
            {
                MessageBox.Show(
                    "Unable to connect to VSTF.Please check your connectivity and configuration" + Environment.NewLine +
                    "Exception Details: " + exp.Message, "OOPS!");
            }
        }

       
    }
}