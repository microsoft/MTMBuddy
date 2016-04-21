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
    public partial class SettingsVSTF : UserControl
    {
        public SettingsVSTF()
        {
            InitializeComponent();
            txtTfsUrl.Text = ConfigurationManager.AppSettings["TFSUrl"];
            txtTeamProject.Text = ConfigurationManager.AppSettings["TeamProject"];
        
            txtTestPlanID.Text = ConfigurationManager.AppSettings["TestPlanID"];
     
        }

        private void saveButton_Click(object sender, RoutedEventArgs e)
        {
           
            
            

            ConfigurationManager.AppSettings["TFSUrl"] = txtTfsUrl.Text;
            ConfigurationManager.AppSettings["TeamProject"] = txtTeamProject.Text;
        
            ConfigurationManager.AppSettings["TestPlanId"] = txtTestPlanID.Text;
       
            try
            {
                MTMInteraction.initializeVSTFUpdate(new Uri(ConfigurationManager.AppSettings["TFSUrl"]),
                    ConfigurationManager.AppSettings["TeamProject"],
                    int.Parse(ConfigurationManager.AppSettings["TestPlanID"]),
                    ConfigurationManager.AppSettings["BuildNumber"]);
                DataGetter.SaveConfig("TFSUrl", txtTfsUrl.Text);
                DataGetter.SaveConfig("TeamProject", txtTeamProject.Text);
                DataGetter.SaveConfig("TestPlanID", txtTestPlanID.Text);

                ConfigurationManager.AppSettings["FirstRun"] = false.ToString();
                DataGetter.SaveConfig("FirstRun", false.ToString());



                DataGetter.diagnostic.AppendLine("Config changed:");
                DataGetter.diagnostic.AppendLine("TFS URL: " + txtTfsUrl.Text);
                DataGetter.diagnostic.AppendLine("Team Project: " + txtTeamProject.Text);
               

                DataGetter.diagnostic.AppendLine("---------------------------------------------------");
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