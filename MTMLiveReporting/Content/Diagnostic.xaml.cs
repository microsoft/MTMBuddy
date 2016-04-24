//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 
using System.Windows;
using System.Windows.Controls;

namespace MTMLiveReporting.Content
{
    /// <summary>
    ///     Interaction logic for Diagnostic.xaml
    /// </summary>
    public partial class Diagnostic 
    {
        public Diagnostic()
        {
            InitializeComponent();
            
            TxtDiagnostic.Text = DataGetter.Diagnostic.ToString();
        }

        private void UserControl_GotFocus(object sender, RoutedEventArgs e)
        {
            TxtDiagnostic.Text = DataGetter.Diagnostic.ToString();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            TxtDiagnostic.Text = DataGetter.Diagnostic.ToString();
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            TxtDiagnostic.Text = "";
            DataGetter.Diagnostic.Clear();
        }
    }
}