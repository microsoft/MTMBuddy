using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using FirstFloor.ModernUI.Windows.Navigation;
using MTMIntegration;

namespace MTMLiveReporting
{
    /// <summary>
    ///     Interaction logic for Splash.xaml
    /// </summary>
    public partial class Splash : UserControl
    {
        public Splash()
        {
            InitializeComponent();
            //if (DataGetter.FirstRun)
            //{
            //    var frame = NavigationHelper.FindFrame(null, this);
            //    frame.Source = new Uri("Pages/Settings.xaml", UriKind.Relative);
            //}
        }

        private void btnProceed_OnClick(object sender, RoutedEventArgs e)
        {
            var frame = NavigationHelper.FindFrame(null, this);
            if (frame != null)
            {
                if (DataGetter.FirstRun)
                {
                    frame.Source = new Uri("Pages/Settings.xaml", UriKind.Relative);
                }
                else
                {
                    frame.Source = new Uri("Pages/Home.xaml", UriKind.Relative);
                }
            }
        }

        private void btnProceed_OnMouseEnter(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Hand;
            btnProceed.Opacity = 1.0;
            //ImageBrush brush = new ImageBrush();
            //brush.ImageSource = new BitmapImage(new Uri("proceed1.png", UriKind.Relative));
            //btnProceed.Background = brush;
        }

        private void btnProceed_OnMouseLeave(object sender, MouseEventArgs e)
        {
            Mouse.OverrideCursor = Cursors.Arrow;
            btnProceed.Opacity = 0.7;
        }
    }
}