using System.Deployment.Application;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Controls;

namespace MTMLiveReporting.Content
{
    /// <summary>
    ///     Interaction logic for About.xaml
    /// </summary>
    public partial class About : UserControl
    {
        public About()
        {
            InitializeComponent();
          
            
            var version = AssemblyVersion();

            versionBlock.Text = version;
        }

        public static string AssemblyVersion()
        {
            if (ApplicationDeployment.IsNetworkDeployed)
            {
                 
                var ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
                return string.Format(System.Globalization.CultureInfo.CurrentCulture,"{0}.{1}.{2}.{3}", ver.Major, ver.Minor, ver.Build, ver.Revision);
            }

            return null;
        }
    }
}