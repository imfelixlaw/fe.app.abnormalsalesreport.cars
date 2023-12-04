using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Cars_Reporting
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

            Version version = assembly.GetName().Version;
            labelVersion.Content = "Version " + version;
            
            myVersion myver = new myVersion();
            labelRev0.Content = myver.getRevision(0);
            labelRev1.Content = myver.getRevision(1);
            labelRev2.Content = myver.getRevision(2);
            labelRev3.Content = myver.getRevision(3);
        }

        private void buttonRptAbnormal_Click(object sender, RoutedEventArgs e)
        {
            using (WindowRptAbnormal wRA = new WindowRptAbnormal())
            {
                wRA.Owner = this;
                wRA.Show();
                wRA.Dispose();
            }
        }

        private void buttonProductByCentre_Click(object sender, RoutedEventArgs e)
        {
            using (WindowRptProductByCentreII WRPC2 = new WindowRptProductByCentreII())
            {
                WRPC2.Owner = this;
                WRPC2.Show();
                WRPC2.Dispose();
            }
        }

        private void buttonProductByCentreI_Click(object sender, RoutedEventArgs e)
        {
            using (WindowRptProductByCentreI WRPC1 = new WindowRptProductByCentreI())
            {
                WRPC1.Owner = this;
                WRPC1.Show();
                WRPC1.Dispose();
            }
        }

        private void buttonCentrePostingDate_Click(object sender, RoutedEventArgs e)
        {
            using (WindowRptCentrePostingDate WRCPD = new WindowRptCentrePostingDate())
            {
                WRCPD.Owner = this;
                WRCPD.Show();
                WRCPD.Dispose();
            }
        }
    }
}
