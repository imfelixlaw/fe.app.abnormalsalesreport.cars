using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Windows;

namespace Cars_Reporting
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
    }

    public class myVersion
    {
        private string [] version = new string[10];
        public myVersion()
        {
            version[0] = "1.1"; //WindowRptAbnormal
            version[1] = "1.0"; //WindowRptProductByCentreI
            version[2] = "1.3"; //WindowRptProductByCentreII
            version[3] = "1.0"; //WindowRptCentrePostingDate
        }

        public string getRevision(int ID)
        {
            return "Revision " + version[ID];
        }
    }

    public class myConn
    {
        public string Setting
        {
            get
            {
                return string.Format(@"SERVER={0};
                    DATABASE={1};
                    UID={2};
                    PASSWORD={3};
                    respect binary flags=false; Compress=true; Pooling=true; Min Pool Size=0; Max Pool Size=100; Connection Lifetime=0",
                        Properties.Settings.Default.myhost,
                        Properties.Settings.Default.mytable,
                        Properties.Settings.Default.myuser,
                        Properties.Settings.Default.mypass);
            }
        }
    }
}
