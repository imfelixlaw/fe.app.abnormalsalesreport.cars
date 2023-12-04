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
using System.Windows.Shapes;
//-- IDisposable -- Start -->
using System.ComponentModel;
//-- IDisposable -- End -->
using MySql.Data.MySqlClient;
using System.Data;

namespace Cars_Reporting
{
    /// <summary>
    /// Interaction logic for WindowRptAbnormalDetails.xaml
    /// </summary>
    public partial class WindowRptAbnormalDetails : Window, IDisposable
    {
        private myConn m = new myConn();
        private MySqlConnection myConn; // Mysql Connection
        private MySqlCommand myCmd; // MySql Command
        private MySqlDataReader myDr; // MySql Data Reader

        //-- IDisposable -- Start -->
        //--> All other IDisposable code on other source code will skip the comment, please refer to here
        private IntPtr handle; //--> Pointer to an external unmanaged resource.
        private Component component = new Component(); //--> Other managed resource this class uses.
        private bool disposed = false; //--> Track whether Dispose has been called.

        public void Dispose()
        {
            Dispose(true);
            /* This object will be cleaned up by the Dispose method.
               Therefore, you should call GC.SupressFinalize to take this object off the finalization queue
               and prevent finalization code for this object from executing a second time. */
            GC.SuppressFinalize(this);
        }

        /* Dispose(bool disposing) executes in two distinct scenarios.
           If disposing equals true, the method has been called directly or indirectly by a user's code.
           Managed and unmanaged resources can be disposed.
           If disposing equals false, the method has been called by the runtime from inside the finalizer
           and you should not reference other objects. Only unmanaged resources can be disposed. */
        protected virtual void Dispose(bool disposing)
        {
            //--> Check to see if Dispose has already been called.
            if (!this.disposed)
            {
                //--> If disposing equals true, dispose all managed and unmanaged resources.
                if (disposing)
                {
                    component.Dispose(); // Dispose managed resources.
                }
                /* Call the appropriate methods to clean up unmanaged resources here.
                   If disposing is false, only the following code is executed. */
                CloseHandle(handle);
                handle = IntPtr.Zero;
                disposed = true; //--> Note disposing has been done.
            }
        }

        //--> Use interop to call the method necessary to clean up the unmanaged resource.
        [System.Runtime.InteropServices.DllImport("Kernel32")]
        private extern static Boolean CloseHandle(IntPtr handle);

        /* Use C# destructor syntax for finalization code.
           This destructor will run only if the Dispose method does not get called.
           It gives your base class the opportunity to finalize.
           Do not provide destructors in types derived from this class. */
        ~WindowRptAbnormalDetails() //--> Change to Object Name
        {
            /* Do not re-create Dispose clean-up code here.
               Calling Dispose(false) is optimal in terms of readability and maintainability. */
            Dispose(false);
        }
        //-- IDisposable -- End -->

        List<sccontent> scc = new List<sccontent>();

        public WindowRptAbnormalDetails(string centre, string site, string date, string jobno, string carno, string telno, string total)
        {
            InitializeComponent();
            myConn = new MySqlConnection(m.Setting); // Create MySQL Connection
            textBoxCentre.Text = centre;
            textBoxDate.Text = date;
            textBoxReceiptNo.Text = jobno;
            textBoxCarNo.Text = carno;
            textBoxTelNo.Text = telno;
            textBoxTotal.Text = total;
            ServiceCode();
            FetchContent(jobno, site);
        }

        private string AssignServiceCode(string sc)
        {
            foreach (sccontent s in scc)
            {
                if (s.scode == sc)
                    return s.sname;
            }
            return string.Empty;
        }

        private void ServiceCode()
        {
            try
            {
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                using (myCmd = new MySqlCommand(@"SELECT `ServiceCode`, `ServiceName` FROM `tblservicecode`;", myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            scc.Add(new sccontent
                            {
                                scode = myDr.GetValue(0).ToString(),
                                sname = myDr.GetValue(1).ToString()
                            });
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                using (myCmd = new MySqlCommand(@"SELECT `ServiceCode`, `ServiceType` FROM `tblvoucher`;", myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            scc.Add(new sccontent
                            {
                                scode = myDr.GetValue(0).ToString(),
                                sname = myDr.GetValue(1).ToString()
                            });
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void FetchContent(string jobno, string site)
        {
            try
            {
                List<jscontent> jsc = new List<jscontent>();
                List<jscontent> jsc2 = new List<jscontent>();
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                using (myCmd = new MySqlCommand(string.Format(@"
                        SELECT `ServiceCode`, `Charge`
                        FROM `tbljobsheetservice`
                        WHERE `JobNo` = {0}
                        AND `Site` = {1};", jobno, site), myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            jsc.Add(new jscontent {
                                Service = AssignServiceCode(myDr.GetValue(0).ToString()),
                                Amount = myDr.GetValue(1).ToString()});
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
                dataGridContents.ItemsSource = jsc;
                dataGridContents.DataContext = jsc;
                dataGridContents.AutoGenerateColumns = true;
                dataGridContents.Items.Refresh();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
