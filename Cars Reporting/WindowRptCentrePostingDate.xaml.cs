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

namespace Cars_Reporting
{
    /// <summary>
    /// Interaction logic for WindowRptCentrePostingDate.xaml
    /// </summary>
    public partial class WindowRptCentrePostingDate : Window, IDisposable
    {
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
        ~WindowRptCentrePostingDate() //--> Change to Object Name
        {
            /* Do not re-create Dispose clean-up code here.
               Calling Dispose(false) is optimal in terms of readability and maintainability. */
            Dispose(false);
        }
        //-- IDisposable -- End -->
		private myConn m = new myConn();
        private MySqlConnection myConn; // Mysql Connection
        private MySqlCommand myCmd; // MySql Command
        private MySqlDataReader myDr; // MySql Data Reader
        
        public WindowRptCentrePostingDate()
        {
            InitializeComponent();
			myConn = new MySqlConnection(m.Setting); // Create MySQL Connection
            GenerateMonthYearComboBox();
        }

        private void GenerateMonthYearComboBox()
        {
            DateTime Today = DateTime.Now; // Get Today
            int[] YearArray = new int[5], // Year
                  MonthArray = new int[12] { 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12 }; // Month
            for (int i = 0; i < 5; i++) { YearArray[i] = (Today.Year) - i; } // Getting Year from -3 toyear +1
            comboBoxYear.DataContext = YearArray; // Storing Year
            comboBoxYear.SelectedIndex = 0; // Select ToYear
            comboBoxMonth.DataContext = MonthArray; // Storing Month
            comboBoxMonth.SelectedIndex = (Today.Month - 1); // Select ToMonth
        }
    }
}
