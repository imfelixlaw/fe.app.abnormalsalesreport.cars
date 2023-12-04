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
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace Cars_Reporting
{
    /// <summary>
    /// Interaction logic for WindowRptAbnormal.xaml
    /// </summary>
    public partial class WindowRptAbnormal : Window, IDisposable
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
        ~WindowRptAbnormal() //--> Change to Object Name
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

        private List<jstorage> alljs = new List<jstorage>(); // all job list with gp coat
        private List<jstorage> suspectedjs = new List<jstorage>(); // suspected job list (contain item other than gp coat)
        private List<jsdata> jsd = new List<jsdata>(); // Jobsheet data
        private List<sccontent> scc = new List<sccontent>();
        private string sGPCoatSC;
        private bool isGenerated = false;

        public WindowRptAbnormal()
        {
            InitializeComponent();
            datePickerEndDate.SelectedDate = DateTime.Now.Date; //--> Set End Date to today
            myConn = new MySqlConnection(m.Setting); // Create MySQL Connection
            sGPCoatSC = Get_GPCoat_ServiceCode(); // get the service code of GPCoat
            StreamWriter sw = new StreamWriter("data.txt");
            sw.Write(sGPCoatSC);
            sw.Close();
            myVersion myver = new myVersion();
            labelRelease.Content = myver.getRevision(0);
        }

        private void buttonGenerateReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(datePickerStartDate.SelectedDate.ToString())
                    || string.IsNullOrEmpty(datePickerEndDate.SelectedDate.ToString()))
                {
                    throw new Exception("Please select valid date range");
                }
                if (datePickerStartDate.SelectedDate > datePickerEndDate.SelectedDate)
                {
                    throw new Exception("Date range is out of range, the start Date cannot greater than end date");
                }
                dataGridResult.DataContext = null;
                dataGridResult.ItemsSource = null;
                dataGridResult.Items.Clear();
                
                DateTime dtStartdate = (DateTime)datePickerStartDate.SelectedDate,
                    dtEnddate = (DateTime)datePickerEndDate.SelectedDate;
                string sStartDate = dtStartdate.ToString("yyyy-MM-dd"),
                    sEndDate = dtEnddate.ToString("yyyy-MM-dd");
                alljs = Get_All_Jobsheet(sStartDate, sEndDate); // Get All Jobsheet Contain GPCoat
                jsd = Get_JobSheetData(alljs); // Finalize the data
                dataGridResult.DataContext = jsd;
                dataGridResult.IsReadOnly = true;
                dataGridResult.ItemsSource = jsd;
                dataGridResult.AutoGenerateColumns = true;
                dataGridResult.Items.Refresh();
                foreach (DataGridColumn col in dataGridResult.Columns)
                {
                    if (col.Header.ToString() == "Site")
                    {
                        dataGridResult.Columns.Remove(col);
                        break;
                    }
                }
                isGenerated = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //-- Data Generate Function -- Start -->
        private List<jsdata> Get_JobSheetData(List<jstorage> tmp)
        {
            List<jsdata> newjsd = new List<jsdata>(); // temporary list to store data
             try
            {
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                foreach (jstorage j in tmp)
                {
                    using (myCmd = new MySqlCommand(string.Format(@"
                        SELECT c.`RegCentre` AS `Centre`,
                            js.`Site` as `Site`,
                            DATE_FORMAT(js.`DateCreated`,'%d-%m-%Y') AS `Date`,
                            js.`JobNo` AS `Receipt No`,
                            js.`CarNo` AS `Car No`,
                            js.`TelHP` AS `Tel No`,
                            js.`TotalCharge` AS `Total`
                        FROM `tbljobsheet` AS js
                        INNER JOIN `tblregcentre` AS c ON c.`Site` = js.`Site`
                        WHERE js.`JobNo` = {0}
                        AND js.`Site` = {1}
                        GROUP BY `Centre`, `Date`, `Receipt No`, `Car No`, `Tel No`;", j.jobno, j.site), myConn))
                    {
                        myCmd.CommandTimeout = 0;
                        using (myDr = myCmd.ExecuteReader())
                        {
                            while (myDr.Read())
                            {
                                newjsd.Add(new jsdata {
                                    Centre = myDr.GetValue(0).ToString(),
                                    Site = myDr.GetValue(1).ToString(),
                                    Date = myDr.GetValue(2).ToString(),
                                    ReceiptNo = myDr.GetValue(3).ToString(),
                                    CarNo = myDr.GetValue(4).ToString(),
                                    TelNo = myDr.GetValue(5).ToString(),
                                    Total = myDr.GetValue(6).ToString()});
                            }
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return newjsd;
        }

        // All Jobsheet got GP Coat and other service
        private List<jstorage> Get_All_Jobsheet(string sFrom, string sTo)
        {
            List<jstorage> tmplist = new List<jstorage>();
            try
            {
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                using (myCmd = new MySqlCommand(string.Format(@"
SELECT jss.`Site`, jss.`JobNo`,
	(SELECT COUNT(jss2.`JobNo`) FROM `tbljobsheetservice` AS jss2
	WHERE jss2.`JobNo` = jss.`JobNo` AND jss2.`Site` = jss.`Site`) AS `Qty`,
	(SELECT COUNT(jss3.`JobNo`) FROM `tbljobsheetservice` AS jss3
	WHERE jss3.`JobNo` = jss.`JobNo` AND jss3.`Site` = jss.`Site` AND jss3.`ServiceCode` IN ({2})) AS `GPQty`	
FROM `tbljobsheetservice` AS jss
WHERE jss.`DateCreated` BETWEEN '{0} 00:00:00' AND  '{1} 23:59:59' AND jss.`ServiceCode` IN ({2})
GROUP BY jss.`JobNo`, jss.`Site`
HAVING `Qty` <> `GPQty`
ORDER BY jss.`Site`, jss.`JobNo` ASC", sFrom, sTo, sGPCoatSC), myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            tmplist.Add(new jstorage {site = myDr.GetValue(0).ToString(),
                                jobno = myDr.GetValue(1).ToString()});
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return tmplist;
        }

        private string Get_GPCoat_ServiceCode()
        {
            string sServiceCode = string.Empty, ServiceCodeTpl = "'{0}'";
            try
            {
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                // Reading Voucher Table for GPCoat
                using (myCmd = new MySqlCommand(@"SELECT `ServiceCode` FROM `tblvoucher` WHERE `ServiceType` LIKE '%GP%COAT%';", myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            if (!string.IsNullOrEmpty(sServiceCode)) { sServiceCode += ", "; }
                            sServiceCode += string.Format(ServiceCodeTpl, myDr.GetValue(0).ToString());
                        }
                    }
                }

                // Reading Service Code Table for GPCoat
                using (myCmd = new MySqlCommand(@"SELECT `ServiceCode` FROM `tblservicecode` WHERE `ServiceName` LIKE  '%GP%COAT%';", myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            if (!string.IsNullOrEmpty(sServiceCode)) { sServiceCode += ", "; }
                            sServiceCode += string.Format(ServiceCodeTpl, myDr.GetValue(0).ToString());
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return sServiceCode;
        }

        private void buttonViewData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!isGenerated)
                {
                    throw new Exception("You need to generate the data first");
                }
                jsdata selectedData = dataGridResult.SelectedItem as jsdata;
                if (selectedData != null)
                {
                    using (WindowRptAbnormalDetails wRAD =
                        new WindowRptAbnormalDetails(selectedData.Centre, selectedData.Site, selectedData.Date, selectedData.ReceiptNo, selectedData.CarNo, selectedData.TelNo, selectedData.Total))
                    {
                        wRAD.Owner = this;
                        wRAD.ShowDialog();
                        wRAD.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        //-- Data Generate Function -- End -->
        private void buttonClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void buttonExcel_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!isGenerated)
                {
                    throw new Exception("You need to generate the data first");
                }
                // save the application
                if (dataGridResult.Items.Count > 0)
                {
                    string fileName = String.Empty;
                    System.Windows.Forms.SaveFileDialog saveFileExcel = new System.Windows.Forms.SaveFileDialog();
                    saveFileExcel.Filter = "Excel files |*.xls|All files (*.*)|*.*";
                    saveFileExcel.FilterIndex = 1;
                    saveFileExcel.RestoreDirectory = true;
                    if (saveFileExcel.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        fileName = saveFileExcel.FileName;
                    }
                    else
                    {
                        throw new Exception("Must provide a file name to save");
                    }
                    // creating Excel Application
                    Excel._Application app = new Excel.Application();
                    // creating new WorkBook within Excel application
                    Excel._Workbook workbook = app.Workbooks.Add(Type.Missing);
                    // creating new Excelsheet in workbook
                    Excel._Worksheet worksheet = null;
                    // see the excel sheet behind the program
                    //Funny
                    app.Visible = false;
                    // get the reference of first sheet. By default its name is Sheet1.
                    // store its reference to worksheet
                    //Fixed:(Excel.Worksheet)
                    worksheet = (Excel.Worksheet)workbook.Sheets["Sheet1"];
                    worksheet = (Excel.Worksheet)workbook.ActiveSheet;
                    // changing the name of active sheet
                    worksheet.Name = "Abnormal Report";
                    // storing header part in Excel

                    int iRow = 1;
                    worksheet.Cells[iRow, 1].Value = "Abnormal Report (GPCoat)";
                    worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                    iRow += 2;
                    worksheet.Cells[iRow, 1].Value = "Centre";
                    worksheet.Cells[iRow, 2].Value = "ReceiptNo";
                    worksheet.Cells[iRow, 3].Value = "Date";
                    worksheet.Cells[iRow, 4].Value = "CarNo";
                    worksheet.Cells[iRow, 5].Value = "TelNo";
                    worksheet.Cells[iRow, 6].Value = "Total";
                    worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                    worksheet.Cells[iRow, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[iRow, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[iRow, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[iRow, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[iRow, 5].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    worksheet.Cells[iRow, 6].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    iRow++;

                    foreach (jsdata j in jsd)
                    {
                        worksheet.Cells[iRow, 1].EntireRow.NumberFormat = "@"; // Make Header Row Reading Like Text
                        worksheet.Cells[iRow, 1].Value = j.Centre;
                        worksheet.Cells[iRow, 2].Value = j.ReceiptNo;
                        worksheet.Cells[iRow, 3].Value = j.Date;
                        worksheet.Cells[iRow, 4].Value = j.CarNo;
                        worksheet.Cells[iRow, 5].Value = j.TelNo;
                        worksheet.Cells[iRow, 6].Value = j.Total;
                        if (checkBoxIncludeItems.IsChecked == true)
                        {
                            ServiceCode();
                            List<jscontent> jsc = new List<jscontent>();
                            List<jscontent> jsc2 = new List<jscontent>();
                            if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                            using (myCmd = new MySqlCommand(string.Format(@"
                                SELECT `ServiceCode`, `Charge`
                                FROM `tbljobsheetservice`
                                WHERE `JobNo` = {0}
                                AND `Site` = {1};", j.ReceiptNo, j.Site), myConn))
                            {
                                myCmd.CommandTimeout = 0;
                                using (myDr = myCmd.ExecuteReader())
                                {
                                    while (myDr.Read())
                                    {
                                        iRow++;
                                        worksheet.Cells[iRow, 1].EntireRow.NumberFormat = "@"; // Make Header Row Reading Like Text
                                        worksheet.Cells[iRow, 2].Value = "Itemize";
                                        worksheet.Cells[iRow, 3].Value = AssignServiceCode(myDr.GetValue(0).ToString());
                                        worksheet.Cells[iRow, 6].Value = myDr.GetValue(1).ToString();
                                    }
                                }
                            }
                            if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
                        }
                        iRow++;
                    }

                    //Fixed-old code :11 para->add 1:Type.Missing
                    workbook.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    // Exit from the application
                    app.Quit();
                    MessageBox.Show("Excel file is successfully created");
                }
                else
                {
                    throw new Exception("No Data Founded");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void ServiceCode()
        {
            try
            {
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                // Reading Service Code Table for GPCoat
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
                // Reading Voucher Table for GPCoat
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
        
        private string AssignServiceCode(string sc)
        {
            foreach (sccontent s in scc)
            {
                if (s.scode == sc)
                    return s.sname;
            }
            return string.Empty;
        }
    }
}