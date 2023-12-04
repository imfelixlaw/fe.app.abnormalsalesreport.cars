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
    /// Interaction logic for WindowRptProductByCentre.xaml
    /// </summary>
    public partial class WindowRptProductByCentreI : Window, IDisposable
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
        ~WindowRptProductByCentreI() //--> Change to Object Name
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
        //private MySqlDataAdapter myDa;
        // Site List
        private List<ssite> lsiteDef = new List<ssite>(); // List Default Value from database
        private List<ssite> lsiteAll = new List<ssite>(); // List all not selected
        private List<ssite> lsiteSel = new List<ssite>(); // List selected

        // Product List (service code) / (voucher code)
        private List<sproduct> lprodDef = new List<sproduct>(); // List Default Value from database
        private List<sproduct> lprodAll = new List<sproduct>(); // List all not selected
        private List<sproduct> lprodSel = new List<sproduct>(); // List selected
        
        // Generator, for sql order by listing
        private const int OrderByCentre = 0;
        private const int OrderByProduct = 1;

        private int iTotalCentre = 0;
        private int iTotalProduct = 0;


        public WindowRptProductByCentreI()
        {
            InitializeComponent();
            myConn = new MySqlConnection(m.Setting); // Create MySQL Connection
            LoadCentreList();
            LoadProductList();
            DisplayNoResult();
            ValidataButtonAll();
            myVersion myver = new myVersion();
            labelRelease.Content = myver.getRevision(1);
        }

        private void datePicker_SelectedDateChanged(object sender, SelectionChangedEventArgs e)
        {
            if (datePickerEndDate.SelectedDate < datePickerStartDate.SelectedDate)
            {
                MessageBox.Show("Incorrect Date Selection");
            }
        }

        private void HideSelectText()
        {
            if (listBoxCentreSelected.Items.Count > 0)
            {
                textBlockSelectCentre.Visibility = Visibility.Hidden;
            }
            else
            {
                textBlockSelectCentre.Visibility = Visibility.Visible;
            }
            if (listBoxProductSelected.Items.Count > 0)
            {
                textBlockSelectProduct.Visibility = Visibility.Hidden;
            }
            else
            {
                textBlockSelectProduct.Visibility = Visibility.Visible;
            }
        }
                
        private void LoadCentreList()
        {
            try
            {
                iTotalCentre = 0;
                listBoxCentreAll.ItemsSource = null;
                listBoxCentreSelected.ItemsSource = null;
                listBoxCentreAll.Items.Clear();
                listBoxCentreSelected.Items.Clear();
                lsiteAll.Clear();
                lsiteDef.Clear();
                lsiteSel.Clear();

                string sAppendIncludeClose = string.Empty;
                if (checkBoxIncludeCloseCentre.IsChecked == false)
                {
                    sAppendIncludeClose = " WHERE `GroupName` <> 'X' ";
                }
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                using (myCmd = new MySqlCommand(string.Format(@"SELECT `Site`, `RegCentre`
                    FROM `tblregcentre`
                    {0}
                    ORDER BY `RegCentre` ASC;", sAppendIncludeClose), myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            iTotalCentre++;
                            lsiteDef.Add(new ssite { site = myDr.GetValue(0).ToString(), sitename = myDr.GetValue(1).ToString() });
                        }
                    }
                    listBoxCentreAll.ItemsSource = lsiteDef;
                    listBoxCentreAll.DisplayMemberPath = "sitename";
                }
                groupBoxCentre.Header = "Select Centre (" + iTotalCentre + ")";
                ValidataButtonAll();
                DisplayNoResult();
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void LoadProductList()
        {
            try
            {
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                using (myCmd = new MySqlCommand(@"SELECT `ServiceCode`, `ServiceType`
                    FROM `tblvoucher`
                    ORDER BY `ServiceType` ASC;", myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            iTotalProduct++;
                            lprodDef.Add(new sproduct { pcode = myDr.GetValue(0).ToString(), pname = myDr.GetValue(1).ToString() });
                        }
                    }
                    listBoxProductAll.ItemsSource = lprodDef;
                    listBoxProductAll.DisplayMemberPath = "pname";
                }

                using (myCmd = new MySqlCommand(@"SELECT `ServiceCode`, `ServiceName`
                    FROM `tblservicecode`
                    ORDER BY `ServiceName` ASC;", myConn))
                {
                    myCmd.CommandTimeout = 0;
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            iTotalProduct++;
                            lprodDef.Add(new sproduct { pcode = myDr.GetValue(0).ToString(), pname = myDr.GetValue(1).ToString() });
                        }
                    }
                    lprodDef.Sort((x, y) => string.Compare( x.pname, y.pname));

                    listBoxProductAll.ItemsSource = lprodDef;
                    listBoxProductAll.DisplayMemberPath = "pname";
                }
                groupBoxProduct.Header += " (" + iTotalProduct + ")";
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void buttonAddCentre_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (listBoxCentreAll.SelectedIndex != -1)
                {
                    int iSelectedIndex = listBoxCentreAll.SelectedIndex;
                    ssite selectedCentre = listBoxCentreAll.SelectedItem as ssite;
                    lsiteSel.Add(selectedCentre);
                    listBoxCentreSelected.ItemsSource = lsiteSel;
                    listBoxCentreSelected.DisplayMemberPath = "sitename";
                    HideSelectText();
                    RefreshCentreList();
                    listBoxCentreSelected.Items.Refresh();
                    if (listBoxCentreAll.Items.Count != 0)
                    {
                        listBoxCentreAll.SelectedIndex = iSelectedIndex;
                    }
                }
                else
                {
                    MessageBox.Show("Select an item from list");
                }
                DisplayNoResult();
            }
            catch { } // dismiss any posible error
        }

        private void buttonAddCentreAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (ssite site in listBoxCentreAll.Items)
                {
                    lsiteSel.Add(site);
                }
                listBoxCentreSelected.ItemsSource = lsiteSel;
                listBoxCentreSelected.DisplayMemberPath = "sitename";
                RefreshCentreList();
                HideSelectText();
                listBoxCentreSelected.Items.Refresh();
            }
            catch { }
        }

        private void buttonAddProduct_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (listBoxProductAll.SelectedIndex != -1)
                {
                    int iSelectedIndex = listBoxProductAll.SelectedIndex;
                    sproduct selectedProduct = listBoxProductAll.SelectedItem as sproduct;
                    lprodSel.Add(selectedProduct);
                    listBoxProductSelected.ItemsSource = lprodSel;
                    listBoxProductSelected.DisplayMemberPath = "pname";
                    HideSelectText();
                    RefreshProductList();
                    listBoxProductSelected.Items.Refresh();
                    if (listBoxProductAll.Items.Count != 0)
                    {
                        listBoxProductAll.SelectedIndex = iSelectedIndex;
                    }
                }
                else
                {
                    MessageBox.Show("Select an item from list");
                }
                DisplayNoResult();
            }
            catch { } // dismiss any posible error
        }

        private void RefreshCentreList()
        {
            try
            {
                if (textBoxSearchCentre.Text.Length == 0)
                {
                    lsiteAll.Clear();
                    foreach (ssite diff in lsiteDef.Except(lsiteSel)) // for each all item in the default list and it not include in the selected list then add to all centre list
                    {
                        lsiteAll.Add(diff);
                    }
                    listBoxCentreAll.ItemsSource = null;
                    listBoxCentreAll.Items.Clear();
                    listBoxCentreAll.ItemsSource = lsiteAll;
                    listBoxCentreAll.DisplayMemberPath = "sitename";
                    listBoxCentreAll.Items.Refresh();
                }
                else
                {
                    SearchCentreFilter(textBoxSearchCentre.Text);
                }
                ValidataButtonAll();
            }
            catch { } // dismiss error
        }
        
        /// <summary>
        /// Button Add All Centre and Remove All Centre
        /// </summary>
        private void ValidataButtonAll()
        {
            if (listBoxCentreAll.Items.Count == 0)
            {
                buttonAddCentreAll.IsEnabled = false;
            }
            else
            {
                buttonAddCentreAll.IsEnabled = true;
            }

            if (listBoxCentreSelected.Items.Count == 0)
            {
                buttonRemoveCentreAll.IsEnabled = false;
            }
            else
            {
                buttonRemoveCentreAll.IsEnabled = true;
            }
        }

        private void RefreshProductList()
        {
            try
            {
                if (textBoxSearchProduct.Text.Length == 0)
                {
                    lprodAll.Clear();
                    foreach (sproduct diff in lprodDef.Except(lprodSel)) // for each all item in the default list and it not include in the selected list then add to all centre list
                    {
                        lprodAll.Add(diff);
                    }
                    listBoxProductAll.ItemsSource = null;
                    listBoxProductAll.Items.Clear();
                    listBoxProductAll.ItemsSource = lprodAll;
                    listBoxProductAll.DisplayMemberPath = "pname";
                    listBoxProductAll.Items.Refresh();
                }
                else
                {
                    SearchProductFilter(textBoxSearchProduct.Text);
                }
            }
            catch { } // dismiss error
        }

        private void textBoxSearchCentre_GotFocus(object sender, RoutedEventArgs e)
        {
            textBlockSearchCentre.Visibility = Visibility.Hidden;
        }
        
        private void textBoxSearchProduct_GotFocus(object sender, RoutedEventArgs e)
        {
            textBlockSearchProduct.Visibility = Visibility.Hidden;
        }
        
        private void textBoxSearchCentre_LostFocus(object sender, RoutedEventArgs e)
        {
            if (textBoxSearchCentre.Text.Length == 0)
            {
                textBlockSearchCentre.Visibility = Visibility.Visible;
            }
        }

        private void textBoxSearchProduct_LostFocus(object sender, RoutedEventArgs e)
        {
            if (textBlockSearchProduct.Text.Length == 0)
            {
                textBlockSearchProduct.Visibility = Visibility.Visible;
            }
        }

        private void textBoxSearchCentre_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (textBoxSearchCentre.Text.Length == 0)
            {
                RefreshCentreList();
            }
            else
            {
                SearchCentreFilter(textBoxSearchCentre.Text);
            }
            DisplayNoResult();
        }

        private void textBoxSearchProduct_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (textBlockSearchProduct.Text.Length == 0)
            {
                RefreshProductList();
            }
            else
            {
                SearchProductFilter(textBoxSearchProduct.Text);
            }
            DisplayNoResult();
        }

        private void DisplayNoResult()
        {
            if (listBoxCentreAll.Items.Count == 0 && textBoxSearchCentre.Text.Length > 0)
            {
                textBlockCentreNoResult.Visibility = Visibility.Visible;
            }
            else
            {
                textBlockCentreNoResult.Visibility = Visibility.Hidden;
            }

            if (listBoxProductAll.Items.Count == 0 && textBoxSearchProduct.Text.Length > 0)
            {
                textBlockProductNoResult.Visibility = Visibility.Visible;
            }
            else
            {
                textBlockProductNoResult.Visibility = Visibility.Hidden;
            }
        }

        private void SearchCentreFilter(string sFilter)
        {
            try
            {
                lsiteAll.Clear();
                foreach (ssite site in lsiteDef.Except(lsiteSel))
                {
                    if (0 <= site.sitename.IndexOf(sFilter, StringComparison.InvariantCultureIgnoreCase)) // Search Ignore Case
                    {
                        lsiteAll.Add(site);
                    }
                }
                listBoxCentreAll.ItemsSource = null;
                listBoxCentreAll.Items.Clear();
                listBoxCentreAll.ItemsSource = lsiteAll;
                listBoxCentreAll.DisplayMemberPath = "sitename";
                listBoxCentreAll.Items.Refresh();
            }
            catch { }
        }

        private void SearchProductFilter(string sFilter)
        {
            try
            {
                lprodAll.Clear();
                foreach (sproduct prod in lprodDef.Except(lprodSel))
                {
                    if (0 <= prod.pname.IndexOf(sFilter, StringComparison.InvariantCultureIgnoreCase)) // Search Ignore Case
                    {
                        lprodAll.Add(prod);
                    }
                }
                listBoxProductAll.ItemsSource = null;
                listBoxProductAll.Items.Clear();
                listBoxProductAll.ItemsSource = lprodAll;
                listBoxProductAll.DisplayMemberPath = "pname";
                listBoxProductAll.Items.Refresh();
            }
            catch { }
        }

        private void buttonRemoveCentre_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (listBoxCentreSelected.SelectedIndex != -1)
                {
                    int iSelectedIndex = listBoxCentreSelected.SelectedIndex;
                    ssite selectedCentre = listBoxCentreSelected.SelectedItem as ssite;
                    lsiteSel = lsiteSel.Except(lsiteSel.Where(a => a.sitename == selectedCentre.sitename)).ToList();
                    listBoxCentreSelected.ItemsSource = null;
                    listBoxCentreSelected.Items.Clear();
                    listBoxCentreSelected.ItemsSource = lsiteSel;
                    listBoxCentreSelected.DisplayMemberPath = "sitename";
                    listBoxCentreSelected.Items.Refresh();
                    RefreshCentreList();
                    HideSelectText();
                    if (listBoxCentreSelected.Items.Count != 0)
                    {
                        listBoxCentreSelected.SelectedIndex = iSelectedIndex;
                    }
                }
                else
                {
                    MessageBox.Show("Select an item from list");
                }
            }
            catch { }
        }

        private void buttonRemoveCentreAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                foreach (ssite site in listBoxCentreSelected.Items)
                {
                    lsiteAll.Add(site);
                }

                lsiteSel = lsiteSel.Except(lsiteAll).ToList();
                listBoxCentreSelected.ItemsSource = null;
                listBoxCentreSelected.Items.Clear();
                listBoxCentreSelected.ItemsSource = lsiteSel;
                listBoxCentreSelected.DisplayMemberPath = "sitename";
                
                listBoxCentreAll.ItemsSource = lsiteAll;
                listBoxCentreAll.DisplayMemberPath = "sitename";
                RefreshCentreList();
                HideSelectText();
                listBoxCentreSelected.Items.Refresh();
            }
            catch { }
        }
        private void buttonRemoveProduct_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (listBoxProductSelected.SelectedIndex != -1)
                {
                    int iSelectedIndex = listBoxProductSelected.SelectedIndex;
                    sproduct selectedProduct = listBoxProductSelected.SelectedItem as sproduct;
                    lprodSel = lprodSel.Except(lprodSel.Where(a => a.pname == selectedProduct.pname)).ToList();
                    listBoxProductSelected.ItemsSource = null;
                    listBoxProductSelected.Items.Clear();
                    listBoxProductSelected.ItemsSource = lprodSel;
                    listBoxProductSelected.DisplayMemberPath = "pname";
                    listBoxProductSelected.Items.Refresh();
                    RefreshProductList();
                    HideSelectText();
                    if (listBoxProductSelected.Items.Count != 0)
                    {
                        listBoxProductSelected.SelectedIndex = iSelectedIndex;
                    }
                }
                else
                {
                    MessageBox.Show("Select an item from list");
                }
            }
            catch { }
        }

        private void buttonClose_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
            this.Dispose();
        }

        private void buttonGenerate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (datePickerStartDate.SelectedDate == null || datePickerEndDate.SelectedDate == null)
                {
                    throw new Exception("Please check the date selection, it cannot be empty");
                }

                if (listBoxCentreSelected.Items.Count == 0)
                {
                    throw new Exception("No centre is selected...");
                }

                if (listBoxProductSelected.Items.Count == 0)
                {
                    throw new Exception("No product is selected...");
                }
                System.Windows.Forms.SaveFileDialog saveFileExcel = new System.Windows.Forms.SaveFileDialog();
                saveFileExcel.Filter = "Excel files |*.xls|All files (*.*)|*.*";
                saveFileExcel.FilterIndex = 1;
                saveFileExcel.RestoreDirectory = true;
                string fileNameExcel = String.Empty;
                if (saveFileExcel.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    fileNameExcel = saveFileExcel.FileName;
                }
                else
                {
                    throw new Exception("Must provide a file name to save");
                }
                FetchData(OrderByCentre, fileNameExcel);
            }
            catch(Exception ex)
            {
                 MessageBox.Show(ex.Message);
            }
        }

        private void FetchData(int order, string fileNameExcel)
        {
            try
            {
                if (string.IsNullOrEmpty(fileNameExcel)) // generally no required but as extra protection
                {
                    throw new Exception("Must provide a file name to save");
                }
                DateTime dtStartdate = (DateTime)datePickerStartDate.SelectedDate,
                        dtEnddate = (DateTime)datePickerEndDate.SelectedDate;
                string sStartDate = dtStartdate.ToString("yyyy-MM-dd"),
                    sEndDate = dtEnddate.ToString("yyyy-MM-dd"),
                    sSite = string.Empty,
                    sProduct = string.Empty;

                foreach (ssite site in lsiteSel)
                {
                    if (sSite != string.Empty)
                    {
                        sSite += ", ";
                    }
                    sSite += "'" + site.site + "'";
                }

                foreach (sproduct product in lprodSel)
                {
                    if (sProduct != string.Empty)
                    {
                        sProduct += ", ";
                    }
                    sProduct += "'" + product.pcode + "'";
                }

                string sql = @"
SELECT rc.`RegCentre`, jss.`Description`, COUNT(jss.`ID`) AS `Qty`, SUM(jss.`Charge`) AS `SumOfCharge`, SUM(jss.`Commission`) AS `SumOfComm`
FROM `tbljobsheetservice` AS jss
INNER JOIN `tblregcentre` AS rc ON rc.`Site` = jss.`Site`
WHERE jss.`Site` IN ({0})
AND jss.`ServiceCode` IN ({1})
AND jss.`DateCreated` BETWEEN '{2} 00:00:00' AND  '{3} 23:59:59'  
GROUP BY rc.`RegCentre`, jss.`Description`
ORDER BY rc.`RegCentre`, jss.`Description` ASC;
";
                sql = string.Format(sql, sSite, sProduct, sStartDate, sEndDate);
                /*
                StreamWriter FileWriter;
                FileWriter = new StreamWriter("SQL.txt");
                FileWriter.Write(sql); // write to file
                FileWriter.Close(); // close file
                */
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
                worksheet.Name = "Product By Centre Report (Sam)";
                // storing header part in Excel
                int iRow = 1;
                worksheet.Cells[iRow, 1].Value = "Product By Centre Report (Sam)";
                worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                iRow++;
                worksheet.Cells[iRow, 1].Value = string.Format("From {0} to {1}", sStartDate, sEndDate);
                ((Excel.Range)worksheet.Columns["A", Type.Missing]).ColumnWidth = 34;
                ((Excel.Range)worksheet.Columns["B", Type.Missing]).ColumnWidth = 10;
                ((Excel.Range)worksheet.Columns["C", Type.Missing]).ColumnWidth = 14;
                ((Excel.Range)worksheet.Columns["D", Type.Missing]).ColumnWidth = 14;
                worksheet.Cells[1, 1].EntireColumn.NumberFormat = "@"; // Make Like Text
                //worksheet.Cells[1, 2].EntireColumn.NumberFormat = "@"; // Currency Format
                worksheet.Cells[1, 3].EntireColumn.NumberFormat = "#,##0.00_);(#,##0.00)"; // Currency Format
                worksheet.Cells[1, 4].EntireColumn.NumberFormat = "#,##0.00_);(#,##0.00)"; // Currency Format
                if (myConn.State == ConnectionState.Closed) { myConn.Open(); } // Open Connection if is Closed
                string sCentreSite = string.Empty;
                double dTotal = 0, dTotalComm = 0;
                int iRowRead = 0, iQty = 0;
                using (myCmd = new MySqlCommand(sql, myConn))
                {
                    using (myDr = myCmd.ExecuteReader())
                    {
                        while (myDr.Read())
                        {
                            iRowRead++;
                            if (sCentreSite != myDr.GetValue(0).ToString())
                            {
                                if (!string.IsNullOrEmpty(sCentreSite))
                                {
                                    iRow += 2;
                                    worksheet.Cells[iRow, 2].Value = iQty;
                                    worksheet.Cells[iRow, 3].Value = dTotal;
                                    worksheet.Cells[iRow, 4].Value = dTotalComm;
                                    worksheet.Cells[iRow, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    worksheet.Cells[iRow, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                    worksheet.Cells[iRow, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                }
                                iRow += 2;
                                sCentreSite = myDr.GetValue(0).ToString();
                                worksheet.Cells[iRow, 1].Value = "Centre";
                                worksheet.Cells[iRow, 2].Value = sCentreSite;
                                worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                                iRow += 2;
                                worksheet.Cells[iRow, 1].Value = "Product";
                                worksheet.Cells[iRow, 2].Value = "Qty";
                                worksheet.Cells[iRow, 3].Value = "Sum of Charge";
                                worksheet.Cells[iRow, 4].Value = "Sum of Comm";
                                worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                                worksheet.Cells[iRow, 1].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                worksheet.Cells[iRow, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                worksheet.Cells[iRow, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                worksheet.Cells[iRow, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                                iRow++;
                                dTotal = 0;
                                dTotalComm = 0;
                                iQty = 0;
                            }
                            else
                            {
                                iRow++;
                                string sProductCode = myDr.GetValue(1).ToString(); // Product Code
                                int iQtyOfProduct = myDr.GetInt32(2);
                                double dSumOfCharge = myDr.GetDouble(3);
                                double dSumOfComm = myDr.GetDouble(4);
                                worksheet.Cells[iRow, 1].Value = sProductCode;
                                worksheet.Cells[iRow, 2].Value = iQtyOfProduct; // Product Code
                                worksheet.Cells[iRow, 3].Value = dSumOfCharge; // Sell Price
                                worksheet.Cells[iRow, 4].Value = dSumOfComm; // Sell Price
                                //worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                                iQty += iQtyOfProduct;
                                dTotal += dSumOfCharge;
                                dTotalComm += dSumOfComm;
                            }
                        }
                        if (iRowRead > 0)
                        {
                            iRow += 2;
                            worksheet.Cells[iRow, 2].Value = iQty;
                            worksheet.Cells[iRow, 3].Value = dTotal;
                            worksheet.Cells[iRow, 4].Value = dTotalComm;
                            worksheet.Cells[iRow, 2].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            worksheet.Cells[iRow, 3].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            worksheet.Cells[iRow, 4].Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                            iRow += 2;
                            worksheet.Cells[iRow, 2].Value = string.Format("Total Record ({0})", iRowRead);
                            worksheet.Cells[iRow, 1].EntireRow.Font.Bold = true; // Make Font Bold
                        }
                    }
                }
                if (myConn.State == ConnectionState.Open) { myConn.Close(); } // Close Connection if is Open
                //Fixed-old code :11 para->add 1:Type.Missing
                workbook.SaveAs(fileNameExcel, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                // Exit from the application
                app.Quit();
                MessageBox.Show("Excel file is successfully created");
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }

        private void checkBoxIncludeCloseCentre_Checked(object sender, RoutedEventArgs e)
        {
            LoadCentreList();
        }

        private void checkBoxIncludeCloseCentre_Unchecked(object sender, RoutedEventArgs e)
        {
            LoadCentreList();
        }
    }
}
