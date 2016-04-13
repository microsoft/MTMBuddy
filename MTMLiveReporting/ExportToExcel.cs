
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

using System.Runtime.InteropServices;
using System.Data.OleDb;
using System.Data;
using System.Collections;
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Globalization;
using ReportingLayer;


namespace MTMLiveReporting
{
    public class ExportToExcel<T, U>
        where T : class
        where U : List<T>
    {
        public List<T> dataToExport;
        private Excel.Application _excelApp = null;
        private Excel.Workbooks _books = null;
        private Excel._Workbook _book = null;
        private Excel.Sheets _sheets = null;
        private Excel._Worksheet _sheet = null;
        private Excel.Range _range = null;
        private Excel.Font _font = null;

        public void GenerateExcel()
        {
            if (dataToExport != null)
            {
                if (dataToExport.Count != 0)
                {
                    //MergeData();
                    // Create Excel Objects
                    _excelApp = new Excel.Application();
                    _books = (Excel.Workbooks)_excelApp.Workbooks;
                    //_book = (Excel._Workbook)(_books.);
                    _book = _excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                    _sheets = (Excel.Sheets)_book.Worksheets;
                    //_sheet = (Excel._Worksheet)(_sheets.get_Item(1));
                    _sheet = (Excel.Worksheet)_book.Sheets[1];
                    // Create Header Row
                    PropertyInfo[] headerInfo = typeof(T).GetProperties();
                    List<object> objHeaders = new List<object>();
                    for (int n = 0; n < headerInfo.Length; n++)
                    {
                        objHeaders.Add(headerInfo[n].Name);
                    }

                    var headerToAdd = objHeaders.ToArray();
                    AddExcelRows("A1", 1, headerToAdd.Length, headerToAdd);
                    _font = _range.Font;
                    _font.Bold = true;

                    // Write Data
                    object[,] objData = new object[dataToExport.Count, headerToAdd.Length];

                    for (int j = 0; j < dataToExport.Count; j++)
                    {
                        var item = dataToExport[j];
                        for (int i = 0; i < headerToAdd.Length; i++)
                        {
                            var y = typeof(T).InvokeMember(headerToAdd[i].ToString(),
                            BindingFlags.GetProperty, null, item, null);
                            objData[j, i] = (y == null) ? "" : y.ToString();
                        }
                    }
                    AddExcelRows("A2", dataToExport.Count, headerToAdd.Length, objData);
                    _range = _sheet.get_Range("A1");
                    _range = _range.get_Resize(dataToExport.Count + 1, headerToAdd.Length);
                    _range.Columns.AutoFit();
                    var path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments), "Reports","ReportData_" + typeof(T).Name + ".xlsx");
                    _book.SaveCopyAs(path);
                    _excelApp.Visible = true;
                }


                
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_sheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_sheets);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_book);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_books);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(_excelApp);
                _sheet = null;
                _sheets = null;
                _book = null;
                _books = null;
                _excelApp = null;
                GC.Collect();

                         }
        }


        private void AddExcelRows(string startRange, int rowCount, int colCount, object values)
        {
            _range = _sheet.get_Range(startRange);
            _range = _range.get_Resize(rowCount, colCount);
            _range.set_Value(Missing.Value, values);
        }
    }



    public static class ExcelToDS
    {
        public static DataSet getDS(string strfilename)
        {
            


           

            _excelApp = new Microsoft.Office.Interop.Excel.Application();
            _excelApp.Visible = false;
            _workBook = _excelApp.Workbooks.Open(strfilename, Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value, Missing.Value,
                                        Missing.Value, Missing.Value);
            _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_workBook.Worksheets.get_Item(1);

            DataSet dataSet = ReadExcelAsDataSet(_workSheet);
            _excelApp.Quit();
            return dataSet;
        }

        public static List<string> getsheets(string strfilename)
        {
            List<string> excelsheets = new List<string>();
            OleDbConnection conn;
            string ConnectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;" + @"Data Source=" + strfilename + ";" + @"Extended Properties=Excel 12.0;";


            conn = new OleDbConnection(ConnectionString);
            conn.Open();
            System.Data.DataTable dt = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            foreach (DataRow drow in dt.Rows)
                excelsheets.Add(drow["TABLE_NAME"].ToString());

            conn.Close();
            return excelsheets;
        }
        #region Data Driven Test Scripts - Excel Helper
        private static Microsoft.Office.Interop.Excel.Application _excelApp = null;
        private static Microsoft.Office.Interop.Excel.Workbook _workBook = null;
        private static Microsoft.Office.Interop.Excel.Worksheet _workSheet = null;
        private static Microsoft.Office.Interop.Excel.Range _cellRange = null;
        /// <summary>
        /// To dispose all instances of Excel App
        /// </summary>
        public static void DisposeExcelAppTraces()
        {
            
                _excelApp.Workbooks.Close();
                _excelApp.Quit();

                Marshal.ReleaseComObject(_workBook);
                _workBook = null;
                Marshal.ReleaseComObject(_cellRange);
                _cellRange = null;
                Marshal.ReleaseComObject(_excelApp);
                _excelApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();

            

        }
        /// <summary>
        /// Reads cell values of a Excel Row
        /// </summary>
        /// <param name="workSheet"></param>
        /// <param name="rowIndex"></param>
        /// <param name="columnCount"></param>
        /// <returns></returns>
        public static string[] ReadExcelRow(Microsoft.Office.Interop.Excel.Worksheet workSheet, int rowIndex, int columnCount = 0)
        {
            ArrayList strArrayList = new ArrayList();
            try
            {

                if (columnCount == 0) { columnCount = workSheet.UsedRange.Columns.Count; }
                for (int c = 1; c <= columnCount; c++)
                {
                    string cellValue = (string)workSheet.Cells[rowIndex, c].Text();
                    strArrayList.Add(cellValue);
                }
            }
            catch (Exception)
            {
                DisposeExcelAppTraces();
            }
            return (string[])strArrayList.ToArray(typeof(string));
        }

        /// <summary>
        /// Read Excel File as Data Set
        /// </summary>
        /// <param name="WorkSheetName"></param>
        /// <returns></returns>
        public static DataSet ReadExcelAsDataSet(Microsoft.Office.Interop.Excel.Worksheet _workSheet)
        {
            DataSet rsltDataSet = new DataSet();
            System.Data.DataTable dataTable = new System.Data.DataTable();
            try
            {
                int totalRows = _workSheet.UsedRange.Rows.Count;
                int totalColumns = _workSheet.UsedRange.Columns.Count;

                // Read first row from excel and add them Headers to Data Table
                string[] headerColumns = ReadExcelRow(_workSheet, 1, totalColumns);
                for (int c = 0; c < totalColumns; c++)
                {
                    dataTable.Columns.Add(new DataColumn(headerColumns[c], typeof(string)));
                }

                // Read rest of the excel file and add them to data table
                for (int i = 2; i <= totalRows; i++)
                {
                    string[] rowValues = ReadExcelRow(_workSheet, i, totalColumns);
                    DataRow dataRow = dataTable.NewRow();
                    for (int j = 0; j < totalColumns; j++)
                    {
                        dataRow[j] = rowValues[j];
                    }
                    dataTable.Rows.Add(dataRow);
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Error: " + ex.Message);
            }
            rsltDataSet.Tables.Add(dataTable);
            return rsltDataSet;
        }

        /// <summary>
        /// Method to read list rows from excel
        /// </summary>
        /// <param name="filterType"></param>
        /// <param name="filterValue"></param>
        /// <returns></returns>
        public static List<DataRow> ReadDataFromExcel(string ExcelFilePath, Dictionary<string, string> filterCriteria = null)
        {
            List<DataRow> dataRowsLst = new List<DataRow>();

            try
            {
                _excelApp = new Microsoft.Office.Interop.Excel.Application();
                _excelApp.Visible = false;
                _workBook = _excelApp.Workbooks.Open(ExcelFilePath, Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Missing.Value, Missing.Value,
                                            Missing.Value, Missing.Value);
                _workSheet = (Microsoft.Office.Interop.Excel.Worksheet)_workBook.Worksheets.get_Item(1);

                DataSet dataSet = ReadExcelAsDataSet(_workSheet);
                foreach (DataRow dataRow in dataSet.Tables[0].Rows)
                {
                    bool matched = true;
                    foreach (string key in filterCriteria.Keys)
                    {
                        if (!dataRow[key].ToString().Equals(filterCriteria[key].ToString()))
                        {
                            matched = false;
                        }
                    }
                    if (matched)
                        dataRowsLst.Add(dataRow);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception occurred in ReadDataFromExcel() method. Exception: " + ex.Message + ":" + ex.StackTrace);
                throw new Exception();
            }
            finally
            {
                DisposeExcelAppTraces();
            }

            return dataRowsLst;
        }

        #endregion

    }

    public static class DStoExcel
    {
        public static bool dropExcel(DataSet ds, string path)
        {
            Application newExcel = new Application();
            newExcel.Visible = false;
            Workbook wbk = newExcel.Workbooks.Add();
            Worksheet wrksht = wbk.Worksheets.Add(Type.Missing);
            System.Data.DataTable table = ds.Tables[0];
            for (int i = 0; i < table.Columns.Count; i++)
            {
                wrksht.Cells[1, i + 1] = table.Columns[i].ColumnName;
            }
            //    wrksht.Cells[1, 1] = table.Columns[0].ColumnName;
            //wrksht.Cells[1,2]="Result";
            //wrksht.Cells[1,3]="Comments";
            //wrksht.Cells[1,4]="State";
            int r = 2;

            foreach (DataRow row in table.Rows)
            {
                for (int c = 1; c <= table.Columns.Count; c++)
                {
                    wrksht.Cells[r, c] = row[c - 1];
                }

                r++;
            }


            try
            {
                wbk.SaveAs(path);
                wbk.Save();
                wbk.Close(false);
            }
            catch
            {
                return false;
            }
            newExcel.Quit();
            return true;

        }
    }
}
