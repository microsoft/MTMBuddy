//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 

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
    public class ExportToExcel<T, TU>
        where T : class
        where TU : List<T>
    {
        public List<T> DataToExport;
        private Excel.Application _excelApp = null;
        private Excel.Workbooks _books = null;
        private Excel._Workbook _book = null;
        private Excel.Sheets _sheets = null;
        private Excel._Worksheet _sheet = null;
        private Excel.Range _range = null;
        private Excel.Font _font = null;

        public void GenerateExcel()
        {
            if (DataToExport != null)
            {
                if (DataToExport.Count != 0)
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
                    object[,] objData = new object[DataToExport.Count, headerToAdd.Length];

                    for (int j = 0; j < DataToExport.Count; j++)
                    {
                        var item = DataToExport[j];
                        for (int i = 0; i < headerToAdd.Length; i++)
                        {
                            var y = typeof(T).InvokeMember(headerToAdd[i].ToString(),
                            BindingFlags.GetProperty, null, item, null);
                            objData[j, i] = (y == null) ? "" : y.ToString();
                        }
                    }
                    AddExcelRows("A2", DataToExport.Count, headerToAdd.Length, objData);
                    _range = _sheet.get_Range("A1");
                    _range = _range.get_Resize(DataToExport.Count + 1, headerToAdd.Length);
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



   
}
