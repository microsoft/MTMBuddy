//------------------------------------------------------------------------------------------------------- 
// Copyright (C) Microsoft. All rights reserved. 
// Licensed under the MIT license. See LICENSE.txt file in the project root for full license information. 
//------------------------------------------------------------------------------------------------------- 

using System;
using System.Collections.Generic;

using System.Reflection;

using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Linq;


namespace MTMLiveReporting
{
    public class ExportToExcel<T>
        where T : class
    {
        public List<T> DataToExport;
        private Application _excelApp;
        private Workbooks _books;
        private _Workbook _book ;
        private Sheets _sheets ;
        private _Worksheet _sheet ;
        private Range _range;
        private Font _font ;

        public void GenerateExcel()
        {
            if (DataToExport == null) return;
            if (DataToExport.Count != 0)
            {
                
                // Create Excel Objects
                _excelApp = new Application();
                _books = _excelApp.Workbooks;
               
                _book = _excelApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
                _sheets = _book.Worksheets;
                
                _sheet = (Worksheet)_book.Sheets[1];
                // Create Header Row
                PropertyInfo[] headerInfo = typeof(T).GetProperties();

                var headerToAdd = headerInfo.Select(t => t.Name).Cast<object>().ToArray();
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
                        objData[j, i] = y?.ToString() ?? "";
                    }
                }
                AddExcelRows("A2", DataToExport.Count, headerToAdd.Length, objData);
                _range = _sheet.Range["A1"];
                _range = _range.Resize[DataToExport.Count + 1, headerToAdd.Length];
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


        private void AddExcelRows(string startRange, int rowCount, int colCount, object values)
        {
            _range = _sheet.Range[startRange];
            _range = _range.Resize[rowCount, colCount];
            _range.set_Value(Missing.Value, values);
        }
    }



   
}
