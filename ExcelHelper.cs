using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using Missing = System.Reflection.Missing;

namespace EskomParser
{
    public class ExcelHelper : IDisposable
    {
        #region Fields

        private Excel.Application _excel;
        private int _workingColumn;
        private Excel.Worksheet _worksheet;
        private Excel.Workbook _workbook;

        #endregion

        #region Properties

        public int WorkingColumn
        {
            get { return _workingColumn; }
            set { _workingColumn = value; }
        }

        #endregion

        #region Methods

        public void Close()
        {
            _excel.Workbooks[1].Close(true, Missing.Value, Missing.Value);
            _excel.Quit();
            Marshal.ReleaseComObject(_workbook);
            Marshal.ReleaseComObject(_worksheet);
            Marshal.ReleaseComObject(_excel);
            _excel = null;
            _worksheet = null;
            _workbook = null;
        }

        public void Dispose()
        {
            if (_excel != null)
            {
                Marshal.ReleaseComObject(_excel);
            }
        }

        public DateTime? GetDateTime(int row, int column)
        {
            DateTime? value = null;

            var cell = (Excel.Range)_worksheet.Cells[row, column];
            if (cell.Value != null)
            {
                if (cell.Value is DateTime)
                {
                    value = (DateTime?)cell.Value;
                }
            }

            Marshal.ReleaseComObject(cell);
            cell = null;

            return value;
        }

        public string GetText(int row, int column)
        {
            var cell = (Excel.Range) _worksheet.Cells[row, column];
            string value = string.Empty;

            if (cell.Value != null)
            {
                value = cell.Value.ToString();
            }
            
            Marshal.ReleaseComObject(cell);
            cell = null;

            return value;
        }

        public Excel.Workbook GetWorkbook(int index)
        {
            return _excel.Workbooks[index];
        }

        public Excel.Worksheet GetWorksheet(Excel.Workbook workbook, int index)
        {
            return (Excel.Worksheet)workbook.Worksheets[index];
        }

        public void Initialize()
        {
            Initialize(false);
        }

        public void Initialize(bool visible)
        {
            _excel = new Excel.Application();
            _excel.Visible = visible;
        }

        public void Load(string excelFile)
        {
            _excel.Workbooks.Open(excelFile, Missing.Value, Missing.Value, Missing.Value, 
                Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }

        public int RowOf(string value, int column)
        {
            int returnValue = -1;
            int row = 1;
            int lastEmptyRow = -1;
            int emptyRowCount = 0;

            while (returnValue == -1)
            {
                var stringValue = GetText(row, column);
                if (!string.IsNullOrEmpty(stringValue))
                {
                    if (stringValue.Equals(value))
                    {
                        returnValue = row;
                    }
                    else
                    {
                        lastEmptyRow = -1;
                    }
                }
                else
                {
                    if (lastEmptyRow == -1)
                    {
                        lastEmptyRow = row;
                    }
                    else if (lastEmptyRow + 1 == row)
                    {
                        lastEmptyRow = row;
                        emptyRowCount++;
                        if (emptyRowCount > 20)
                        {
                            break;
                        }
                    }
                }
                row++;
            }

            return returnValue;
        }

        public void Save()
        {
            _excel.Workbooks[1].Save();
        }

        public void SaveAs(string excelFile)
        {
            _excel.Workbooks[1].SaveAs(excelFile, Missing.Value, Missing.Value, Missing.Value,
                                       Missing.Value, Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive,
                                       Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
        }

        public void SetDecimal(decimal value, int row, int column)
        {
            var cell = (Excel.Range)_worksheet.Cells[row, column];
            cell.Value = value;

            Marshal.ReleaseComObject(cell);
            cell = null;
        }

        public void SetDouble(double value, int row, int column)
        {
            var cell = (Excel.Range)_worksheet.Cells[row, column];
            cell.Value = value;

            Marshal.ReleaseComObject(cell);
            cell = null;
        }

        public void SetInteger(int value, int row, int column)
        {
            var cell = (Excel.Range)_worksheet.Cells[row, column];
            cell.Value = value;

            Marshal.ReleaseComObject(cell);
            cell = null;
        }

        public void SetText(string text, int row, int column)
        {
            var cell = (Excel.Range) _worksheet.Cells[row, column];
            cell.Value = text;

            Marshal.ReleaseComObject(cell);
            cell = null;
        }

        public void SetWorkingColumn(int month, int year)
        {
            bool foundColumn = false;
            int columnCounter = 1;

            while (!foundColumn)
            {
                var value = GetDateTime(1, columnCounter);
                if (value != null)
                {
                    if ((value.Value.Month == month) &&
                        (value.Value.Year == year))
                    {
                        foundColumn = true;
                        WorkingColumn = columnCounter;
                    }
                }
                columnCounter++;
            }

        }

        public void SetWorksheet(int index)
        {
            _workbook = GetWorkbook(1);
            _worksheet = GetWorksheet(_workbook, index);
        }

        #endregion
    }
}
