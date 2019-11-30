using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelLib
{
    public class ExcelApi
    {
        Excel.Application app = null;
        Excel._Workbook wb = null;
        bool m_workbookOpen = false;
        Excel.Sheets ws;

        public ExcelApi(bool visible = false)
        {
            app = new Excel.Application();
            app.Visible = visible;
        }
        public static void CloseExcel()
        {
            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Equals("EXCEL"))
                {
                    clsProcess.Kill();
                }
            }

        }
        public bool NewFile(string fileName)
        {
            if (File.Exists(fileName) == false)
            {
                wb = app.Workbooks.Add();
                wb.SaveAs(fileName);
                m_workbookOpen = true;
                ws = wb.Worksheets;
                return true;
            }
            return false;
        }
        public bool OpenFile(string fileName)
        {
            if (File.Exists(fileName) == false)
                return false;

            wb = app.Workbooks.Open(fileName);
            ws = wb.Worksheets;

            return true;
        }
        public int SheetCount
        {
            get
            {

                return ws.Count;
            }
        }

        public string SheetName(int index)
        {
            return ws[index].Name;
        }

        public void Save(string fileName = "")
        {
            if (fileName == string.Empty)
                wb.Save();
            else
                wb.SaveAs(fileName);
        }
        public bool AddWorkSheetAtTheEnd(string name, out string outMessage)
        {
            outMessage = string.Empty;

            try
            {
                var xlNewSheet = (Excel.Worksheet)ws.Add(Type.Missing, ws[ws.Count], Type.Missing, Type.Missing);
                xlNewSheet.Name = name;
                //xlNewSheet.Cells[1, 1] = "New 555555 content";
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool UpdateSheetName(string oldName, string newName, out string outMessage)
        {
            outMessage = string.Empty;
            try
            {
                for (int i = 0; i < ws.Count; i++)
                {
                    if (ws[i].Name == oldName)
                    {
                        ws[i].Name = newName;
                        break;
                    }
                }
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool UpdateSheetName(int index, string newName, out string outMessage)
        {
            outMessage = string.Empty;
            try
            {
                ws[index].Name = newName;
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool UpdateFirstSheetName(string newName, out string outMessage)
        {
            outMessage = string.Empty;
            try
            {
                ws[1].Name = newName;
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool UpdateLastSheetName(string newName, out string outMessage)
        {
            outMessage = string.Empty;
            try
            {
                ws[ws.Count].Name = newName;
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool AddWorkSheetAtTheBegin(string name, out string outMessage)
        {
            outMessage = string.Empty;

            try
            {
                var xlNewSheet = (Excel.Worksheet)ws.Add(ws[1], Type.Missing, Type.Missing, Type.Missing);
                xlNewSheet.Name = name;
                //xlNewSheet.Cells[1, 1] = "New 555555 content";
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }
        public void Close(bool terminate)
        {
            if (wb != null)
            {
                wb.Save();
                wb.Close();
                wb = null;
            }
            m_workbookOpen = false;
            if (app != null && terminate == true)
            {
                app.Quit();
                app = null;
            }
        }

        public int TotalRows(int sheetIndex)
        {
            if (sheetIndex < 1)
            {
                throw (new SystemException("Sheet index start from 1"));
            }
            return ws[sheetIndex].Rows.Count;
        }
        public int TotalCols(int sheetIndex)
        {
            if (sheetIndex < 1)
            {
                throw (new SystemException("Sheet index start from 1"));
            }
            return ws[sheetIndex].Columns.Count;
        }

        public bool ReadCell(int sheetIndex, int rowIndex, int colIndex, out object value, out string outMessage)
        {
            if (sheetIndex < 1)
            {
                value = 0;
                outMessage = "Sheet index start from 1";
                return false;
            }
            value = 0;
            try
            {
                outMessage = string.Empty;
                value = ws[sheetIndex].Cells[rowIndex, colIndex].Value;
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }
        public bool ReadCell(int sheetIndex, int rowIndex, int colIndex, out string value, out string outMessage)
        {
            if (sheetIndex < 1)
            {
                value = string.Empty;
                outMessage = "Sheet index start from 1";
                return false;
            }

            value = string.Empty;
            try
            {
                outMessage = string.Empty;
                value = ws[sheetIndex].Cells[rowIndex, colIndex].Value;
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool WriteLine(int sheetIndex, int startRowIndex, int startColIndex, List<object> data, out string outMessage)
        {
            if (sheetIndex  < 1)
            {
                outMessage = "Sheet index start from 1";
                return false;
            }
            outMessage = string.Empty;
            try
            {
                object[,] ar = new object[1, data.Count];
                for (int i = 0; i < data.Count; i++)
                {
                    ar[0,i] = data[i];
                }
                WriteArray<object>(sheetIndex, startRowIndex, startColIndex, ar);
                /*
                
                */
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool WriteStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, T s, out string outMessage)
        {
            if (sheetIndex < 1)
            {              
                outMessage = "Sheet index start from 1";
                return false;
            }

            outMessage = string.Empty;
            try
            {
                int i = 0;
                // write the header 
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                 BindingFlags.NonPublic |
                                                                 BindingFlags.Public))
                {
                    //Console.WriteLine("{0} = {1}", field.Name, field.GetValue(s));
                    string name = field.Name;
                    string[] sname = name.Split('>');
                    name = sname[0].Trim('<');
                    ws[sheetIndex].Cells[startRowIndex, startColIndex + i] = name;
                    i++;
                }
                i = 0;
                
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                 BindingFlags.NonPublic |
                                                                 BindingFlags.Public))
                {
                    //Console.WriteLine("{0} = {1}", field.Name, field.GetValue(s));
                    ws[sheetIndex].Cells[startRowIndex + 1, startColIndex + i] = field.GetValue(s);
                    i++;
                }
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }

            return true;

        }

        public void WriteArray<T>(int sheetIndex,int startRow, int startColumn, T[,] array)
        {
            var row = array.GetLength(0);
            var col = array.GetLength(1);
            Range c1 = (Range)ws[sheetIndex].Cells[startRow, startColumn];
            Range c2 = (Range)ws[sheetIndex].Cells[startRow + row - 1, startColumn + col - 1];
            Range range = ws[sheetIndex].Range[c1, c2];
            range.Value = array;
        }

        public bool WriteStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, List<T> s, out string outMessage)
        {
            if (sheetIndex < 1)
            {
                outMessage = "Sheet index start from 1";
                return false;
            }

            outMessage = string.Empty;
            try
            {
                int i = 0;
                // write the header 
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                 BindingFlags.NonPublic |
                                                                 BindingFlags.Public))
                {
                    //Console.WriteLine("{0} = {1}", field.Name, field.GetValue(s));
                    string name = field.Name;
                    string[] sname = name.Split('>');
                    name = sname[0].Trim('<');
                    ws[sheetIndex].Cells[startRowIndex, startColIndex + i] = name;
                    i++;
                }

                object[,] data = new object[s.Count, i];
                int startColIndex1 = 0;
                int startRowIndex1 = 0;
                for (int index = 0; index < s.Count; index++)
                {
                    startColIndex1 = 0;
                    i = 0;
                    foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                     BindingFlags.NonPublic |
                                                                     BindingFlags.Public))
                    {
                        //Console.WriteLine("{0} = {1}", field.Name, field.GetValue(s[index]));
                        //ws[sheetIndex].Cells[startRowIndex + 1 + index, startColIndex + i] = field.GetValue(s[index]);
                        data[startRowIndex1, startColIndex1 + i] = field.GetValue(s[index]);
                        i++;
                    }
                    startRowIndex1++;
                }
                WriteArray(sheetIndex, startRowIndex + 1, startColIndex, data);
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
            return true;
        }

        public bool WriteCell(int sheetIndex , 
                              int rowIndex, 
                              int colIndex, 
                              object value, 
                              out string outMessage)
        {
            outMessage = string.Empty;

            try
            {
               
                ws[sheetIndex].Cells[rowIndex, colIndex] = value;
                //ws[sheetIndex].Cells[xname, 1].Font.Bold = true;
                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;            
                return false;
            }
        }

        public bool WriteCell(int sheetIndex,
                              int rowIndex,
                              int colIndex,
                              object value,
                              bool bold,
                              Color foreColor,
                              Color backColor,
                              out string outMessage)
        {
            outMessage = string.Empty;

            try
            {

                ws[sheetIndex].Cells[rowIndex, colIndex] = value;
                if (bold == true)
                    ws[sheetIndex].Cells[rowIndex, colIndex].Font.Bold = true;

                 ws[sheetIndex].Cells[rowIndex, colIndex].Font.Color = foreColor;           
                 ws[sheetIndex].Cells[rowIndex, colIndex].interior.color = backColor;

                return true;
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
        }

        public bool ReadStruct<T>(int sheetIndex, 
                                  int startRowIndex, 
                                  int startColIndex, 
                                  ref List<T> s, 
                                  int rowCount, 
                                  out string outMessage) where T : class, new()
        {

            outMessage = string.Empty;
            try
            {
                int i = 0;


                int numOfFields = 0;
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                     BindingFlags.NonPublic |
                                                                     BindingFlags.Public))
                {
                    numOfFields++;
                }


                Range range = (Excel.Range)ws[sheetIndex].Range[ws[sheetIndex].Cells[startRowIndex, startColIndex], ws[sheetIndex].Cells[startRowIndex + rowCount, startColIndex + numOfFields]];
                 
                object[,] values = (object[,])range.Value2;

                startRowIndex = 1;
               
                for (int index = 0; index <= rowCount; index++)
                {
                    i = 0;
                    T x = new T();
                    startColIndex = 1;
                    foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                     BindingFlags.NonPublic |
                                                                     BindingFlags.Public))
                    {
                        string ft = field.FieldType.Name;
                        object d = values[startRowIndex, startColIndex + i];
                        if (d == null)
                        {
                            continue;
                        }
                        string ft1 = d.GetType().Name;
                        if (ft != ft1)
                        {
                            if (ft == "String")
                            {
                                field.SetValue(x, d.ToString());
                            }
                        }
                        else
                        {
                            field.SetValue(x, d);
                        }
                        i++;
                    }
                    s.Add(x);
                    startRowIndex++;
                }
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
            return true;
        }
        public bool ReadStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, ref T s, out string outMessage) where T : class
        {

          


            outMessage = string.Empty;

              

            try
            {

                int i = 0;
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                 BindingFlags.NonPublic |
                                                                 BindingFlags.Public))
                {

                    //Console.WriteLine("{0}", field.Name);
                    string ft = field.FieldType.Name;
                    ReadCell(sheetIndex, startRowIndex, startColIndex + i, out object d, out outMessage);
                    if (d == null)
                    {
                        continue;
                    }
                    string ft1 = d.GetType().Name;
                    if (ft != ft1)
                    {
                        if (ft == "String")
                        {
                            field.SetValue(s, d.ToString());
                        }
                    }
                    else
                    {
                        field.SetValue(s, d);
                    }
                    i++;
                }
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }

            return true;

        }

        public bool ReadColumnList(int sheetIndex,
                             int startRowIndex,
                             int startColIndex,
                             out List<object> list,
                             int rowCount,
                             out string outMessage)
        {

            outMessage = string.Empty;
            list = new List<object>();
            try
            {

                Range range = (Excel.Range)ws[sheetIndex].Range[ws[sheetIndex].Cells[startRowIndex, startColIndex], ws[sheetIndex].Cells[startRowIndex + rowCount, startColIndex]];
                object[,] values = (object[,])range.Value2;

                for (int i = 1; i <= rowCount; i++)
                {                   
                    list.Add(values[1,i]);
                }
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
            return true;
        }
        public string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public int ExcelColumnNameToNumber(string col_name)
        {
            int result = 0;

            // Process each letter.
            for (int i = 0; i < col_name.Length; i++)
            {
                result *= 26;
                char letter = col_name[i];

                // See if it's out of bounds.
                if (letter < 'A') letter = 'A';
                if (letter > 'Z') letter = 'Z';

                // Add in the value of this letter.
                result += (int)letter - (int)'A' + 1;
            }
            return result;
        }
        public bool ReadRowList(int sheetIndex,
                                int startRowIndex,
                                int startColIndex,
                                out List<object> list,
                                int colCount,
                                out string outMessage)
        {

            outMessage = string.Empty;
            list = new List<object>();
            try
            {

                Range range = (Excel.Range)ws[sheetIndex].Range[ws[sheetIndex].Cells[startRowIndex, startColIndex], ws[sheetIndex].Cells[startRowIndex, startColIndex + colCount]];
                object[,] values = (object[,])range.Value2;


                for (int i = 1; i <= colCount; i++)
                {                    
                    list.Add(values[1,i]);
                }
            }
            catch (Exception err)
            {
                outMessage = err.Message;
                return false;
            }
            return true;
        }
    }
}
