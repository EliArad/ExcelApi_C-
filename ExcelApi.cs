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

            foreach (Process clsProcess in Process.GetProcesses())
            {
                if (clsProcess.ProcessName.Equals("EXCEL"))
                {
                    clsProcess.Kill();
                }
            }

            app = new Excel.Application();
            app.Visible = false;

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
                    if (ws[i + 1].Name == oldName)
                    {
                        ws[i + 1].Name = newName;
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
                ws[index + 1].Name = newName;
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

        public bool AddWorkSheetAtThebegin(string name, out string outMessage)
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
        public void Close()
        {
            if (wb != null)
            {
                wb.Save();
                wb.Close();
            }
            m_workbookOpen = false;
            if (app != null)
                app.Quit();
        }

        public int TotalRows(int sheetIndex)
        {
            return ws[sheetIndex + 1].Rows.Count;
        }
        public int TotalCols(int sheetIndex)
        { 
          return ws[sheetIndex + 1].Columns.Count;
        }

        public bool ReadCell(int sheetIndex, int rowIndex, int colIndex, object value, out string outMessage)
        {
            try
            {
                outMessage = string.Empty;
                value = ws[sheetIndex + 1].Cells[rowIndex, 1].Value;
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
            outMessage = string.Empty;
            try
            {
                for (int i = 0; i < data.Count; i++)
                {
                    ws[sheetIndex + 1].Cells[startRowIndex, startColIndex + i] = data[i];
                }
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
                    ws[sheetIndex + 1].Cells[startRowIndex, startColIndex + i] = name;
                    i++;
                }
                i = 0;
                
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                 BindingFlags.NonPublic |
                                                                 BindingFlags.Public))
                {
                    //Console.WriteLine("{0} = {1}", field.Name, field.GetValue(s));
                    ws[sheetIndex + 1].Cells[startRowIndex + 1, startColIndex + i] = field.GetValue(s);
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

        public bool WriteStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, List<T> s, out string outMessage)
        {

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
                    ws[sheetIndex + 1].Cells[startRowIndex, startColIndex + i] = name;
                    i++;
                }

                for (int index = 0; index < s.Count; index++)
                {
                    i = 0;
                    foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                     BindingFlags.NonPublic |
                                                                     BindingFlags.Public))
                    {
                        //Console.WriteLine("{0} = {1}", field.Name, field.GetValue(s[index]));
                        ws[sheetIndex + 1].Cells[startRowIndex + 1 + index, startColIndex + i] = field.GetValue(s[index]);
                        i++;
                    }
                }
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
               
                ws[sheetIndex + 1].Cells[rowIndex, colIndex] = value;
                //ws[sheetIndex + 1].Cells[xname, 1].Font.Bold = true;
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

                ws[sheetIndex + 1].Cells[rowIndex, colIndex] = value;
                if (bold == true)
                    ws[sheetIndex + 1].Cells[rowIndex, colIndex].Font.Bold = true;

                 ws[sheetIndex + 1].Cells[rowIndex, colIndex].Font.Color = foreColor;           
                 ws[sheetIndex + 1].Cells[rowIndex, colIndex].interior.color = backColor;

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
                Excel.Range xlRange = ws[sheetIndex + 1].UsedRange;

                for (int index = 0; index < rowCount; index++)
                {
                    i = 0;
                    T x = new T();
                        
                    foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                     BindingFlags.NonPublic |
                                                                     BindingFlags.Public))
                    {
                        object d = xlRange.Cells[startRowIndex, startColIndex + i].Value2.ToString();
                        field.SetValue(x, d);
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

           
            Excel.Range xlRange = ws[sheetIndex + 1].UsedRange;

            outMessage = string.Empty;
                      
            try
            {

                int i = 0;
                foreach (var field in typeof(T).GetFields(BindingFlags.Instance |
                                                                 BindingFlags.NonPublic |
                                                                 BindingFlags.Public))
                {

                    object d = xlRange.Cells[startRowIndex, startColIndex + i].Value2.ToString();
                    field.SetValue(s,  d);
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
    }
}
