Excel API -Eli Arad

My intend was to work with strucure and list of structure in excel interop excel automaticly.
the functions should be simple and self explain to write fast structures to excel

for example:

			Employee emp = new Employee
            {
                Email_ID = "eeee",
                Employee_ID = "42323232",
                Name = "Eli Arad",
                Number = "e11999"
            };

            Employee emp1 = new Employee
            {
                Email_ID = "Eli",
                Employee_ID = "027104918",
                Name = "El9999",
                Number = "050480277"
            };

            List<Employee> employees = new List<Employee>();
			employees.Add(emp);
            employees.Add(emp1);

here we have  a two structures and we want to write them with the header and the data using simple function.
The function will use the Range capatabiliy to write the data fast:

 ExcelApi.CloseExcel(); // this function close all running excel process
 ExcelApi t = new ExcelApi();  // initialize the API
           
            string fileName = "c:\\MyExcelFile.xlsx";
            if (t.NewFile(fileName) == true)
            {
            
                if (t.WriteStruct<Employee>(1, 10, 2, employees, out outMessage) == false)
                {

                }
		    }
we see here that we have  a WriteStruct function which is a template base struct.
The function uses the c# reflection to get the name of the class fields for the header and the field values.

To read the structure i used the same tecnique.

 List<Employee> remp1 = new List<Employee>();
 t.ReadStruct<Employee>(1 ,11, 2, ref remp1, 2, out outMessage);

 the Api let you read a structure which is a list and a list of strcture.

 All indexes start from 1, as the C# excel API does.

The API can be expand to use more and more features and i intend to update it offten.


Functions:

ReadRowList will read an amount of values from a specific row( starting from a of course)
it will use the Range to read from memory rather then read cell by cell which is very slow compare to range read.

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


	public int ExcelColumnNameToNumber(string col_name);
	public string GetExcelColumnName(int columnNumber);
	public bool ReadColumnList(int sheetIndex,
                             int startRowIndex,
                             int startColIndex,
                             out List<object> list,
                             int rowCount,
                             out string outMessage);
	public bool ReadStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, ref T s, out string outMessage) where T : class

	public bool ReadStruct<T>(int sheetIndex, 
                                  int startRowIndex, 
                                  int startColIndex, 
                                  ref List<T> s, 
                                  int rowCount, 
                                  out string outMessage) where T : class, new()


	// Slow write 
	public bool WriteCell(int sheetIndex,
								  int rowIndex,
								  int colIndex,
								  object value,
								  bool bold,
								  Color foreColor,
								  Color backColor,
								  out string outMessage)

	public bool WriteStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, List<T> s, out string outMessage)

	public bool WriteStruct<T>(int sheetIndex, int startRowIndex, int startColIndex, T s, out string outMessage)
	public bool UpdateLastSheetName(string newName, out string outMessage)        
    public bool AddWorkSheetAtTheBegin(string name, out string outMessage)
	public bool UpdateSheetName(int index, string newName, out string outMessage)
	public bool AddWorkSheetAtTheEnd(string name, out string outMessage);
    public bool UpdateFirstSheetName(string newName, out string outMessage)
    public int SheetCount{get;}
    public string SheetName(int index)
    public void Save(string fileName = "")
	public bool NewFile(string fileName)
    public bool OpenFile(string fileName)
	public static void CloseExcel()