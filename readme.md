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


