using ExcelLib;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelTestCase
{
    class Program
    {
        class Employee
        {

            public string Name { get; set; }
            public string Number { get; set; }
            public string Employee_ID { get; set; }
            public string Email_ID { get; set; }
        }

        static void Main(string[] args)
        {
            string outMessage;
            ExcelApi t = new ExcelApi();
            string fileName = "c:\\MyExcelFile.xlsx";
            if (t.NewFile(fileName) == true)
            {
                Console.WriteLine("File Created");
                t.UpdateSheetName(0, "Eli Arad 1", out outMessage);
                for (int i = 1; i < 5; i++)
                {
                    if (t.AddWorkSheetAtTheEnd("Eli Arad " + (i + 1), out outMessage) == false)
                    {

                    }
                }
                //t.UpdateSheetName(0, "7777", out outMessage);
            }
            else
            {
                t.OpenFile(fileName);
                /*
                if (t.AddWorkSheetAtTheEnd("gggggggg", out outMessage) == false)
                {

                }
                
                if (t.UpdateLastSheetName("121212", out outMessage) == false)
                {

                }
                */
                if (t.WriteCell(0, 1, 1, "Hello world", out outMessage) == false)
                {
                    Console.WriteLine("error: " + outMessage);
                }
                if (t.WriteCell(1, 1, 1, "Hello world", out outMessage) == false)
                {
                    Console.WriteLine("error: " + outMessage);
                }

                if (t.WriteCell(2, 1, 1, "Hello world", true, Color.Red , Color.Transparent, out outMessage) == false)
                {
                    Console.WriteLine("error: " + outMessage);
                }

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



                if (t.WriteStruct<Employee>(0, 5, 2, emp, out outMessage) == false)
                {

                }
                Employee remp = new Employee();
                t.ReadStruct<Employee>(0, 6, 2, ref remp, out outMessage);

                
                if (t.WriteStruct<Employee>(0, 10, 2, employees, out outMessage) == false)
                {

                }

                List<Employee> remp1 = new List<Employee>();
                t.ReadStruct<Employee>(0, 11, 2, ref remp1, 2, out outMessage);

                List<object> data1 = new List<object>();

                data1.Add("Eli");
                data1.Add("1");
                data1.Add("0.2323");
                data1.Add("Arad");
                data1.Add("12112");            
                t.WriteLine(0, 20, 1, data1, out outMessage);

            }
            t.Close();
        }

    }
}
