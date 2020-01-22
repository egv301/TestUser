using System;
using System.Data.OleDb;
using System.Data;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using Microsoft.Office.Interop.Excel;




namespace ConsoleApplication4
{
    class Program
    {
       static void Main(string[] args)
       {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=C:\Users\User\Desktop\db\Database11.accdb; Persist Security Info=False;";
            string sql = @"select * from Users where SocialNumber = '12345543211234'";

            object id = null;
            object name = null;
            object birthDate = null;
            object PhoneNumber = null;
            object Address = null;
            object SocialNumber = null;


            using(OleDbConnection connection = new OleDbConnection(connectionString))
            {
                connection.Open();

                OleDbCommand command = new OleDbCommand(sql, connection);
                OleDbDataReader reader = command.ExecuteReader();

                if (reader.HasRows) // если есть данные
                {
                    // выводим названия столбцов
                    Console.WriteLine("{0}\t{1}\t{2}\t{3}\t{4}\t{5}", reader.GetName(0), reader.GetName(1), reader.GetName(2), reader.GetName(3), reader.GetName(4), reader.GetName(5));

                    while (reader.Read()) // построчно считываем данные
                    {
                        id = reader.GetValue(0);
                        name = reader.GetValue(1);
                        birthDate = reader.GetValue(2);
                        PhoneNumber = reader.GetValue(3);
                        Address = reader.GetValue(4);
                        SocialNumber = reader.GetValue(5);

                        Console.WriteLine("{0} \t{1} \t{2} \t{3} \t{4} \t{5}", id, name, birthDate, PhoneNumber, Address, SocialNumber);
                    }

                    
                }

                Microsoft.Office.Interop.Excel.Application excel = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook sheet = excel.Workbooks.Open(@"C:\Users\User\Documents\Visual Studio 2010\Projects\ConsoleApplication1\Test\Template\example.xlsx");
                Microsoft.Office.Interop.Excel.Worksheet x = excel.ActiveSheet as Microsoft.Office.Interop.Excel.Worksheet;

                x.Cells[3, 2] = id;
                x.Cells[4, 2] = name;
                x.Cells[5, 2] = birthDate;
                x.Cells[6, 2] = PhoneNumber;
                x.Cells[7, 2] = Address;
                x.Cells[8, 2] = SocialNumber;


                x.Cells[4, 4] = id;
                x.Cells[4, 5] = name;
                x.Cells[4, 6] = birthDate;
                x.Cells[4, 7] = PhoneNumber;
                x.Cells[4, 8] = Address;
                x.Cells[4, 9] = SocialNumber;

                sheet.SaveAs(@"C:\Users\User\Documents\Visual Studio 2010\Projects\Test\ConsoleApplication1\Result\Example.xlsx");
                sheet.Close(true, Type.Missing, Type.Missing);
                excel.Quit();

                Console.Read();
            }
       }
    }
}
