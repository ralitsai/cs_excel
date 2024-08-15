using System;
using System.IO;
using OfficeOpenXml;

namespace ConsoleApp1 {
    class Program {
        static void Main(string[] args) {
            // ako neshto ne ti trugne probvai da izteglish nakuv NuGet Package Manager, Microsoft Tools za VS (ako chetesh koda na VS), vsqkakvi c# ili .net extentioni
            // ako pak ne stane probvai da otvorish command prompt, da vuvedesh "cd " i kude ti e faila, sled tova "dotnet add package EPPlus"      

            string file_path = @"C:\Users\PC1\OneDrive\Desktop\cs_excel\ConsoleApp1\test_excel.xlsx";   // vuvedi kude ti e faila

            if (!File.Exists(file_path)) {
                Console.WriteLine("File not found.");   // proverqva dali si go vuvel pravilno
                return;
            }

            Console.Write("Enter the name of the sheet you want to read the cells from: ");
            string sheet_to_read = Console.ReadLine();   // napishi ot koi sheet shte chetesh poleto 

            Console.Write("Enter the name of the sheet you want to put the cells into: ");
            string output_sheet = Console.ReadLine();   // napishi na koi sheet da ide informaciqta

            Console.Write("Enter the cell address to read from: ");
            string cell_address_read = Console.ReadLine();    // napishi adresa na poleto koeto ti trqbva (A1, B3, D5...)

            Console.Write("Enter the cell address to write to: ");   // napishi adresa na poleto v noviq sheet kudeto shte otide informaciqta ot onova druoto
            string cell_address_write = Console.ReadLine();

            using (var package = new ExcelPackage(new FileInfo(file_path))) {
                ExcelWorksheet read_worksheet = package.Workbook.Worksheets[sheet_to_read];
                if (read_worksheet == null) {
                    Console.WriteLine("This sheet is not found.");   // proverqva dali si vuvel sushtestvuvasht sheet i printira tova ako ne si
                    return;
                }

                ExcelWorksheet output_worksheet = package.Workbook.Worksheets[output_sheet];
                if (output_worksheet == null) {
                    Console.WriteLine("This sheet is not found, but a new one will be created.");   // proverqva dali toq sheet v koito iskash da vuvedesh sushtestvuva i ako ne pravi takuv
                    output_worksheet = package.Workbook.Worksheets.Add(output_sheet);
                }

                var cell_value = read_worksheet.Cells[cell_address_read].Text;
                if (cell_value == null || string.IsNullOrEmpty(cell_value)) {
                    Console.WriteLine("The cell is empty.");   // proverqva dali poleto ne e prazno i printira tova ako e
                    return;
                } else {
                    output_worksheet.Cells[cell_address_write].Value = cell_value;
                    package.Save();   // zapazva promenite
                    Console.WriteLine($"{cell_value} written successfully.");   // printira ti kakvo ima v poleto koeto si vuvel i che se e zapisalo
                }
            }
        }
    }
}