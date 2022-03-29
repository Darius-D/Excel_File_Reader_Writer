using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading.Tasks;

namespace ExcelWriter
{
    class Program
    {
        static async Task Main(string[] args)
        {
            //required for the nugget package to work
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // Tells the path of the file and name of file.
            var file = new FileInfo(@"C:\Users\ddubo\Desktop\FHL\Demos\ExcelWriterDemo.xlsx");

            var people = GetSetupData();

            //Saves data async to the file. 
            await SaveExcelFile(people, file);

            List<PersonModel> peopleFromExcel = await LoadExcelFile(file);

            foreach(var p in peopleFromExcel)
            {
                Console.WriteLine($"ID:{p.Id} \n First Name:{p.FirstName} \n Last Name:{p.LastName} \n");
            }
        }
        //Pulling data from an excel file.
        private static async Task<List<PersonModel>> LoadExcelFile(FileInfo file)
        {
            List<PersonModel> output = new();
            // would normally make a try catch block to makesure file exists.
            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var worksheet = package.Workbook.Worksheets[0];

            int row = 3;
            int col = 1;

            while(string.IsNullOrWhiteSpace(worksheet.Cells[row,col].Value?.ToString()) == false)
            {
                PersonModel p = new();

                p.Id = int.Parse(worksheet.Cells[row, col].Value.ToString());
                p.FirstName = worksheet.Cells[row, col + 1].Value.ToString();
                p.LastName = worksheet.Cells[row, col + 2].Value.ToString();
                output.Add(p);
                row += 1;
            }

            return output;

        }

        private static async Task SaveExcelFile(List<PersonModel> people, FileInfo file)
        {
            DeleteIfExists(file);

            //using statement ensures the resources are disposed at the end of this method.
            using(var package = new ExcelPackage(file))// this opens the excel file.
            {
                //This adds a new worksheet. Can not add new sheet if it already exists.
                var worksheet = package.Workbook.Worksheets.Add("MainReport");
                // This selects the rows and says "starting at A1 insert all elements in the collection. 
                //Th collection in this example is people. the second argument is true meaning the property names
                //will be added as a header before the data to label the columns. 
                var rangeCells = worksheet.Cells["A2"].LoadFromCollection(people, true);

                
                rangeCells.AutoFitColumns();
                //formats header
                worksheet.Cells["A1"].Value = "Our Cool Report";
                worksheet.Cells["A1:C1"].Merge = true;
                worksheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Row(1).Style.Font.Size = 24;
                worksheet.Row(1).Style.Font.Color.SetColor(Color.Gray);

                worksheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                worksheet.Rows.Style.Font.Bold = true;
                worksheet.Column(3).Width = 20;



                // Save file before disposing.
                await package.SaveAsync();
            }
        }

        private static void DeleteIfExists(FileInfo file)
        {
            if (file.Exists) file.Delete();
        }

        private static List<PersonModel> GetSetupData()
        {
            List<PersonModel> output = new()
            {
                new() { Id = 1, FirstName = "Darius", LastName = "Dubose" },
                new() { Id = 2, FirstName = "John", LastName = "Epes" },
                new() { Id = 3, FirstName = "Ben", LastName = "Daley" }
            };
            return output;
        }


    }
}
