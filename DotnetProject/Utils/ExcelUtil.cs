using ClosedXML.Excel;
using DotnetProject.Models;

namespace DotnetProject.Utils
{
    public class ExcelUtil
    {
        public string GenerateExcelReport()
        {
            List<Person> people = InitPeople();
            // Create a new Excel package
            using (var excelPackage = new XLWorkbook())
            {
                // Add a new worksheet to the Excel package
                var worksheet = excelPackage.Worksheets.Add("Report");

                // Add data to the worksheet
                worksheet.Cell("A1").Value = "ID";
                worksheet.Cell("B1").Value = "Name";
                worksheet.Cell("C1").Value = "Age";
                worksheet.Cell("D1").Value = "BirthPlace";
                worksheet.Cell("E1").Value = "Mobile";

                // Change the background color of the header cells
                worksheet.Cell("A1").Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Cell("B1").Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Cell("C1").Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Cell("D1").Style.Fill.BackgroundColor = XLColor.LightBlue;
                worksheet.Cell("E1").Style.Fill.BackgroundColor = XLColor.LightBlue;

                for (int i = 0; i < people.Count; i++)
                {
                    worksheet.Cell($"A{i+2}").Value = people[i].Id;
                    worksheet.Cell($"B{i+2}").Value = people[i].Name;
                    worksheet.Cell($"C{i+2}").Value = people[i].Age;
                    worksheet.Cell($"D{i+2}").Value = people[i].BirthPlace;
                    worksheet.Cell($"E{i+2}").Value = people[i].Mobile;
                }

                // Adjust column width to fit content
                worksheet.Columns().AdjustToContents();

                // Save the Excel package to a file
                string filename = $"report_{DateTime.Now.ToString("yyyy-MM-dd HH-mm-ss")}.xlsx";
                var filePath = $"wwwroot/report/{filename}";
                excelPackage.SaveAs(filePath);
                return filename;
            }
        }

        public List<Person> InitPeople()
        {
            var people = new List<Person>();
            people.Add(new Person()
            {
                Name = "Morteza",
                Age = 36,
                BirthPlace = "Tehran",
                Mobile = "+989122802700",
                Id = 1,
            });

            people.Add(new Person()
            {
                Name = "Tom",
                Age = 30,
                BirthPlace = "Shiraz",
                Mobile = "+989121236565",
                Id = 1,
            });

            return people;
        }
    }
}
