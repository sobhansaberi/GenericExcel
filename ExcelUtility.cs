using ClosedXML.Excel;
using System.ComponentModel;

namespace Application.Common.Utilities;
public class ExcelUtility
{
    public static string GetExcelFile<T>(IReadOnlyList<T> dataList, string fileName) where T : class
    {
        using (var workbook = new XLWorkbook())
        {
            // Add a worksheet to the workbook
            var worksheet = workbook.Worksheets.Add("Sheet1");

            var type = typeof(T);
            var properties = type.GetProperties();

            // Define headers
            int index = 1;
            foreach (var property in properties)
            {
                var name = string.Empty;
                var displayNameAtt = (DisplayNameAttribute?)property.GetCustomAttributes(typeof(DisplayNameAttribute), true).FirstOrDefault();

                if (displayNameAtt != null)
                    name = displayNameAtt.DisplayName;
                else
                    name = property.Name;

                worksheet.Cell(1, index).Value = name; index++;
            }

            // Populate data from the list
            int row = 2; // Start from row 2 to leave space for headers

            foreach (var data in dataList)
            {
                int column = 1;
                foreach (var property in properties)
                {
                    var item = property.GetValue(data);
                    if (item != null) worksheet.Cell(row, column).Value = item.ToString();

                    column++;
                }
                row++;
            }

            // Save the workbook to a file
            fileName = $"{fileName}.xlsx";
            string folderPath = $"wwwroot/Reports/{fileName}";
            workbook.SaveAs(folderPath);

            return fileName;
        }
    }
}
