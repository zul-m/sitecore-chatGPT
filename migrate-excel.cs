using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Sitecore.Data;
using Sitecore.Data.Items;
using Sitecore.Data.Fields;

namespace ExcelToSitecoreMigration
{
    class Program
    {
        static void Main(string[] args)
        {
            // Open the Excel file
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Open("C:\\ExcelData.xlsx");
            Worksheet worksheet = workbook.Worksheets[1];

            // Get the data range from the Excel file
            Range excelRange = worksheet.UsedRange;

            // Loop through each row of the data
            for (int row = 1; row <= excelRange.Rows.Count; row++)
            {
                // Get the Sitecore item where the data will be imported
                Item parentItem = Sitecore.Context.Database.GetItem(new ID("{GUID of parent item}"));

                // Create a new item under the parent item
                Item newItem = parentItem.Add("Item Name", new TemplateID(new ID("{GUID of template}")));

                // Loop through each column of the data row
                for (int col = 1; col <= excelRange.Columns.Count; col++)
                {
                    // Get the field name from the first row of the Excel file
                    string fieldName = excelRange.Cells[1, col].Value2.ToString();

                    // Get the field value from the current row of the Excel file
                    string fieldValue = excelRange.Cells[row, col].Value2.ToString();

                    // Set the value of the field on the Sitecore item
                    newItem.Editing.BeginEdit();
                    newItem[fieldName] = fieldValue;
                    newItem.Editing.EndEdit();
                }
            }

            // Close the Excel file
            workbook.Close();
            excel.Quit();
        }
    }
}