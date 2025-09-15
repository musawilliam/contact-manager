using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ContactManager
{
    public class ExcelService
    {
        public List<Contact> ImportFromExcel(string filePath)
        {
            var contacts = new List<Contact>();
            
            try
            {
                using (var workbook = new XLWorkbook(filePath))
                {
                    var worksheet = workbook.Worksheet(1); // First worksheet
                    var rows = worksheet.RowsUsed();
                    
                    // Skip header row, start from row 2
                    foreach (var row in rows.Skip(1))
                    {
                        var contact = new Contact
                        {
                            Name = GetCellValue(row.Cell(1)),
                            Surname = GetCellValue(row.Cell(2)),
                            Number = GetCellValue(row.Cell(3)),
                            Used = false
                        };
                        
                        if (!string.IsNullOrEmpty(contact.Name))
                        {
                            contacts.Add(contact);
                        }
                    }
                }
                
                MessageBox.Show($"Successfully imported {contacts.Count} contacts.", 
                    "Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error importing Excel file: {ex.Message}", "Import Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            
            return contacts;
        }

        public void ExportToExcel(List<Contact> contacts, string filePath)
        {
            try
            {
                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Contacts");
                    
                    // Add headers
                    worksheet.Cell("A1").Value = "Name";
                    worksheet.Cell("B1").Value = "Surname";
                    worksheet.Cell("C1").Value = "Number";
                    worksheet.Cell("E1").Value = "Used";
                    worksheet.Cell("F1").Value = "Created Date";
                    
                    // Format headers
                    var headerRange = worksheet.Range("A1:F1");
                    headerRange.Style.Font.Bold = true;
                    headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
                    
                    // Add data
                    for (int i = 0; i < contacts.Count; i++)
                    {
                        int row = i + 2;
                        worksheet.Cell(row, 1).Value = contacts[i].Name;
                        worksheet.Cell(row, 2).Value = contacts[i].Surname;
                        worksheet.Cell(row, 3).Value = contacts[i].Number;
                        worksheet.Cell(row, 5).Value = contacts[i].Used ? "Yes" : "No";
                        worksheet.Cell(row, 6).Value = contacts[i].CreatedDate.ToString("yyyy-MM-dd HH:mm:ss");
                    }
                    
                    // Auto-fit columns
                    worksheet.Columns().AdjustToContents();
                    
                    workbook.SaveAs(filePath);
                }
                
                MessageBox.Show($"Data exported successfully to {filePath}", "Export Complete", 
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error exporting to Excel: {ex.Message}", "Export Error", 
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string GetCellValue(IXLCell cell)
        {
            return cell.GetValue<string>() ?? "";
        }
    }
}