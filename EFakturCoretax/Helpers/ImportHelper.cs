using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using EFakturCoretax.Models;
using SAPbouiCOM.Framework;

namespace EFakturCoretax.Helpers
{
    public static class ImportHelper
    {
        public static string GetPathFile()
        {
            string res = string.Empty;

            try
            {
                Thread t = new Thread(() =>
                {
                    using (var dummyForm = new System.Windows.Forms.Form
                    {
                        TopMost = true,
                        ShowInTaskbar = false,
                        WindowState = System.Windows.Forms.FormWindowState.Minimized
                    })
                    using (var dialog = new System.Windows.Forms.OpenFileDialog
                    {
                        Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*",
                        Title = "Select Excel file"
                    })
                    {
                        dummyForm.Show();
                        dummyForm.Hide();

                        if (dialog.ShowDialog(dummyForm) == System.Windows.Forms.DialogResult.OK)
                        {
                            res = dialog.FileName.Trim('"').Trim();
                        }
                    }
                });

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();
                return res;
            }
            catch (Exception ex)
            {
                throw new Exception( "Error selecting file: " + ex.Message);
            }
        }


        public static List<InvoiceModel> ImportInvoices(string filePath)
        {
            var list = new List<InvoiceModel>();

            using (var workbook = new XLWorkbook(filePath))
            {
                var worksheet = workbook.Worksheets.Worksheet(1); // First sheet
                var rows = worksheet.RangeUsed().RowsUsed().Skip(1); // Skip header

                foreach (var row in rows)
                {
                    var model = new InvoiceModel
                    {
                        InvoiceNo = row.Cell(1).GetValue<string>().Trim(),
                        CoretaxNo = row.Cell(2).GetValue<string>().Trim()
                    };

                    if (!string.IsNullOrEmpty(model.InvoiceNo) || !string.IsNullOrEmpty(model.CoretaxNo))
                        list.Add(model);
                }
            }

            return list;
        }
    }
}
