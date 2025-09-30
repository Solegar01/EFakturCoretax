using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Xml.Serialization;

namespace EFakturCoretax.Helpers
{
    public static class ExportHelper
    {
        private static string _lastDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private class WindowWrapper : IWin32Window
        {
            private readonly IntPtr _hwnd;
            public WindowWrapper(IntPtr handle) { _hwnd = handle; }
            public IntPtr Handle => _hwnd;
        }

        // ==================== Export XML ====================
        public static bool ExportXml(TaxInvoiceBulk invoice)
        {
            try
            {
                if (invoice == null) throw new ArgumentNullException(nameof(invoice));
                
                string filePath = GetSaveFilePath("XML Files (*.xml)|*.xml", "TaxInvoice.xml");
                if (string.IsNullOrEmpty(filePath)) return false;

                XmlSerializer serializer = new XmlSerializer(typeof(TaxInvoiceBulk));
                using (var writer = new StreamWriter(filePath))
                {
                    serializer.Serialize(writer, invoice);
                }
                return true;
            }
            catch (Exception)
            {

                throw;
            }
        }

        // ==================== Export CSV ====================
        public static bool ExportCsv(TaxInvoiceBulk invoice)
        {
            try
            {
                if (invoice == null) throw new ArgumentNullException(nameof(invoice));

                string filePath = GetSaveFilePath("CSV Files (*.csv)|*.csv", "TaxInvoice.csv");
                if (string.IsNullOrEmpty(filePath)) return false;

                var sb = new StringBuilder();

                // CSV Header
                sb.AppendLine("TaxInvoiceDate,TaxInvoiceOpt,TrxCode,AddInfo,CustomDoc,CustomDocMonthYear,RefDesc,FacilityStamp,SellerIDTKU,BuyerTin,BuyerDocument,BuyerCountry,BuyerDocumentNumber,BuyerName,BuyerAdress,BuyerEmail,BuyerIDTKU,GoodServiceOpt,GoodServiceCode,GoodServiceName,Unit,Price,Qty,TotalDiscount,TaxBase,OtherTaxBase,VATRate,VAT,STLGRate,STLG");

                foreach (var taxInvoice in invoice.ListOfTaxInvoice.TaxInvoiceCollection)
                {
                    foreach (var gs in taxInvoice.ListOfGoodService.GoodServiceCollection)
                    {
                        sb.AppendLine(string.Join(",", new[]
                        {
                CsvEscape(taxInvoice.TaxInvoiceDate),
                CsvEscape(taxInvoice.TaxInvoiceOpt),
                CsvEscape(taxInvoice.TrxCode),
                CsvEscape(taxInvoice.AddInfo),
                CsvEscape(taxInvoice.CustomDoc),
                CsvEscape(taxInvoice.CustomDocMonthYear),
                CsvEscape(taxInvoice.RefDesc),
                CsvEscape(taxInvoice.FacilityStamp),
                CsvEscape(taxInvoice.SellerIDTKU),
                CsvEscape(taxInvoice.BuyerTin),
                CsvEscape(taxInvoice.BuyerDocument),
                CsvEscape(taxInvoice.BuyerCountry),
                CsvEscape(taxInvoice.BuyerDocumentNumber),
                CsvEscape(taxInvoice.BuyerName),
                CsvEscape(taxInvoice.BuyerAdress),
                CsvEscape(taxInvoice.BuyerEmail),
                CsvEscape(taxInvoice.BuyerIDTKU),
                CsvEscape(gs.Opt),
                CsvEscape(gs.Code),
                CsvEscape(gs.Name),
                CsvEscape(gs.Unit),
                gs.Price.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.Qty.ToString(),
                gs.TotalDiscount.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.TaxBase.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.OtherTaxBase.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.VATRate.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.VAT.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.STLGRate.ToString(System.Globalization.CultureInfo.InvariantCulture),
                gs.STLG.ToString(System.Globalization.CultureInfo.InvariantCulture)
            }));
                    }
                }

                // Write with UTF-8 BOM for Excel compatibility
                File.WriteAllText(filePath, sb.ToString(), new UTF8Encoding(true));
                return true;
            }
            catch (Exception)
            {

                throw;
            }
        }

        /// <summary>
        /// Escape CSV values (quotes, commas, line breaks)
        /// </summary>
        private static string CsvEscape(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Contains(",") || value.Contains("\"") || value.Contains("\n"))
            {
                return "\"" + value.Replace("\"", "\"\"") + "\"";
            }
            return value;
        }


        // ==================== Helper: SaveFileDialog in STA thread ====================
        public static string GetSaveFilePath(string filter, string defaultFileName)
        {
            try
            {
                string filePath = null;

                string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                string fileNameWithTimestamp =
                    Path.GetFileNameWithoutExtension(defaultFileName) + "_" +
                    timestamp +
                    Path.GetExtension(defaultFileName);

                var t = new Thread(() =>
                {
                    string initialDir = Directory.Exists(_lastDirectory)
                    ? _lastDirectory
                    : Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                    using (var saveDialog = new SaveFileDialog
                    {
                        Filter = filter,
                        Title = "Save File",
                        FileName = fileNameWithTimestamp,
                        RestoreDirectory = true,
                        CheckPathExists = true,
                        InitialDirectory = initialDir
                    })
                    {
                        var owner = new WindowWrapper(GetForegroundWindow());

                        if (saveDialog.ShowDialog(owner) == DialogResult.OK)
                        {
                            filePath = saveDialog.FileName;
                            // update last directory
                            _lastDirectory = Path.GetDirectoryName(filePath);
                        }
                    }
                });

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();

                return filePath;
            }
            catch (Exception)
            {

                throw;
            }
        }
    }
}
