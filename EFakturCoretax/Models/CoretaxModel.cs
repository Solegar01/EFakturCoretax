using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.Models
{
    public class CoretaxModel
    {
        public int? DocEntry { get; set; }
        public int? DocNum { get; set; }
        public int? SeriesId { get; set; }
        public string SeriesName { get; set; }
        public DateTime? DocDate { get; set; }
        public DateTime? PostDate { get; set; }

        // Mapping ke [U_T2_AR_INV], [U_T2_AR_DP], [U_T2_AR_CM]
        public bool IsARInvoice { get; set; }
        public bool IsARDownPayment { get; set; }
        public bool IsARCreditMemo { get; set; }

        // Mapping ke From-To Document Range
        public string FromDocNum { get; set; }
        public string ToDocNum { get; set; }

        public string FromDocEntry { get; set; }
        public string ToDocEntry { get; set; }

        // Mapping ke Date Range
        public DateTime? FromDate { get; set; }
        public DateTime? ToDate { get; set; }

        // Mapping ke Customer
        public string FromCust { get; set; }
        public string ToCust { get; set; }

        // Mapping ke Branch
        public string FromBranch { get; set; }
        public string ToBranch { get; set; }

        // Mapping ke Outlet
        public string FromOutlet { get; set; }
        public string ToOutlet { get; set; }

        // Status (Open/Close/Cancel)
        public string Status { get; set; } = "O";

        // Detail lines
        public List<InvoiceDataModel> Detail { get; set; } = new List<InvoiceDataModel>();
    }

    public class InvoiceDataModel
    {
        public string TIN { get; set; }
        public string DocEntry { get; set; }
        public string LineNum { get; set; }
        public string InvDate { get; set; }
        public string NoDocument { get; set; }
        public string ObjectType { get; set; }
        public string ObjectName { get; set; }
        public string BPCode { get; set; }
        public string BPName { get; set; }
        public string SellerIDTKU { get; set; }
        public string BuyerDocument { get; set; }
        public string NomorNPWP { get; set; }
        public string NPWPName { get; set; }
        public string NPWPAddress { get; set; }
        public string BuyerIDTKU { get; set; }
        public string ItemCode { get; set; }
        public string DefItemCode { get; set; }
        public string ItemName { get; set; }
        public string ItemUnit { get; set; }
        public decimal ItemPrice { get; set; }
        public decimal Qty { get; set; }
        public decimal TotalDisc { get; set; }
        public decimal TaxBase { get; set; }
        public decimal OtherTaxBase { get; set; }
        public decimal VATRate { get; set; }
        public decimal AmountVAT { get; set; }
        public decimal STLGRate { get; set; }
        public decimal STLG { get; set; }
        public decimal CoretaxVatAmount { get; set; }
        public decimal CoretaxVatRate { get; set; }
        public string JenisPajak { get; set; }
        public string KetTambahan { get; set; }
        public string PajakPengganti { get; set; }
        public string Referensi { get; set; }
        public string Status { get; set; }
        public string KodeDokumenPendukung { get; set; }
        public string BranchCode { get; set; }
        public string BranchName { get; set; }
        public string AddInfo { get; set; }
        public string BuyerCountry { get; set; }
        public string BuyerEmail { get; set; }
        public string OutletCode { get; set; }
        public string OutletName { get; set; }
        public string Revise { get; set; }
    }
}
