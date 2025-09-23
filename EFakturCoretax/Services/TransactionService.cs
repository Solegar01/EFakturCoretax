using EFakturCoretax.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.Services
{
    public static class TransactionService
    {
        public static Task<int> AddDataCoretax(CoretaxModel model)
        {
            return Task.Run(() => {
                SAPbobsCOM.Company oCompany = null;
                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                try
                {

                    oCompany = CompanyService.GetCompany();
                    oCompanyService = oCompany.GetCompanyService();
                    int seriesId = 0;
                    seriesId = GetSeriesIdCoretax();
                    // 1. Get GeneralService for UDO
                    oGeneralService = oCompanyService.GetGeneralService("T2_CORETAX");

                    // 2. Create header
                    oGeneralData = (GeneralData)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralData);

                    // --- Header fields ---
                    oGeneralData.SetProperty("Series", seriesId);
                    oGeneralData.SetProperty("U_T2_Posting_Date", DateTime.Now);
                    oGeneralData.SetProperty("U_T2_AR_INV", model.IsARInvoice ? "Y" : "N");
                    oGeneralData.SetProperty("U_T2_AR_DP", model.IsARDownPayment ? "Y" : "N");
                    oGeneralData.SetProperty("U_T2_AR_CM", model.IsARCreditMemo ? "Y" : "N");

                    if (model.FromDate != null) oGeneralData.SetProperty("U_T2_From_Date", model.FromDate);
                    if (model.ToDate != null) oGeneralData.SetProperty("U_T2_To_Date", model.ToDate);
                    if (!string.IsNullOrEmpty(model.FromDocNum)) oGeneralData.SetProperty("U_T2_From_Doc", model.FromDocNum);
                    if (!string.IsNullOrEmpty(model.ToDocNum)) oGeneralData.SetProperty("U_T2_To_Doc", model.ToDocNum);
                    if (!string.IsNullOrEmpty(model.FromDocEntry)) oGeneralData.SetProperty("U_T2_From_Doc_Entry", model.FromDocEntry);
                    if (!string.IsNullOrEmpty(model.ToDocEntry)) oGeneralData.SetProperty("U_T2_To_Doc_Entry", model.ToDocEntry);

                    if (!string.IsNullOrEmpty(model.FromCust)) oGeneralData.SetProperty("U_T2_From_Cust", model.FromCust);
                    if (!string.IsNullOrEmpty(model.ToCust)) oGeneralData.SetProperty("U_T2_To_Cust", model.ToCust);
                    if (!string.IsNullOrEmpty(model.FromBranch)) oGeneralData.SetProperty("U_T2_From_Branch", model.FromBranch);
                    if (!string.IsNullOrEmpty(model.ToBranch)) oGeneralData.SetProperty("U_T2_To_Branch", model.ToBranch);
                    if (!string.IsNullOrEmpty(model.FromOutlet)) oGeneralData.SetProperty("U_T2_From_Outlet", model.FromOutlet);
                    if (!string.IsNullOrEmpty(model.ToOutlet)) oGeneralData.SetProperty("U_T2_To_Outlet", model.ToOutlet);

                    // 3. Add detail lines
                    GeneralDataCollection oChildren = oGeneralData.Child("T2_CORETAX_DT");
                    foreach (var line in model.Detail)
                    {
                        GeneralData oChild = oChildren.Add();

                        if (!string.IsNullOrEmpty(line.TIN)) oChild.SetProperty("U_T2_TIN", line.TIN);
                        if (!string.IsNullOrEmpty(line.DocEntry)) oChild.SetProperty("U_T2_DocEntry", line.DocEntry);
                        if (!string.IsNullOrEmpty(line.LineNum)) oChild.SetProperty("U_T2_LineNum", line.LineNum);
                        if (!string.IsNullOrEmpty(line.InvDate)) oChild.SetProperty("U_T2_Inv_Date", line.InvDate);
                        if (!string.IsNullOrEmpty(line.NoDocument)) oChild.SetProperty("U_T2_No_Doc", line.NoDocument);
                        if (!string.IsNullOrEmpty(line.ObjectType)) oChild.SetProperty("U_T2_Object_Type", line.ObjectType);
                        if (!string.IsNullOrEmpty(line.ObjectName)) oChild.SetProperty("U_T2_Object_Name", line.ObjectName);
                        if (!string.IsNullOrEmpty(line.BPCode)) oChild.SetProperty("U_T2_BP_Code", line.BPCode);
                        if (!string.IsNullOrEmpty(line.BPName)) oChild.SetProperty("U_T2_BP_Name", line.BPName);
                        if (!string.IsNullOrEmpty(line.SellerIDTKU)) oChild.SetProperty("U_T2_Seller_IDTKU", line.SellerIDTKU);
                        if (!string.IsNullOrEmpty(line.BuyerDocument)) oChild.SetProperty("U_T2_Buyer_Doc", line.BuyerDocument);
                        if (!string.IsNullOrEmpty(line.NomorNPWP)) oChild.SetProperty("U_T2_Nomor_NPWP", line.NomorNPWP);
                        if (!string.IsNullOrEmpty(line.NPWPName)) oChild.SetProperty("U_T2_NPWP_Name", line.NPWPName);
                        if (!string.IsNullOrEmpty(line.NPWPAddress)) oChild.SetProperty("U_T2_NPWP_Address", line.NPWPAddress);
                        if (!string.IsNullOrEmpty(line.BuyerIDTKU)) oChild.SetProperty("U_T2_Buyer_IDTKU", line.BuyerIDTKU);
                        if (!string.IsNullOrEmpty(line.ItemCode)) oChild.SetProperty("U_T2_Item_Code", line.ItemCode);
                        if (!string.IsNullOrEmpty(line.ItemName)) oChild.SetProperty("U_T2_Item_Name", line.ItemName);
                        if (!string.IsNullOrEmpty(line.ItemUnit)) oChild.SetProperty("U_T2_Item_Unit", line.ItemUnit);

                        // --- numeric fields (pastikan UDF tipe number/price/rate) ---
                        if (line.ItemPrice > 0) oChild.SetProperty("U_T2_Item_Price", line.ItemPrice);
                        if (line.Qty > 0) oChild.SetProperty("U_T2_Qty", line.Qty);
                        if (line.TotalDisc > 0) oChild.SetProperty("U_T2_Total_Disc", line.TotalDisc);
                        if (line.TaxBase > 0) oChild.SetProperty("U_T2_Tax_Base", line.TaxBase);
                        if (line.OtherTaxBase > 0) oChild.SetProperty("U_T2_Other_Tax_Base", line.OtherTaxBase);
                        if (line.VATRate > 0) oChild.SetProperty("U_T2_VAT_Rate", line.VATRate);
                        if (line.AmountVAT > 0) oChild.SetProperty("U_T2_Amount_VAT", line.AmountVAT);
                        if (line.STLGRate > 0) oChild.SetProperty("U_T2_STLG_Rate", line.STLGRate);
                        if (line.STLG > 0) oChild.SetProperty("U_T2_STLG", line.STLG);

                        // --- string fields ---
                        if (!string.IsNullOrEmpty(line.JenisPajak)) oChild.SetProperty("U_T2_Jenis_Pajak", line.JenisPajak);
                        if (!string.IsNullOrEmpty(line.KetTambahan)) oChild.SetProperty("U_T2_Ket_Tambahan", line.KetTambahan);
                        if (!string.IsNullOrEmpty(line.PajakPengganti)) oChild.SetProperty("U_T2_Pajak_Pengganti", line.PajakPengganti);
                        if (!string.IsNullOrEmpty(line.Referensi)) oChild.SetProperty("U_T2_Referensi", line.Referensi);
                        if (!string.IsNullOrEmpty(line.Status)) oChild.SetProperty("U_T2_Status", line.Status);
                        if (!string.IsNullOrEmpty(line.KodeDokumenPendukung)) oChild.SetProperty("U_T2_Kode_Dok_Pendukung", line.KodeDokumenPendukung);
                        if (!string.IsNullOrEmpty(line.BranchCode)) oChild.SetProperty("U_T2_Branch_Code", line.BranchCode);
                        if (!string.IsNullOrEmpty(line.BranchName)) oChild.SetProperty("U_T2_Branch_Name", line.BranchName);
                        if (!string.IsNullOrEmpty(line.OutletCode)) oChild.SetProperty("U_T2_Outlet_Code", line.OutletCode);
                        if (!string.IsNullOrEmpty(line.OutletName)) oChild.SetProperty("U_T2_Outlet_Name", line.OutletName);
                        if (!string.IsNullOrEmpty(line.AddInfo)) oChild.SetProperty("U_T2_Add_Info", line.AddInfo);
                        if (!string.IsNullOrEmpty(line.BuyerCountry)) oChild.SetProperty("U_T2_Buyer_Country", line.BuyerCountry);
                        if (!string.IsNullOrEmpty(line.BuyerEmail)) oChild.SetProperty("U_T2_Buyer_Email", line.BuyerEmail);
                    }

                    // 4. Save data
                    oGeneralParams = oGeneralService.Add(oGeneralData);
                    int newDocEntry = Convert.ToInt32(oGeneralParams.GetProperty("DocEntry"));

                    return Task.FromResult(newDocEntry);
                }
                catch (Exception ex)
                {
                    // kalau mau return false aja biar lebih aman
                    throw new Exception($"Error AddDataCoretax: {ex.Message}", ex);
                }
                finally
                {
                    if (oGeneralData != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                    if (oGeneralParams != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralParams);
                    if (oGeneralService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    if (oCompanyService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompanyService);
                }
            });
        }

        public static Task<int> UpdateDataCoretax(CoretaxModel model)
        {
            return Task.Run(() => {
                SAPbobsCOM.Company oCompany = null;
                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.GeneralData oGeneralData = null;
                try
                {
                    oCompany = CompanyService.GetCompany();
                    oCompanyService = oCompany.GetCompanyService();

                    // 1. Get GeneralService for UDO
                    oGeneralService = oCompanyService.GetGeneralService("T2_CORETAX");

                    // 2. Load existing data by DocEntry
                    oGeneralParams = (GeneralDataParams)oGeneralService.GetDataInterface(GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", model.DocEntry);   // pastikan model.DocEntry sudah ada

                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);

                    // --- Header fields ---
                    oGeneralData.SetProperty("U_T2_AR_INV", model.IsARInvoice ? "Y" : "N");
                    oGeneralData.SetProperty("U_T2_AR_DP", model.IsARDownPayment ? "Y" : "N");
                    oGeneralData.SetProperty("U_T2_AR_CM", model.IsARCreditMemo ? "Y" : "N");

                    if (model.FromDate != null) oGeneralData.SetProperty("U_T2_From_Date", model.FromDate);
                    if (model.ToDate != null) oGeneralData.SetProperty("U_T2_To_Date", model.ToDate);
                    if (!string.IsNullOrEmpty(model.FromDocNum)) oGeneralData.SetProperty("U_T2_From_Doc", model.FromDocNum);
                    if (!string.IsNullOrEmpty(model.ToDocNum)) oGeneralData.SetProperty("U_T2_To_Doc", model.ToDocNum);
                    if (!string.IsNullOrEmpty(model.FromDocEntry)) oGeneralData.SetProperty("U_T2_From_Doc_Entry", model.FromDocEntry);
                    if (!string.IsNullOrEmpty(model.ToDocEntry)) oGeneralData.SetProperty("U_T2_To_Doc_Entry", model.ToDocEntry);

                    if (!string.IsNullOrEmpty(model.FromCust)) oGeneralData.SetProperty("U_T2_From_Cust", model.FromCust);
                    if (!string.IsNullOrEmpty(model.ToCust)) oGeneralData.SetProperty("U_T2_To_Cust", model.ToCust);
                    if (!string.IsNullOrEmpty(model.FromBranch)) oGeneralData.SetProperty("U_T2_From_Branch", model.FromBranch);
                    if (!string.IsNullOrEmpty(model.ToBranch)) oGeneralData.SetProperty("U_T2_To_Branch", model.ToBranch);
                    if (!string.IsNullOrEmpty(model.FromOutlet)) oGeneralData.SetProperty("U_T2_From_Outlet", model.FromOutlet);
                    if (!string.IsNullOrEmpty(model.ToOutlet)) oGeneralData.SetProperty("U_T2_To_Outlet", model.ToOutlet);

                    // 3. Add detail lines
                    GeneralDataCollection oChildren = oGeneralData.Child("T2_CORETAX_DT");

                    while (oChildren.Count > 0)
                    {
                        oChildren.Remove(0);
                    }

                    foreach (var line in model.Detail)
                    {
                        GeneralData oChild = oChildren.Add();

                        if (!string.IsNullOrEmpty(line.TIN)) oChild.SetProperty("U_T2_TIN", line.TIN);
                        if (!string.IsNullOrEmpty(line.DocEntry)) oChild.SetProperty("U_T2_DocEntry", line.DocEntry);
                        if (!string.IsNullOrEmpty(line.LineNum)) oChild.SetProperty("U_T2_LineNum", line.LineNum);
                        if (!string.IsNullOrEmpty(line.InvDate)) oChild.SetProperty("U_T2_Inv_Date", line.InvDate);
                        if (!string.IsNullOrEmpty(line.NoDocument)) oChild.SetProperty("U_T2_No_Doc", line.NoDocument);
                        if (!string.IsNullOrEmpty(line.ObjectType)) oChild.SetProperty("U_T2_Object_Type", line.ObjectType);
                        if (!string.IsNullOrEmpty(line.ObjectName)) oChild.SetProperty("U_T2_Object_Name", line.ObjectName);
                        if (!string.IsNullOrEmpty(line.BPCode)) oChild.SetProperty("U_T2_BP_Code", line.BPCode);
                        if (!string.IsNullOrEmpty(line.BPName)) oChild.SetProperty("U_T2_BP_Name", line.BPName);
                        if (!string.IsNullOrEmpty(line.SellerIDTKU)) oChild.SetProperty("U_T2_Seller_IDTKU", line.SellerIDTKU);
                        if (!string.IsNullOrEmpty(line.BuyerDocument)) oChild.SetProperty("U_T2_Buyer_Doc", line.BuyerDocument);
                        if (!string.IsNullOrEmpty(line.NomorNPWP)) oChild.SetProperty("U_T2_Nomor_NPWP", line.NomorNPWP);
                        if (!string.IsNullOrEmpty(line.NPWPName)) oChild.SetProperty("U_T2_NPWP_Name", line.NPWPName);
                        if (!string.IsNullOrEmpty(line.NPWPAddress)) oChild.SetProperty("U_T2_NPWP_Address", line.NPWPAddress);
                        if (!string.IsNullOrEmpty(line.BuyerIDTKU)) oChild.SetProperty("U_T2_Buyer_IDTKU", line.BuyerIDTKU);
                        if (!string.IsNullOrEmpty(line.ItemCode)) oChild.SetProperty("U_T2_Item_Code", line.ItemCode);
                        if (!string.IsNullOrEmpty(line.ItemName)) oChild.SetProperty("U_T2_Item_Name", line.ItemName);
                        if (!string.IsNullOrEmpty(line.ItemUnit)) oChild.SetProperty("U_T2_Item_Unit", line.ItemUnit);

                        // --- numeric fields (pastikan UDF tipe number/price/rate) ---
                        if (line.ItemPrice > 0) oChild.SetProperty("U_T2_Item_Price", line.ItemPrice);
                        if (line.Qty > 0) oChild.SetProperty("U_T2_Qty", line.Qty);
                        if (line.TotalDisc > 0) oChild.SetProperty("U_T2_Total_Disc", line.TotalDisc);
                        if (line.TaxBase > 0) oChild.SetProperty("U_T2_Tax_Base", line.TaxBase);
                        if (line.OtherTaxBase > 0) oChild.SetProperty("U_T2_Other_Tax_Base", line.OtherTaxBase);
                        if (line.VATRate > 0) oChild.SetProperty("U_T2_VAT_Rate", line.VATRate);
                        if (line.AmountVAT > 0) oChild.SetProperty("U_T2_Amount_VAT", line.AmountVAT);
                        if (line.STLGRate > 0) oChild.SetProperty("U_T2_STLG_Rate", line.STLGRate);
                        if (line.STLG > 0) oChild.SetProperty("U_T2_STLG", line.STLG);

                        // --- string fields ---
                        if (!string.IsNullOrEmpty(line.JenisPajak)) oChild.SetProperty("U_T2_Jenis_Pajak", line.JenisPajak);
                        if (!string.IsNullOrEmpty(line.KetTambahan)) oChild.SetProperty("U_T2_Ket_Tambahan", line.KetTambahan);
                        if (!string.IsNullOrEmpty(line.PajakPengganti)) oChild.SetProperty("U_T2_Pajak_Pengganti", line.PajakPengganti);
                        if (!string.IsNullOrEmpty(line.Referensi)) oChild.SetProperty("U_T2_Referensi", line.Referensi);
                        if (!string.IsNullOrEmpty(line.Status)) oChild.SetProperty("U_T2_Status", line.Status);
                        if (!string.IsNullOrEmpty(line.KodeDokumenPendukung)) oChild.SetProperty("U_T2_Kode_Dok_Pendukung", line.KodeDokumenPendukung);
                        if (!string.IsNullOrEmpty(line.BranchCode)) oChild.SetProperty("U_T2_Branch_Code", line.BranchCode);
                        if (!string.IsNullOrEmpty(line.BranchName)) oChild.SetProperty("U_T2_Branch_Name", line.BranchName);
                        if (!string.IsNullOrEmpty(line.OutletCode)) oChild.SetProperty("U_T2_Outlet_Code", line.OutletCode);
                        if (!string.IsNullOrEmpty(line.OutletName)) oChild.SetProperty("U_T2_Outlet_Name", line.OutletName);
                        if (!string.IsNullOrEmpty(line.AddInfo)) oChild.SetProperty("U_T2_Add_Info", line.AddInfo);
                        if (!string.IsNullOrEmpty(line.BuyerCountry)) oChild.SetProperty("U_T2_Buyer_Country", line.BuyerCountry);
                        if (!string.IsNullOrEmpty(line.BuyerEmail)) oChild.SetProperty("U_T2_Buyer_Email", line.BuyerEmail);
                    }
                    // 5. Save update
                    oGeneralService.Update(oGeneralData);

                    return Task.FromResult((int)oGeneralData.GetProperty("DocEntry"));
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error UpdateDataCoretax: {ex.Message}", ex);
                }
                finally
                {
                    if (oGeneralData != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                    if (oGeneralParams != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralParams);
                    if (oGeneralService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    if (oCompanyService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompanyService);
                }
            });
        }

        public static Task CloseCoretax(int docEntry)
        {
            return Task.Run(() =>
            {
                SAPbobsCOM.Company oCompany = null;
                SAPbobsCOM.CompanyService oCompanyService = null;
                SAPbobsCOM.GeneralService oGeneralService = null;
                SAPbobsCOM.GeneralDataParams oGeneralParams = null;
                SAPbobsCOM.GeneralData oGeneralData = null;

                try
                {
                    oCompany = CompanyService.GetCompany();
                    oCompanyService = oCompany.GetCompanyService();
                    oGeneralService = (SAPbobsCOM.GeneralService)oCompanyService.GetGeneralService("T2_CORETAX");

                    // identify document
                    oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(
                        SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oGeneralParams.SetProperty("DocEntry", docEntry);

                    // update posting date before close
                    oGeneralData = oGeneralService.GetByParams(oGeneralParams);
                    //oGeneralData.SetProperty("U_T2_Posting_Date", DateTime.Today);
                    //oGeneralService.Update(oGeneralData);

                    // close the document
                    oGeneralService.Close(oGeneralParams);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error CloseCoretax: {ex.Message}", ex);
                }
                finally
                {
                    if (oGeneralData != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralData);
                    if (oGeneralParams != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralParams);
                    if (oGeneralService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oGeneralService);
                    if (oCompanyService != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oCompanyService);
                }
            });
        }


        public static Task<CoretaxModel> GetCoretaxByKey(int docEntry)
        {
            return Task.Run(() => {
                try
                {
                    Company oCompany = CompanyService.GetCompany();

                    SAPbobsCOM.CompanyService oCompanyService = oCompany.GetCompanyService();

                    SAPbobsCOM.GeneralService oGeneralService = oCompanyService.GetGeneralService("T2_CORETAX");

                    SAPbobsCOM.GeneralDataParams oParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(
                        SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                    oParams.SetProperty("DocEntry", docEntry);

                    SAPbobsCOM.GeneralData oHeader = oGeneralService.GetByParams(oParams);

                    var series = GetSeriesByDocEntry("T2_CORETAX", docEntry);

                    DateTime? fromDate = null;
                    DateTime? toDate = null;
                    DateTime? postDate = null;
                    DateTime? docDate = null;
                    string fromDateStr = oHeader.GetProperty("U_T2_From_Date").ToString();
                    string toDateStr = oHeader.GetProperty("U_T2_To_Date").ToString();
                    string postDateStr = oHeader.GetProperty("U_T2_Posting_Date").ToString();
                    //string docDateStr = oHeader.GetProperty("U_T2_Doc_Date").ToString();
                    if (!string.IsNullOrEmpty(fromDateStr) && !fromDateStr.StartsWith("30/12/1899"))
                    {
                        fromDate = Convert.ToDateTime(oHeader.GetProperty("U_T2_From_Date"));
                    }
                    if (!string.IsNullOrEmpty(toDateStr) && !toDateStr.StartsWith("30/12/1899"))
                    {
                        toDate = Convert.ToDateTime(oHeader.GetProperty("U_T2_To_Date"));
                    }
                    //if (!string.IsNullOrEmpty(docDateStr) && !docDateStr.StartsWith("30/12/1899"))
                    //{
                    //    docDate = Convert.ToDateTime(oHeader.GetProperty("U_T2_Doc_Date"));
                    //}
                    if (!string.IsNullOrEmpty(postDateStr) && !postDateStr.StartsWith("30/12/1899"))
                    {
                        postDate = Convert.ToDateTime(oHeader.GetProperty("U_T2_Posting_Date"));
                    }

                    var model = new CoretaxModel
                    {
                        DocEntry = Convert.ToInt32(oHeader.GetProperty("DocEntry")),
                        DocNum = Convert.ToInt32(oHeader.GetProperty("DocNum")),
                        DocDate = docDate,
                        PostDate = postDate,
                        SeriesId = series.SeriesId,
                        SeriesName = series.SeriesName,
                        IsARInvoice = (oHeader.GetProperty("U_T2_AR_INV")?.ToString() == "Y"),
                        IsARDownPayment = (oHeader.GetProperty("U_T2_AR_DP")?.ToString() == "Y"),
                        IsARCreditMemo = (oHeader.GetProperty("U_T2_AR_CM")?.ToString() == "Y"),

                        FromDocNum = oHeader.GetProperty("U_T2_From_Doc")?.ToString(),
                        ToDocNum = oHeader.GetProperty("U_T2_To_Doc")?.ToString(),
                        FromDocEntry = oHeader.GetProperty("U_T2_From_Doc_Entry")?.ToString(),
                        ToDocEntry = oHeader.GetProperty("U_T2_To_Doc_Entry")?.ToString(),
                        FromDate = fromDate,
                        ToDate = toDate,

                        FromCust = oHeader.GetProperty("U_T2_From_Cust")?.ToString(),
                        ToCust = oHeader.GetProperty("U_T2_To_Cust")?.ToString(),
                        FromBranch = oHeader.GetProperty("U_T2_From_Branch")?.ToString(),
                        ToBranch = oHeader.GetProperty("U_T2_To_Branch")?.ToString(),
                        FromOutlet = oHeader.GetProperty("U_T2_From_Outlet")?.ToString(),
                        ToOutlet = oHeader.GetProperty("U_T2_To_Outlet")?.ToString(),

                        Status = oHeader.GetProperty("Status")?.ToString() ?? "O"
                    };

                    SAPbobsCOM.GeneralDataCollection children = oHeader.Child("T2_CORETAX_DT"); // UDO child table name
                    foreach (SAPbobsCOM.GeneralData line in children)
                    {
                        var detail = new InvoiceDataModel
                        {
                            TIN = line.GetProperty("U_T2_TIN")?.ToString(),
                            DocEntry = line.GetProperty("U_T2_DocEntry")?.ToString(),
                            LineNum = line.GetProperty("U_T2_LineNum")?.ToString(),
                            InvDate = line.GetProperty("U_T2_Inv_Date")?.ToString(),
                            NoDocument = line.GetProperty("U_T2_No_Doc")?.ToString(),
                            ObjectType = line.GetProperty("U_T2_Object_Type")?.ToString(),
                            ObjectName = line.GetProperty("U_T2_Object_Name")?.ToString(),
                            BPCode = line.GetProperty("U_T2_BP_Code")?.ToString(),
                            BPName = line.GetProperty("U_T2_BP_Name")?.ToString(),
                            SellerIDTKU = line.GetProperty("U_T2_Seller_IDTKU")?.ToString(),
                            BuyerDocument = line.GetProperty("U_T2_Buyer_Doc")?.ToString(),
                            NomorNPWP = line.GetProperty("U_T2_Nomor_NPWP")?.ToString(),
                            NPWPName = line.GetProperty("U_T2_NPWP_Name")?.ToString(),
                            NPWPAddress = line.GetProperty("U_T2_NPWP_Address")?.ToString(),
                            BuyerIDTKU = line.GetProperty("U_T2_Buyer_IDTKU")?.ToString(),
                            ItemCode = line.GetProperty("U_T2_Item_Code")?.ToString(),
                            ItemName = line.GetProperty("U_T2_Item_Name")?.ToString(),
                            ItemUnit = line.GetProperty("U_T2_Item_Unit")?.ToString(),
                            ItemPrice = line.GetProperty("U_T2_Item_Price") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_Item_Price")),
                            Qty = line.GetProperty("U_T2_Qty") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_Qty")),
                            TotalDisc = line.GetProperty("U_T2_Total_Disc") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_Total_Disc")),
                            TaxBase = line.GetProperty("U_T2_Tax_Base") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_Tax_Base")),
                            OtherTaxBase = line.GetProperty("U_T2_Other_Tax_Base") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_Other_Tax_Base")),
                            VATRate = line.GetProperty("U_T2_VAT_Rate") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_VAT_Rate")),
                            AmountVAT = line.GetProperty("U_T2_Amount_VAT") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_Amount_VAT")),
                            STLGRate = line.GetProperty("U_T2_STLG_Rate") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_STLG_Rate")),
                            STLG = line.GetProperty("U_T2_STLG") == null ? 0 : Convert.ToDouble(line.GetProperty("U_T2_STLG")),
                            JenisPajak = line.GetProperty("U_T2_Jenis_Pajak")?.ToString(),
                            KetTambahan = line.GetProperty("U_T2_Ket_Tambahan")?.ToString(),
                            PajakPengganti = line.GetProperty("U_T2_Pajak_Pengganti")?.ToString(),
                            Referensi = line.GetProperty("U_T2_Referensi")?.ToString(),
                            Status = line.GetProperty("U_T2_Status")?.ToString(),
                            KodeDokumenPendukung = line.GetProperty("U_T2_Kode_Dok_Pendukung")?.ToString(),
                            BranchCode = line.GetProperty("U_T2_Branch_Code")?.ToString(),
                            BranchName = line.GetProperty("U_T2_Branch_Name")?.ToString(),
                            OutletCode = line.GetProperty("U_T2_Outlet_Code")?.ToString(),
                            OutletName = line.GetProperty("U_T2_Outlet_Name")?.ToString(),
                            AddInfo = line.GetProperty("U_T2_Add_Info")?.ToString(),
                            BuyerCountry = line.GetProperty("U_T2_Buyer_Country")?.ToString(),
                            BuyerEmail = line.GetProperty("U_T2_Buyer_Email")?.ToString()
                        };

                        model.Detail.Add(detail);
                    }

                    return Task.FromResult(model);
                }
                catch (Exception ex)
                {
                    throw new Exception($"Error while getting Coretax UDO DocEntry={docEntry}: {ex.Message}", ex);
                }
            });
        }

        public static (int SeriesId, string SeriesName) GetSeriesByDocEntry(string udoTable, int docEntry)
        {
            Company oCompany = CompanyService.GetCompany();
            Recordset oRs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

            try
            {
                string sql = $@"
                SELECT T0.Series, ISNULL(T1.SeriesName, '') AS SeriesName
                FROM [@{udoTable}] T0
                LEFT JOIN NNM1 T1 ON T0.Series = T1.Series AND T1.ObjectCode = '{udoTable}'
                WHERE T0.DocEntry = {docEntry}";

                oRs.DoQuery(sql);

                if (!oRs.EoF)
                {
                    int seriesId = Convert.ToInt32(oRs.Fields.Item("Series").Value);
                    string seriesName = oRs.Fields.Item("SeriesName").Value.ToString();
                    return (seriesId, seriesName);
                }

                return (0, null);
            }
            catch (Exception ex)
            {
                throw new Exception("Error get series by DocEntry: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }

        public static string GetSeriesName(int seriesId)
        {
            Company oCompany = CompanyService.GetCompany();
            Recordset oRs = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);
            string seriesName = "";
            try
            {
                string sql = $@"
                SELECT ISNULL(T0.SeriesName, '') AS SeriesName
                FROM NNM1 T0
                WHERE T0.Series = {seriesId}";

                oRs.DoQuery(sql);

                if (!oRs.EoF)
                {
                    seriesName = oRs.Fields.Item("SeriesName").Value.ToString();
                }

                return seriesName;
            }
            catch (Exception ex)
            {
                throw new Exception("Error get series: " + ex.Message);
            }
            finally
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oRs);
            }
        }

        public static int GetSeriesIdCoretax()
        {
            try
            {
                Company oCompany = CompanyService.GetCompany();
                Recordset oRecordset = (Recordset)oCompany.GetBusinessObject(BoObjectTypes.BoRecordset);

                // panggil SQL function
                string sql = "SELECT dbo.T2_GET_SERIES_ID_CORETAX() AS SeriesId";
                oRecordset.DoQuery(sql);

                if (!oRecordset.EoF)
                {
                    return Convert.ToInt32(oRecordset.Fields.Item("SeriesId").Value);
                }
                else
                {
                    throw new Exception("SeriesId not found.");
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Error GetSeriesIdCoretax2025: {ex.Message}", ex);
            }
        }

        public static Task<List<FilterDataModel>> GetDataFilter(
            Dictionary<string, bool> selectedCkBox,
            string dtFrom, string dtTo, string docFrom, string docTo, string custFrom, string custTo,
            string branchFrom, string branchTo, string outFrom, string outTo
            )
        {
            return Task.Run(() => {
                List<FilterDataModel> result = new List<FilterDataModel>();
                try
                {
                    SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
                    List<string> docList = selectedCkBox.Where((ck) => ck.Value == true).Select((c) => c.Key).ToList();
                    SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                    string query = $@"EXEC [dbo].[T2_SP_CORETAX_HEADER] 
    @selected_doc = '{string.Join(",", docList)}'";

                    List<string> filters = new List<string>();

                    if (!string.IsNullOrEmpty(dtFrom))
                    {
                        filters.Add($"@from_date = '{dtFrom}'");
                    }
                    if (!string.IsNullOrEmpty(dtTo))
                    {
                        filters.Add($"@to_date = '{dtTo}'");
                    }
                    if (!string.IsNullOrEmpty(docFrom))
                    {
                        filters.Add($"@from_docentry = {docFrom}");
                    }
                    if (!string.IsNullOrEmpty(docTo))
                    {
                        filters.Add($"@to_docentry = {docTo}");
                    }
                    if (!string.IsNullOrEmpty(custFrom))
                    {
                        filters.Add($"@from_cust = '{custFrom}'");
                    }
                    if (!string.IsNullOrEmpty(custTo))
                    {
                        filters.Add($"@to_cust = '{custTo}'");
                    }
                    if (!string.IsNullOrEmpty(branchFrom))
                    {
                        filters.Add($"@from_branch = {branchFrom}");
                    }
                    if (!string.IsNullOrEmpty(branchTo))
                    {
                        filters.Add($"@to_branch = {branchTo}");
                    }
                    if (!string.IsNullOrEmpty(outFrom))
                    {
                        filters.Add($"@from_outlet = '{outFrom}'");
                    }
                    if (!string.IsNullOrEmpty(outTo))
                    {
                        filters.Add($"@to_outlet = '{outTo}'");
                    }

                    // join filters with commas
                    if (filters.Count > 0)
                    {
                        query += ", " + string.Join(", ", filters);
                    }

                    rs.DoQuery(query);

                    while (!rs.EoF)
                    {
                        var model = new FilterDataModel
                        {
                            DocEntry = rs.Fields.Item("DocEntry").Value?.ToString(),
                            DocNo = rs.Fields.Item("NoDocument").Value?.ToString(),
                            CardCode = rs.Fields.Item("BPCode").Value?.ToString(),
                            CardName = rs.Fields.Item("BPName").Value?.ToString(),
                            ObjType = rs.Fields.Item("ObjectType").Value?.ToString(),
                            ObjName = rs.Fields.Item("ObjectName").Value?.ToString(),
                            PostDate = rs.Fields.Item("InvDate").Value?.ToString(),
                            BranchCode = rs.Fields.Item("BranchCode").Value?.ToString(),
                            BranchName = rs.Fields.Item("BranchName").Value?.ToString(),
                            OutletCode = rs.Fields.Item("OutletCode").Value?.ToString(),
                            OutletName = rs.Fields.Item("OutletName").Value?.ToString(),
                            Selected = false,
                        };

                        result.Add(model);
                        rs.MoveNext();
                    }
                }
                catch (Exception)
                {

                    throw;
                }
                return Task.FromResult(result);
            });
        }

        public static Task<List<InvoiceDataModel>> GetDataGenerate(List<FilterDataModel> filteredHeader)
        {
            return Task.Run(() =>
            {
                List<InvoiceDataModel> result = new List<InvoiceDataModel>();
                SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                // --- group headers by ObjType ---
                var dict = filteredHeader
                    .GroupBy(x => x.ObjType)
                    .ToDictionary(g => g.Key, g => g.Select(x => x.DocEntry).ToList());

                // --- build JSON ---
                var payload = dict.Select(g => new
                {
                    ObjType = g.Key,
                    DocEntries = string.Join(",", g.Value)
                }).ToList();

                // --- build JSON manually ---
                string json = "[" + string.Join(",", dict.Select(g =>
                    "{ \"ObjType\": \"" + g.Key + "\", \"DocEntries\": \"" + string.Join(",", g.Value) + "\" }"
                )) + "]";

                // --- execute new SP ---
                string query = $@"EXEC [dbo].[T2_SP_CORETAX_GENERATE] @DocList = N'{json}'";
                rs.DoQuery(query);

                while (!rs.EoF)
                {
                    var model = new InvoiceDataModel
                    {
                        TIN = rs.Fields.Item("TIN").Value?.ToString(),
                        DocEntry = rs.Fields.Item("DocEntry").Value?.ToString(),
                        LineNum = rs.Fields.Item("LineNum").Value?.ToString(),
                        InvDate = rs.Fields.Item("InvDate").Value?.ToString(),
                        NoDocument = rs.Fields.Item("NoDocument").Value?.ToString(),
                        ObjectType = rs.Fields.Item("ObjectType").Value?.ToString(),
                        ObjectName = rs.Fields.Item("ObjectName").Value?.ToString(),
                        BPCode = rs.Fields.Item("BPCode").Value?.ToString(),
                        BPName = rs.Fields.Item("BPName").Value?.ToString(),
                        SellerIDTKU = rs.Fields.Item("SellerIDTKU").Value?.ToString(),
                        BuyerDocument = rs.Fields.Item("BuyerDocument").Value?.ToString(),
                        NomorNPWP = rs.Fields.Item("NomorNPWP").Value?.ToString(),
                        NPWPName = rs.Fields.Item("NPWPName").Value?.ToString(),
                        NPWPAddress = rs.Fields.Item("NPWPAddress").Value?.ToString(),
                        BuyerIDTKU = rs.Fields.Item("BuyerIDTKU").Value?.ToString(),
                        ItemCode = rs.Fields.Item("ItemCode").Value?.ToString(),
                        DefItemCode = rs.Fields.Item("DefItemCode").Value?.ToString(),
                        ItemName = rs.Fields.Item("ItemName").Value?.ToString(),
                        ItemUnit = rs.Fields.Item("ItemUnit").Value?.ToString(),

                        // decimals
                        ItemPrice = Convert.ToDouble(rs.Fields.Item("ItemPrice").Value ?? 0),
                        Qty = Convert.ToDouble(rs.Fields.Item("Qty").Value ?? 0),
                        TotalDisc = Convert.ToDouble(rs.Fields.Item("TotalDisc").Value ?? 0),
                        TaxBase = Convert.ToDouble(rs.Fields.Item("TaxBase").Value ?? 0),
                        OtherTaxBase = Convert.ToDouble(rs.Fields.Item("OtherTaxBase").Value ?? 0),
                        VATRate = Convert.ToDouble(rs.Fields.Item("VATRate").Value ?? 0),
                        AmountVAT = Convert.ToDouble(rs.Fields.Item("AmountVAT").Value ?? 0),
                        STLGRate = Convert.ToDouble(rs.Fields.Item("STLGRate").Value ?? 0),
                        STLG = Convert.ToDouble(rs.Fields.Item("STLG").Value ?? 0),

                        // strings
                        JenisPajak = rs.Fields.Item("JenisPajak").Value?.ToString(),
                        KetTambahan = rs.Fields.Item("KetTambahan").Value?.ToString(),
                        PajakPengganti = rs.Fields.Item("PajakPengganti").Value?.ToString(),
                        Referensi = rs.Fields.Item("Referensi").Value?.ToString(),
                        Status = rs.Fields.Item("Status").Value?.ToString(),
                        KodeDokumenPendukung = rs.Fields.Item("KodeDokumenPendukung").Value?.ToString(),
                        BranchCode = rs.Fields.Item("BranchCode").Value?.ToString(),
                        BranchName = rs.Fields.Item("BranchName").Value?.ToString(),
                        OutletCode = rs.Fields.Item("OutletCode").Value?.ToString(),
                        OutletName = rs.Fields.Item("OutletName").Value?.ToString(),
                        AddInfo = rs.Fields.Item("AddInfo").Value?.ToString(),
                        BuyerCountry = rs.Fields.Item("BuyerCountry").Value?.ToString(),
                        BuyerEmail = rs.Fields.Item("BuyerEmail").Value?.ToString(),
                    };

                    result.Add(model);
                    rs.MoveNext();
                }

                return Task.FromResult(result);
            });
        }


        //public static Task<List<InvoiceDataModel>> GetDataGenerate(List<FilterDataModel> filteredHeader)
        //{
        //    return Task.Run(() =>
        //    {
        //        List<InvoiceDataModel> result = new List<InvoiceDataModel>();
        //        SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
        //        SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

        //        var dict = filteredHeader
        //        .GroupBy(x => x.ObjType)
        //        .ToDictionary(g => g.Key, g => g.ToList());

        //        foreach (var key in dict.Keys)
        //        {
        //            var invDict = dict[key];
        //            if (invDict.Any())
        //            {
        //                var docEntryCsv = string.Join(",", invDict.Select(x => x.DocEntry));
        //                string query = $@"EXEC [dbo].[T2_SP_CORETAX_GENERATE] 
        //                @DocList = '{docEntryCsv}', @ObjectType = '{key}'";
        //                rs.DoQuery(query);

        //                while (!rs.EoF)
        //                {
        //                    var model = new InvoiceDataModel
        //                    {
        //                        //No = rs.Fields.Item("No").Value?.ToString(),
        //                        TIN = rs.Fields.Item("TIN").Value?.ToString(),
        //                        DocEntry = rs.Fields.Item("DocEntry").Value?.ToString(),
        //                        LineNum = rs.Fields.Item("LineNum").Value?.ToString(),
        //                        InvDate = rs.Fields.Item("InvDate").Value?.ToString(),
        //                        NoDocument = rs.Fields.Item("NoDocument").Value?.ToString(),
        //                        ObjectType = rs.Fields.Item("ObjectType").Value?.ToString(),
        //                        ObjectName = rs.Fields.Item("ObjectName").Value?.ToString(),
        //                        BPCode = rs.Fields.Item("BPCode").Value?.ToString(),
        //                        BPName = rs.Fields.Item("BPName").Value?.ToString(),
        //                        SellerIDTKU = rs.Fields.Item("SellerIDTKU").Value?.ToString(),
        //                        BuyerDocument = rs.Fields.Item("BuyerDocument").Value?.ToString(),
        //                        NomorNPWP = rs.Fields.Item("NomorNPWP").Value?.ToString(),
        //                        NPWPName = rs.Fields.Item("NPWPName").Value?.ToString(),
        //                        NPWPAddress = rs.Fields.Item("NPWPAddress").Value?.ToString(),
        //                        BuyerIDTKU = rs.Fields.Item("BuyerIDTKU").Value?.ToString(),
        //                        ItemCode = rs.Fields.Item("ItemCode").Value?.ToString(),
        //                        DefItemCode = rs.Fields.Item("DefItemCode").Value?.ToString(),
        //                        ItemName = rs.Fields.Item("ItemName").Value?.ToString(),
        //                        ItemUnit = rs.Fields.Item("ItemUnit").Value?.ToString(),

        //                        // --- decimals (safe conversion) ---
        //                        ItemPrice = Convert.ToDouble(rs.Fields.Item("ItemPrice").Value ?? 0),
        //                        Qty = Convert.ToDouble(rs.Fields.Item("Qty").Value ?? 0),
        //                        TotalDisc = Convert.ToDouble(rs.Fields.Item("TotalDisc").Value ?? 0),
        //                        TaxBase = Convert.ToDouble(rs.Fields.Item("TaxBase").Value ?? 0),
        //                        OtherTaxBase = Convert.ToDouble(rs.Fields.Item("OtherTaxBase").Value ?? 0),
        //                        VATRate = Convert.ToDouble(rs.Fields.Item("VATRate").Value ?? 0),
        //                        AmountVAT = Convert.ToDouble(rs.Fields.Item("AmountVAT").Value ?? 0),
        //                        STLGRate = Convert.ToDouble(rs.Fields.Item("STLGRate").Value ?? 0),
        //                        STLG = Convert.ToDouble(rs.Fields.Item("STLG").Value ?? 0),

        //                        // --- strings ---
        //                        JenisPajak = rs.Fields.Item("JenisPajak").Value?.ToString(),
        //                        KetTambahan = rs.Fields.Item("KetTambahan").Value?.ToString(),
        //                        PajakPengganti = rs.Fields.Item("PajakPengganti").Value?.ToString(),
        //                        Referensi = rs.Fields.Item("Referensi").Value?.ToString(),
        //                        Status = rs.Fields.Item("Status").Value?.ToString(),
        //                        KodeDokumenPendukung = rs.Fields.Item("KodeDokumenPendukung").Value?.ToString(),
        //                        BranchCode = rs.Fields.Item("BranchCode").Value?.ToString(),
        //                        BranchName = rs.Fields.Item("BranchName").Value?.ToString(),
        //                        OutletCode = rs.Fields.Item("OutletCode").Value?.ToString(),
        //                        OutletName = rs.Fields.Item("OutletName").Value?.ToString(),
        //                        AddInfo = rs.Fields.Item("AddInfo").Value?.ToString(),
        //                        BuyerCountry = rs.Fields.Item("BuyerCountry").Value?.ToString(),
        //                        BuyerEmail = rs.Fields.Item("BuyerEmail").Value?.ToString(),
        //                    };

        //                    result.Add(model);
        //                    rs.MoveNext();
        //                }
        //            }
        //        }
        //        return Task.FromResult(result);
        //    });
        //}

        public static Task<int> GetLastDocNum(int seriesId)
        {
            int docNum = 0;
            try
            {
                SAPbobsCOM.Company oCompany = Services.CompanyService.GetCompany();
                SAPbobsCOM.Recordset rs = (SAPbobsCOM.Recordset)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset);

                string query = $"SELECT NextNumber FROM NNM1 WHERE Series = '{seriesId}'";
                rs.DoQuery(query);
                if (!rs.EoF)
                {
                    docNum = (int)rs.Fields.Item("NextNumber").Value;
                }
                return Task.FromResult(docNum);
            }
            catch (Exception e)
            {

                throw e;
            }
        }

        public static Task UpdateStatusInv(int docNum, List<InvoiceDataModel> datas)
        {
            return Task.Run(() =>
            {
                try
                {
                    SAPbobsCOM.Company oCompany = CompanyService.GetCompany();
                    if (datas != null && datas.Any())
                    {
                        var gInvList = datas
                            .GroupBy(p => new
                            {
                                p.DocEntry,
                                p.NoDocument,
                                p.BPCode,
                                p.BPName,
                                p.ObjectType,
                                p.ObjectName,
                                p.InvDate,
                                p.BranchCode,
                                p.BranchName,
                                p.OutletCode,
                                p.OutletName
                            })
                            .Select(g => new FilterDataModel
                            {
                                DocEntry = g.Key.DocEntry,
                                DocNo = g.Key.NoDocument,
                                CardCode = g.Key.BPCode,
                                CardName = g.Key.BPName,
                                ObjType = g.Key.ObjectType,
                                ObjName = g.Key.ObjectName,
                                PostDate = g.Key.InvDate,
                                BranchCode = g.Key.BranchCode,
                                BranchName = g.Key.BranchName,
                                OutletCode = g.Key.OutletCode,
                                OutletName = g.Key.OutletName,
                                Selected = true
                            })
                            .ToList();

                        foreach (var item in gInvList)
                        {
                            switch (item.ObjType)
                            {
                                case "13": // AR Invoice
                                    UpdateArInvoice(oCompany, int.Parse(item.DocEntry), (docNum), "Y");
                                    break;
                                case "14": // AR Credit Memo
                                    UpdateArCreditMemo(oCompany, int.Parse(item.DocEntry), (docNum), "Y");
                                    break;
                                case "203": // AR Down Payment
                                    UpdateArDownPayment(oCompany, int.Parse(item.DocEntry), (docNum), "Y");
                                    break;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Update Status Invoice error: {ex.Message}");
                }
            });
        }

        public static void UpdateArInvoice(SAPbobsCOM.Company oCompany, int docEntry, int docNum, string status)
        {
            SAPbobsCOM.Documents oInvoice = null;

            try
            {
                // Get the invoice
                oInvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oInvoices);

                if (oInvoice.GetByKey(docEntry))
                {
                    // Set the UDF value
                    oInvoice.UserFields.Fields.Item("U_T2_Exported").Value = status;
                    oInvoice.UserFields.Fields.Item("U_T2_Coretax_No").Value = docNum;

                    // Update the invoice
                    int retCode = oInvoice.Update();
                    if (retCode != 0)
                    {
                        oCompany.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Failed to update OINV UDF. Error {errCode}: {errMsg}");
                    }
                }
            }
            finally
            {
                if (oInvoice != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);
            }

        }

        public static void UpdateArDownPayment(SAPbobsCOM.Company oCompany, int docEntry, int docNum, string status)
        {
            SAPbobsCOM.Documents oInvoice = null;

            try

            {
                // Get the invoice
                oInvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oDownPayments);

                if (oInvoice.GetByKey(docEntry))
                {
                    // Set the UDF value
                    oInvoice.UserFields.Fields.Item("U_T2_Exported").Value = status;
                    oInvoice.UserFields.Fields.Item("U_T2_Coretax_No").Value = docNum;

                    // Update the invoice
                    int retCode = oInvoice.Update();
                    if (retCode != 0)
                    {
                        oCompany.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Failed to update OINV UDF. Error {errCode}: {errMsg}");
                    }
                }
            }
            finally
            {
                if (oInvoice != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);
            }

        }

        public static void UpdateArCreditMemo(SAPbobsCOM.Company oCompany, int docEntry, int docNum, string status)
        {
            SAPbobsCOM.Documents oInvoice = null;

            try

            {
                // Get the invoice
                oInvoice = (SAPbobsCOM.Documents)oCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oCreditNotes);

                if (oInvoice.GetByKey(docEntry))
                {
                    // Set the UDF value
                    oInvoice.UserFields.Fields.Item("U_T2_Exported").Value = status;
                    oInvoice.UserFields.Fields.Item("U_T2_Coretax_No").Value = docNum;

                    // Update the invoice
                    int retCode = oInvoice.Update();
                    if (retCode != 0)
                    {
                        oCompany.GetLastError(out int errCode, out string errMsg);
                        throw new Exception($"Failed to update OINV UDF. Error {errCode}: {errMsg}");
                    }
                }
            }
            finally
            {
                if (oInvoice != null) System.Runtime.InteropServices.Marshal.ReleaseComObject(oInvoice);
            }

        }

    }
}
