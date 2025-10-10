using EFakturCoretax.Models;
using SAPbobsCOM;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.Services
{
    public static class TransactionService
    {
        public static void CloseCoretax(Company oCompany, int docEntry)
        {
            SAPbobsCOM.CompanyService oCompanyService = null;
            SAPbobsCOM.GeneralService oGeneralService = null;
            SAPbobsCOM.GeneralDataParams oGeneralParams = null;
            SAPbobsCOM.GeneralData oGeneralData = null;

            try
            {
                oCompanyService = oCompany.GetCompanyService();
                oGeneralService = (SAPbobsCOM.GeneralService)oCompanyService.GetGeneralService("T2_CORETAX");

                // identify document
                oGeneralParams = (SAPbobsCOM.GeneralDataParams)oGeneralService.GetDataInterface(
                    SAPbobsCOM.GeneralServiceDataInterfaces.gsGeneralDataParams);
                oGeneralParams.SetProperty("DocEntry", docEntry);

                // update posting date before close
                oGeneralData = oGeneralService.GetByParams(oGeneralParams);

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
                            Revise = false,
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

        public static Task<List<InvoiceDataModel>> GetDataGenerate(List<FilterDataModel> filteredHeader, decimal vatRate)
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
                string query = $@"EXEC [dbo].[T2_SP_CORETAX_GENERATE] @VatRate = {vatRate.ToString(CultureInfo.InvariantCulture)}, @DocList = N'{json}'";
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

                        // doubles
                        ItemPrice = Math.Round(Convert.ToDecimal(rs.Fields.Item("ItemPrice").Value, CultureInfo.InvariantCulture),2),
                        Qty = Math.Round(Convert.ToDecimal(rs.Fields.Item("Qty").Value, CultureInfo.InvariantCulture),2),
                        TotalDisc = Math.Round(Convert.ToDecimal(rs.Fields.Item("TotalDisc").Value, CultureInfo.InvariantCulture),2),
                        TaxBase = Math.Round(Convert.ToDecimal(rs.Fields.Item("TaxBase").Value, CultureInfo.InvariantCulture),2),
                        OtherTaxBase = Math.Round(Convert.ToDecimal(rs.Fields.Item("OtherTaxBase").Value, CultureInfo.InvariantCulture),2),
                        VATRate = Math.Round(Convert.ToDecimal(rs.Fields.Item("VATRate").Value, CultureInfo.InvariantCulture),2),
                        AmountVAT = Math.Round(Convert.ToDecimal(rs.Fields.Item("AmountVAT").Value, CultureInfo.InvariantCulture),2),
                        STLGRate = Math.Round(Convert.ToDecimal(rs.Fields.Item("STLGRate").Value, CultureInfo.InvariantCulture),2),
                        STLG = Math.Round(Convert.ToDecimal(rs.Fields.Item("STLG").Value, CultureInfo.InvariantCulture),2),
                        CoretaxVatAmount = Math.Round(Convert.ToDecimal(rs.Fields.Item("CoretaxVatAmount").Value, CultureInfo.InvariantCulture),2),
                        CoretaxVatRate = Math.Round(Convert.ToDecimal(rs.Fields.Item("CoretaxVatRate").Value, CultureInfo.InvariantCulture),2),

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
                        Revise = "N"
                    };

                    result.Add(model);
                    rs.MoveNext();
                }

                return Task.FromResult(result);
            });
        }

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

        public static void UpdateStatusInv(SAPbobsCOM.Company oCompany, string docNum, List<InvoiceDataModel> datas)
        {
            try
            {
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
                            p.OutletName,
                            p.Revise
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
                            Selected = true,
                            Revise = g.Key.Revise == "Y"
                        })
                        .ToList();

                    foreach (var item in gInvList)
                    {
                        switch (item.ObjType)
                        {
                            case "13": // AR Invoice
                                UpdateArInvoice(oCompany, int.Parse(item.DocEntry), docNum, "Y");
                                break;
                            case "14": // AR Credit Memo
                                UpdateArCreditMemo(oCompany, int.Parse(item.DocEntry), docNum, "Y");
                                break;
                            case "203": // AR Down Payment
                                UpdateArDownPayment(oCompany, int.Parse(item.DocEntry), docNum, "Y");
                                break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Update Status Invoice error: {ex.Message}");
            }
        }

        public static void ReviseInvoice(SAPbobsCOM.Company oCompany, string docEntry, string objType)
        {
            try
            {
                switch (objType)
                {
                    case "13": // AR Invoice
                        UpdateArInvoice(oCompany, int.Parse(docEntry), "", "N");
                        break;
                    case "14": // AR Credit Memo
                        UpdateArCreditMemo(oCompany, int.Parse(docEntry), "", "N");
                        break;
                    case "203": // AR Down Payment
                        UpdateArDownPayment(oCompany, int.Parse(docEntry), "", "N");
                        break;
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Revise Invoice error: {ex.Message}");
            }
        }


        public static void UpdateArInvoice(SAPbobsCOM.Company oCompany, int docEntry, string docNum, string status)
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

        public static void UpdateArDownPayment(SAPbobsCOM.Company oCompany, int docEntry, string docNum, string status)
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

        public static void UpdateArCreditMemo(SAPbobsCOM.Company oCompany, int docEntry, string docNum, string status)
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
