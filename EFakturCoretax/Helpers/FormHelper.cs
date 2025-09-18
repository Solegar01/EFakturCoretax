using EFakturCoretax.Models;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.Helpers
{
    public static class FormHelper
    {
        public static void ClearCombo(SAPbouiCOM.Form form, string id)
        {
            try
            {
                //if (!ItemIsExists(form, id)) return;
                SAPbouiCOM.Item comboItem = form.Items.Item(id);
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)comboItem.Specific;

                // Remove any default values
                while (oCombo.ValidValues.Count > 0)
                {
                    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
                }
                oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void ResetSelectCb(SAPbouiCOM.Form oForm, string id)
        {
            try
            {
                SAPbouiCOM.Item itemTo = oForm.Items.Item(id);
                SAPbouiCOM.ComboBox cbTo = (SAPbouiCOM.ComboBox)itemTo.Specific;
                cbTo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void RemoveFocus(SAPbouiCOM.Form oForm)
        {
            try
            {
                if (!ItemIsExists(oForm, "dummy"))
                {
                    var dummyItem = oForm.Items.Add("dummy", SAPbouiCOM.BoFormItemTypes.it_EDIT);
                    dummyItem.Left = 5;
                    dummyItem.Top = 5;
                    dummyItem.Width = 1;
                    dummyItem.Height = 1;
                    ((SAPbouiCOM.EditText)dummyItem.Specific).Value = "";
                }

                // Force focus to dummy edit field
                oForm.Items.Item("dummy").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                oForm.ActiveItem = "dummy";
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static bool ItemIsExists(SAPbouiCOM.Form oForm, string itemUid)
        {
            try
            {
                for (int i = 0; i < oForm.Items.Count; i++)
                {
                    var c = oForm.Items.Item(i).UniqueID;
                    if (oForm.Items.Item(i).UniqueID == itemUid) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        public static bool DsIsExists(SAPbouiCOM.Form oForm, string dsUid)
        {
            try
            {
                for (int i = 0; i < oForm.DataSources.UserDataSources.Count; i++)
                {
                    var c = oForm.DataSources.UserDataSources.Item(i).UID;
                    if (oForm.DataSources.UserDataSources.Item(i).UID == dsUid) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        public static bool DBSIsExists(SAPbouiCOM.Form oForm, string tableName)
        {
            try
            {
                for (int i = 0; i < oForm.DataSources.DBDataSources.Count; i++)
                {
                    var c = oForm.DataSources.DBDataSources.Item(i).TableName;
                    if (oForm.DataSources.DBDataSources.Item(i).TableName == tableName) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        public static bool DtIsExists(SAPbouiCOM.Form oForm, string dsUid)
        {
            try
            {
                for (int i = 0; i < oForm.DataSources.DataTables.Count; i++)
                {
                    var c = oForm.DataSources.DataTables.Item(i).UniqueID;
                    if (oForm.DataSources.DataTables.Item(i).UniqueID == dsUid) // 1-based index
                        return true;
                }
            }
            catch (Exception)
            {

                throw;
            }

            return false;
        }

        public static void ClearEdit(SAPbouiCOM.Form oForm, string id)
        {
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.EditText edit = (SAPbouiCOM.EditText)oForm.Items.Item(id).Specific;
                edit.Value = "";
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void ClearEditDate(SAPbouiCOM.Form oForm, string id, string ds)
        {
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                ClearEdit(oForm, id);
                oForm.DataSources.UserDataSources.Item(ds).Value = "00000000"; // -> tampil 30.12.1899
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void ClearCheckBox(SAPbouiCOM.Form oForm, string id, string ds)
        {
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.CheckBox ck = (SAPbouiCOM.CheckBox)oForm.Items.Item(id).Specific;
                ck.Checked = false;
                oForm.DataSources.UserDataSources.Item(ds).Value = "N";
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetBoldUnderlinedLabel(SAPbouiCOM.Form oForm, string id)
        {
            try
            {
                // Get StaticText item
                SAPbouiCOM.StaticText label = (SAPbouiCOM.StaticText)oForm.Items.Item(id).Specific;

                // Bold Underline
                label.Item.TextStyle = (int)(SAPbouiCOM.BoFontStyle.fs_Bold | SAPbouiCOM.BoFontStyle.fs_Underline);
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void ClearGrid(SAPbouiCOM.Form oForm, string gridItemUID, string dataTableUID)
        {
            try
            {
                // Check if the DataTable exists
                SAPbouiCOM.DataTable dt = null;
                if (DtIsExists(oForm, dataTableUID))
                {
                    dt = oForm.DataSources.DataTables.Item(dataTableUID);
                    dt.Clear();
                }
                else
                {
                    // Recreate DataTable with all columns if it doesn't exist
                    dt = oForm.DataSources.DataTables.Add(dataTableUID);
                    dt.Clear();
                }

                // Bind DataTable to Grid
                SAPbouiCOM.Grid oGrid = (SAPbouiCOM.Grid)oForm.Items.Item(gridItemUID).Specific;
                oGrid.DataTable = dt;

            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void ClearMatrix(SAPbouiCOM.Form oForm, string mtUID, string dtUID = null, string tbUID = null)
        {
            try
            {
                //if (!ItemIsExists(oForm, mtUID)) return;
                SAPbouiCOM.Matrix oMtx = (SAPbouiCOM.Matrix)oForm.Items.Item(mtUID).Specific;
                if (oMtx.RowCount <= 0) return;
                oMtx.Clear();
                if (!string.IsNullOrEmpty(dtUID))
                {
                    // Check if the DataTable exists
                    SAPbouiCOM.DataTable dt = null;
                    if (DtIsExists(oForm, dtUID))
                    {
                        dt = oForm.DataSources.DataTables.Item(dtUID);
                        dt.Clear();
                    }
                    else
                    {
                        // Recreate DataTable with all columns if it doesn't exist
                        dt = oForm.DataSources.DataTables.Add(dtUID);
                        dt.Clear();
                    }
                }

                if (!string.IsNullOrEmpty(tbUID))
                {
                    SAPbouiCOM.DBDataSource db = null;
                    if (DBSIsExists(oForm, tbUID))
                    {
                        db = oForm.DataSources.DBDataSources.Item(tbUID);
                        db.Clear();
                    }
                    else
                    {
                        // Recreate DataTable with all columns if it doesn't exist
                        db = oForm.DataSources.DBDataSources.Add(tbUID);
                        db.Clear();
                    }
                }


            }
            catch (Exception)
            {
                throw;
            }
        }

        public static void SetCustomerCfl(SAPbouiCOM.Form oForm, string id, string txtId)
        {
            // ✅ Check if CFL already exists
            bool cflExists = oForm.ChooseFromLists
                .Cast<SAPbouiCOM.ChooseFromList>()
                .Any(cfl => cfl.UniqueID == id);

            if (!cflExists)
            {
                // Create CFL parameters
                var oCFLParams = (SAPbouiCOM.ChooseFromListCreationParams)
                    Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                oCFLParams.MultiSelection = false;
                oCFLParams.ObjectType = "2"; // Business Partner
                oCFLParams.UniqueID = id;

                // Add CFL to form
                var oCFL = oForm.ChooseFromLists.Add(oCFLParams);

                // Add conditions
                var conditions = oCFL.GetConditions();
                var condition = conditions.Add();
                condition.Alias = "CardType";
                condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                condition.CondVal = "C"; // Customer
                oCFL.SetConditions(conditions);
            }
            var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
            oEdit.ChooseFromListUID = id;      // ✅ use your parameter `id` instead of hardcoded "2"
            oEdit.ChooseFromListAlias = "CardCode";
        }

        public static void SetDocumentCfl(SAPbouiCOM.Form oForm, string id, string txtId, string selectedDoc)
        {
            try
            {
                // ✅ Check if CFL already exists
                bool cflExists = oForm.ChooseFromLists
                    .Cast<SAPbouiCOM.ChooseFromList>()
                    .Any(cfl => cfl.UniqueID == id);
                //string selectedDoc = SelectedCkBox.Where((d) => d.Value).First().Key;
                if (!cflExists)
                {
                    // Create CFL parameters
                    var oCFLParams = (SAPbouiCOM.ChooseFromListCreationParams)
                        Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_ChooseFromListCreationParams);

                    oCFLParams.MultiSelection = false;
                    switch (selectedDoc)
                    {
                        case "OINV":
                            oCFLParams.ObjectType = "13";
                            break;
                        case "ODPI":
                            oCFLParams.ObjectType = "203";
                            break;
                        case "ORIN":
                            oCFLParams.ObjectType = "14";
                            break;
                        default:
                            break;
                    }
                    oCFLParams.UniqueID = id;

                    // Add CFL to form
                    var oCFL = oForm.ChooseFromLists.Add(oCFLParams);

                    // Add conditions
                    var conditions = oCFL.GetConditions();
                    var condition = conditions.Add();
                    condition.Alias = "CANCELED";
                    condition.Operation = SAPbouiCOM.BoConditionOperation.co_EQUAL;
                    condition.CondVal = "N"; // Customer
                    oCFL.SetConditions(conditions);
                }

                var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(txtId).Specific;
                oEdit.ChooseFromListUID = id;
                oEdit.ChooseFromListAlias = "DocNum";
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetEnabled(SAPbouiCOM.Form oForm, string[] itemIds, bool enabled)
        {
            try
            {
                foreach (var id in itemIds)
                {
                    if (!enabled && oForm.ActiveItem == id)
                    {
                        RemoveFocus(oForm);
                    }

                    oForm.Items.Item(id).Enabled = enabled;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetValueEdit(SAPbouiCOM.Form oForm, string id, string val)
        {
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.EditText oEdit = (SAPbouiCOM.EditText)oForm.Items.Item(id).Specific;
                oEdit.Value = val;
                oEdit.Active = false;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetValueCheck(SAPbouiCOM.Form oForm, string id, bool val)
        {
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.CheckBox oCheck = (SAPbouiCOM.CheckBox)oForm.Items.Item(id).Specific;
                oCheck.Checked = val;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetValueCb(SAPbouiCOM.Form oForm, string id, string val)
        {
            try
            {
                //if (!ItemIsExists(oForm, id)) return;
                SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item(id).Specific;
                oCombo.Select(val, SAPbouiCOM.BoSearchKey.psk_ByValue);
                oCombo.Active = false;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static void SetValueDS(SAPbouiCOM.Form oForm, string id, string val)
        {
            try
            {
                if (!DsIsExists(oForm, id)) return;
                oForm.DataSources.UserDataSources.Item(id).Value = val;
            }
            catch (Exception)
            {

                throw;
            }
        }

        public static bool HasCfl(SAPbouiCOM.Form oForm, string cflUID)
        {
            try
            {
                var _ = oForm.ChooseFromLists.Item(cflUID);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static List<InvoiceDataModel> BuildInvoiceDetailList(SAPbouiCOM.DBDataSource oDBDS_Detail)
        {
            var listData = new List<InvoiceDataModel>();

            for (int i = 0; i < oDBDS_Detail.Size; i++)
            {
                var detail = new InvoiceDataModel
                {
                    TIN = oDBDS_Detail.GetValue("U_T2_TIN", i)?.Trim(),
                    DocEntry = oDBDS_Detail.GetValue("U_T2_DocEntry", i)?.Trim(),
                    LineNum = oDBDS_Detail.GetValue("U_T2_LineNum", i)?.Trim(),
                    InvDate = oDBDS_Detail.GetValue("U_T2_Inv_Date", i)?.Trim(),
                    NoDocument = oDBDS_Detail.GetValue("U_T2_No_Doc", i)?.Trim(),
                    ObjectType = oDBDS_Detail.GetValue("U_T2_Object_Type", i)?.Trim(),
                    ObjectName = oDBDS_Detail.GetValue("U_T2_Object_Name", i)?.Trim(),
                    BPCode = oDBDS_Detail.GetValue("U_T2_BP_Code", i)?.Trim(),
                    BPName = oDBDS_Detail.GetValue("U_T2_BP_Name", i)?.Trim(),

                    SellerIDTKU = oDBDS_Detail.GetValue("U_T2_Seller_IDTKU", i)?.Trim(),
                    BuyerDocument = oDBDS_Detail.GetValue("U_T2_Buyer_Doc", i)?.Trim(),
                    NomorNPWP = oDBDS_Detail.GetValue("U_T2_Nomor_NPWP", i)?.Trim(),
                    NPWPName = oDBDS_Detail.GetValue("U_T2_NPWP_Name", i)?.Trim(),
                    NPWPAddress = oDBDS_Detail.GetValue("U_T2_NPWP_Address", i)?.Trim(),
                    BuyerIDTKU = oDBDS_Detail.GetValue("U_T2_Buyer_IDTKU", i)?.Trim(),
                    ItemCode = oDBDS_Detail.GetValue("U_T2_Item_Code", i)?.Trim(),
                    DefItemCode = "000000",
                    ItemName = oDBDS_Detail.GetValue("U_T2_Item_Name", i)?.Trim(),
                    ItemUnit = oDBDS_Detail.GetValue("U_T2_Item_Unit", i)?.Trim(),

                    // --- decimals (safe conversion) ---
                    ItemPrice = SafeToDouble(oDBDS_Detail.GetValue("U_T2_Item_Price", i)),
                    Qty = SafeToDouble(oDBDS_Detail.GetValue("U_T2_Qty", i)),
                    TotalDisc = SafeToDouble(oDBDS_Detail.GetValue("U_T2_Total_Disc", i)),
                    TaxBase = SafeToDouble(oDBDS_Detail.GetValue("U_T2_Tax_Base", i)),
                    OtherTaxBase = SafeToDouble(oDBDS_Detail.GetValue("U_T2_Other_Tax_Base", i)),
                    VATRate = SafeToDouble(oDBDS_Detail.GetValue("U_T2_VAT_Rate", i)),
                    AmountVAT = SafeToDouble(oDBDS_Detail.GetValue("U_T2_Amount_VAT", i)),
                    STLGRate = SafeToDouble(oDBDS_Detail.GetValue("U_T2_STLG_Rate", i)),
                    STLG = SafeToDouble(oDBDS_Detail.GetValue("U_T2_STLG", i)),

                    // --- strings ---
                    JenisPajak = oDBDS_Detail.GetValue("U_T2_Jenis_Pajak", i)?.Trim(),
                    KetTambahan = oDBDS_Detail.GetValue("U_T2_Ket_Tambahan", i)?.Trim(),
                    PajakPengganti = oDBDS_Detail.GetValue("U_T2_Pajak_Pengganti", i)?.Trim(),
                    Referensi = oDBDS_Detail.GetValue("U_T2_Referensi", i)?.Trim(),
                    Status = oDBDS_Detail.GetValue("U_T2_Status", i)?.Trim(),
                    KodeDokumenPendukung = oDBDS_Detail.GetValue("U_T2_Kode_Dok_Pendukung", i)?.Trim(),
                    BranchCode = oDBDS_Detail.GetValue("U_T2_Branch_Code", i)?.Trim(),
                    BranchName = oDBDS_Detail.GetValue("U_T2_Branch_Name", i)?.Trim(),
                    OutletCode = oDBDS_Detail.GetValue("U_T2_Outlet_Code", i)?.Trim(),
                    OutletName = oDBDS_Detail.GetValue("U_T2_Outlet_Name", i)?.Trim(),
                    AddInfo = oDBDS_Detail.GetValue("U_T2_Add_Info", i)?.Trim(),
                    BuyerCountry = oDBDS_Detail.GetValue("U_T2_Buyer_Country", i)?.Trim(),
                    BuyerEmail = oDBDS_Detail.GetValue("U_T2_Buyer_Email", i)?.Trim(),
                };

                listData.Add(detail);
            }

            return listData;
        }

        private static double SafeToDouble(string val)
        {
            return double.TryParse(val?.Trim(), out var result) ? result : 0;
        }

        public static List<TaxInvoice> BuildTaxInvoiceList(List<InvoiceDataModel> invoiceDatas)
        {
            var result = new List<TaxInvoice>();
            try
            {

                var groupedDocs = invoiceDatas
                    .GroupBy(g => new
                    {
                        g.DocEntry,
                        g.NoDocument,
                        g.InvDate,
                        g.JenisPajak,
                        g.Referensi,
                        g.SellerIDTKU,
                        g.NomorNPWP,
                        g.BuyerDocument,
                        g.BuyerCountry,
                        g.BPName,
                        g.NPWPAddress,
                        g.BuyerEmail,
                        g.BuyerIDTKU
                    });

                foreach (var grp in groupedDocs)
                {
                    var header = grp.Key;

                    var taxInv = new TaxInvoice
                    {
                        TaxInvoiceDate = header.InvDate,
                        TaxInvoiceOpt = "Normal",
                        TrxCode = header.JenisPajak,
                        AddInfo = header.JenisPajak == "04" ? "" : grp.First().AddInfo,
                        CustomDoc = "",
                        CustomDocMonthYear = "",
                        RefDesc = header.Referensi,
                        FacilityStamp = "",
                        SellerIDTKU = header.SellerIDTKU,
                        BuyerTin = header.NomorNPWP,
                        BuyerDocument = header.BuyerDocument,
                        BuyerCountry = header.BuyerCountry,
                        BuyerDocumentNumber = header.NomorNPWP,
                        BuyerName = header.BPName,
                        BuyerAdress = header.NPWPAddress,
                        BuyerEmail = header.BuyerEmail,
                        BuyerIDTKU = header.BuyerIDTKU,
                        ListOfGoodService = new ListOfGoodService
                        {
                            GoodServiceCollection = grp
                                .OrderBy(c => c.LineNum)
                                .Select(c => new GoodService
                                {
                                    Code = c.ItemCode,
                                    Name = c.ItemName,
                                    Unit = c.ItemUnit,
                                    Qty = c.Qty,
                                    Price = c.ItemPrice,
                                    TotalDiscount = c.TotalDisc,
                                    TaxBase = c.TaxBase,
                                    VATRate = c.VATRate,
                                    VAT = c.AmountVAT,
                                    STLGRate = c.STLGRate,
                                    STLG = c.STLG
                                }).ToList()
                        }
                    };

                    result.Add(taxInv);
                }

                return result;
            }
            catch (Exception)
            {

                throw;
            }
            
        }
    }
}
