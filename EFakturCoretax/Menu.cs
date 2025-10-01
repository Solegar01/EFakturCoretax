using EFakturCoretax.Helpers;
using EFakturCoretax.Models;
using EFakturCoretax.Services;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax
{
    class Menu
    {
        private Dictionary<string, bool> SelectedCkBox = new Dictionary<string, bool>()
            {
                { "OINV", false},
                { "ODPI", false},
                { "ORIN", false},
            };
        private Dictionary<string, string> addActions = new Dictionary<string, string>() 
        {
            {"1", "Add & View"},
            {"2", "Add & New"},
        };
        private string SelectedAddAction = "1";
        private List<FilterDataModel> FindListModel = new List<FilterDataModel>();
        private Dictionary<string, string> cbBranchValues = new Dictionary<string, string>();
        private Dictionary<string, string> cbOutletValues = new Dictionary<string, string>();
        DateTime? oldFromDt = null;
        DateTime? oldToDt = null;

        int oldFromDocEntry = 0;
        int oldToDocEntry = 0;

        string oldFromCust = string.Empty;
        string oldToCust = string.Empty;

        string oldFromBranch = string.Empty;
        string oldToBranch = string.Empty;

        string oldFromOutlet = string.Empty;
        string oldToOutlet = string.Empty;
        string strDocEntry = string.Empty;

        public void AddMenuItems()
        {
            SAPbouiCOM.Menus oMenus = null;
            SAPbouiCOM.MenuItem oMenuItem = null;

            oMenus = Application.SBO_Application.Menus;

            SAPbouiCOM.MenuCreationParams oCreationPackage = null;
            oCreationPackage = ((SAPbouiCOM.MenuCreationParams)(Application.SBO_Application.CreateObject(SAPbouiCOM.BoCreatableObjectType.cot_MenuCreationParams)));
            oMenuItem = Application.SBO_Application.Menus.Item("43520"); // moudles'

            oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_POPUP;
            oCreationPackage.UniqueID = "EFakturCoretax";
            oCreationPackage.String = "E-Faktur Coretax";
            oCreationPackage.Enabled = true;
            oCreationPackage.Position = -1;

            oMenus = oMenuItem.SubMenus;

            try
            {
                //  If the manu already exists this code will fail
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            {

            }

            try
            {
                // Get the menu collection of the newly added pop-up item
                oMenuItem = Application.SBO_Application.Menus.Item("EFakturCoretax");
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EFakturCoretax.MainForm";
                oCreationPackage.String = "Generate Coretax";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                var udfFormIds = new[]
                {
                    "133",     // AR Invoice
                    "65300",   // AR Down Payment
                    "179"      // AR Credit Memo
                };

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                    && !BusinessObjectInfo.BeforeAction
                    && udfFormIds.Contains(BusinessObjectInfo.FormTypeEx))
                {
                    try
                    {
                        SAPbouiCOM.Form oForm = null;

                        for (int i = 0; i < Application.SBO_Application.Forms.Count; i++)
                        {
                            SAPbouiCOM.Form form = Application.SBO_Application.Forms.Item(i);
                            string formType = form.TypeEx;
                            string formUID = form.UniqueID;

                            if (formType == $"-{BusinessObjectInfo.FormTypeEx}")
                            {
                                oForm = form;
                            }
                        }

                        if (oForm != null)
                        {
                            oForm.Items.Item("U_T2_Exported").Enabled = false;
                            oForm.Items.Item("U_T2_Coretax_No").Enabled = false;
                            var ckRevise = (SAPbouiCOM.CheckBox)oForm.Items.Item("U_T2_Revise_Coretax").Specific;
                            oForm.Items.Item("U_T2_Revise_Coretax").Enabled = true;
                            if (ckRevise.Checked)
                            {
                                oForm.Items.Item("U_T2_Revise_Coretax").Enabled = false;
                            }
                            else
                            {
                                var ckExp = (SAPbouiCOM.CheckBox)oForm.Items.Item("U_T2_Exported").Specific;
                                var txtCoretax = (SAPbouiCOM.EditText)oForm.Items.Item("U_T2_Coretax_No").Specific;
                                if (!ckExp.Checked && string.IsNullOrEmpty(txtCoretax.Value))
                                {
                                    oForm.Items.Item("U_T2_Revise_Coretax").Enabled = false;
                                }
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        // UDF form not available (not open yet)
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                && !BusinessObjectInfo.BeforeAction
                && BusinessObjectInfo.FormTypeEx == "EFakturCoretax.MainForm"
                )
                {
                    string formUID = BusinessObjectInfo.FormUID;

                    System.Threading.Tasks.Task.Run(async () =>
                    {
                        await System.Threading.Tasks.Task.Delay(200); // small delay
                        try
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(formUID);
                            if (oForm != null && oForm.Visible)
                            {
                                try
                                {
                                    FormHelper.StartLoading(oForm, "Loading...", 0, false);

                                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                                    SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                                    var status = oDBDS_Header.GetValue("Status", 0).Trim();

                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    {
                                        var listData = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);
                                        if (listData.Any())
                                        {
                                            var gInvList = listData
                                                        .Where(p => !string.IsNullOrEmpty(p.DocEntry)) // ✅ filter out null/empty DocEntry
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


                                            FindListModel = gInvList;
                                            SetMtFind(oForm); // load matrix
                                        }
                                        else
                                        {
                                            FormHelper.ClearMatrix(oForm, "MtFind", "DT_FILTER");
                                        }

                                        if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                                            oForm.Items.Item("CbStatus").Visible = false;
                                        
                                        oForm.Items.Item("CbSeries").Enabled = false;

                                        oForm.Items.Item("TStatus").Visible = true;

                                        ((SAPbouiCOM.EditText)oForm.Items.Item("TStatus").Specific).Value = status == "O" ? "Open" : "Closed";

                                        oForm.Items.Item("TDocNum").Enabled = false;

                                        ShowFilterGroup(oForm);

                                        SetMtGenerate(oForm);

                                        FormHelper.SetVisible(oForm, new[] { "CkAllDt", "CkAllDoc", "CkAllCust", "CkAllBr", "CkAllOtl" }, true);
                                        if (status == "C")
                                        {
                                            FormHelper.SetEnabled(oForm, new[] { "CkInv", "CkDp", "CkCm" }, false);
                                            FormHelper.SetEnabled(oForm, new[] { "BtFilter", "BtGen" }, false);
                                            FormHelper.SetEnabled(oForm, new[] { "CkAllDt", "CkAllDoc", "CkAllCust", "CkAllBr", "CkAllOtl" }, false);
                                            FormHelper.SetEnabled(oForm, new[] { "TFromDt", "TToDt", "TFromDoc", "TToDoc", "TFromCust", "TToCust", "CbFromBr", "CbToBr", "CbFromOtl", "CbToOtl" }, false);
                                        }

                                        oForm.Items.Item("TPostDate").Enabled = status == "O";
                                        oForm.Items.Item("BtCSV").Enabled = true;
                                        oForm.Items.Item("BtXML").Enabled = true;
                                        if (FormHelper.ItemIsExists(oForm, "BtCbAdd"))
                                        {
                                            oForm.Items.Item("BtCbAdd").Visible = false;
                                        }
                                    }
                                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                    {
                                        NewForm(oForm);
                                    }
                                }
                                catch (Exception)
                                {

                                    throw;
                                }
                                finally
                                {
                                    FormHelper.FinishLoading(oForm);
                                }
                            }
                        }
                        catch { }
                    });
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                && BusinessObjectInfo.BeforeAction
                && BusinessObjectInfo.FormTypeEx == "EFakturCoretax.MainForm")
                {
                    try
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                            if (oDBDS_Detail.Size == 1)
                            {
                                string docEntry = oDBDS_Detail.GetValue("U_T2_DocEntry", 0).Trim();
                                if (string.IsNullOrEmpty(docEntry))
                                {
                                    // Clear Detail
                                    while (oDBDS_Detail.Size > 0)
                                    {
                                        oDBDS_Detail.RemoveRecord(0);
                                    }
                                    oDBDS_Detail.Clear();  // extra reset
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }

                if (BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD
                && BusinessObjectInfo.ActionSuccess
                && !BusinessObjectInfo.BeforeAction
                && BusinessObjectInfo.FormTypeEx == "EFakturCoretax.MainForm")
                {
                    // This is the key of the new object (DocEntry)
                    string objectKeyXml = BusinessObjectInfo.ObjectKey;

                    // Parse XML to get DocEntry
                    var xmlDoc = new System.Xml.XmlDocument();
                    xmlDoc.LoadXml(objectKeyXml);

                    strDocEntry = xmlDoc.SelectSingleNode("//DocEntry")?.InnerText;
                    
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "FormDataEvent error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "EFakturCoretax.MainForm")
                {
                    MainForm activeForm = new MainForm();
                    activeForm.Show();
                }

                if (!pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    if (oForm.TypeEx == "EFakturCoretax.MainForm")
                    {
                       if (pVal.MenuUID == "1281") // Find Mode
                        {
                            try
                            {
                                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                                GetDataSeriesComboBox(oForm);
                                oForm.Items.Item("TDocNum").Enabled = true;
                                oForm.Items.Item("TPostDate").Enabled = true;

                                var statusItem = oForm.Items.Item("TStatus");
                                statusItem.Visible = false;
                                var cbSeriesItem = oForm.Items.Item("CbSeries");
                                cbSeriesItem.Enabled = true;
                                
                                FormHelper.RemoveFocus(oForm);
                                if (!FormHelper.ItemIsExists(oForm, "CbStatus"))
                                {
                                    var selectedStatus = oDBDS_Header.GetValue("Status", 0).Trim();

                                    SAPbouiCOM.Item oNewItem = oForm.Items.Add("CbStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                                    oNewItem.Left = statusItem.Left;
                                    oNewItem.Top = statusItem.Top;
                                    oNewItem.Width = statusItem.Width;
                                    oNewItem.Height = statusItem.Height;
                                    oNewItem.DisplayDesc = true;

                                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                                    oCombo.DataBind.SetBound(true, "@T2_CORETAX", "Status");
                                    // For FIND mode, better use unbound combo
                                    oCombo.ValidValues.Add("", "");
                                    oCombo.ValidValues.Add("O", "Open");
                                    oCombo.ValidValues.Add("C", "Closed");
                                    oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    var cbItem = oForm.Items.Item("CbStatus");
                                    cbItem.Visible = true;
                                    var oCombo = (SAPbouiCOM.ComboBox)cbItem.Specific;
                                    oCombo.Select("", SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }

                                if (FormHelper.ItemIsExists(oForm, "BtCbAdd"))
                                {
                                    oForm.Items.Item("BtCbAdd").Visible = false;
                                }
                                oForm.Items.Item("1").Visible = true;

                                FormHelper.SetVisible(oForm, new[] { "CkAllDt", "CkAllDoc", "CkAllCust", "CkAllBr", "CkAllOtl" }, false);
                                FormHelper.SetEnabled(oForm, new[] { "TFromDt","TToDt","TFromDoc","TToDoc","TFromCust","TToCust","CbFromBr","CbToBr","CbFromOtl","CbToOtl"}, true);
                                FormHelper.SetEnabled(oForm, new[] { "CkInv", "CkDp", "CkCm"}, true);
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            finally
                            {
                                FormHelper.FinishLoading(oForm);
                            }
                        }
                        else if (pVal.MenuUID == "1282" ) // Add
                        {
                            try
                            {
                                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                                NewForm(oForm);
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            finally
                            {
                                FormHelper.FinishLoading(oForm);
                            }
                        }
                        else if (pVal.MenuUID == "1284" || pVal.MenuUID == "1286") // Update, Cancel
                        {
                            try
                            {
                                FormHelper.StartLoading(oForm, "Loading...", 0, false);

                                FormHelper.RemoveFocus(oForm);
                                oForm.Items.Item("TDocNum").Enabled = false;
                                oForm.Items.Item("TStatus").Visible = true;
                                oForm.Items.Item("CbSeries").Enabled = false;
                                oForm.Items.Item("CbSeries").Visible = true;

                                if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                                    oForm.Items.Item("CbStatus").Visible = false;

                                if (FormHelper.ItemIsExists(oForm, "BtCbAdd"))
                                {
                                    oForm.Items.Item("BtCbAdd").Visible = false;
                                }

                                FormHelper.SetVisible(oForm, new[] { "CkAllDt", "CkAllDoc", "CkAllCust", "CkAllBr", "CkAllOtl" }, true);
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            finally
                            {
                                FormHelper.FinishLoading(oForm);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        private void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var udfFormIds = new[] { "-133", "-65300", "-179" };

            if (udfFormIds.Contains(pVal.FormTypeEx))
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                && !pVal.BeforeAction) // After form loaded
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    oForm.Items.Item("U_T2_Exported").Enabled = false;
                    oForm.Items.Item("U_T2_Coretax_No").Enabled = false;
                    oForm.Items.Item("U_T2_Revise_Coretax").Enabled = false;
                }

                if (pVal.ItemUID == "U_T2_Revise_Coretax"
                    && pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                    && pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    var ckRevise = (SAPbouiCOM.CheckBox)oForm.Items.Item("U_T2_Revise_Coretax").Specific;
                    var ckExp = (SAPbouiCOM.CheckBox)oForm.Items.Item("U_T2_Exported").Specific;
                    var txtCoretax = (SAPbouiCOM.EditText)oForm.Items.Item("U_T2_Coretax_No").Specific;
                    if (ckRevise.Checked) // sudah checked → jangan boleh uncheck
                    {
                        ckExp.Checked = false;
                        txtCoretax.Value = "";
                        FormHelper.RemoveFocus(oForm);
                        oForm.Items.Item("U_T2_Revise_Coretax").Enabled = false;
                        oForm.Items.Item("U_T2_Exported").Enabled = false;
                        oForm.Items.Item("U_T2_Coretax_No").Enabled = false;
                    }
                }
            }

            if (pVal.FormTypeEx == "EFakturCoretax.MainForm")
            {
                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    try
                    {

                        FormHelper.StartLoading(oForm,"Loading...", 0, false);

                        foreach (var key in SelectedCkBox.Keys.ToList())
                        {
                            SelectedCkBox[key] = false;
                        }
                        FindListModel.Clear();
                        cbBranchValues = new Dictionary<string, string>();
                        cbOutletValues = new Dictionary<string, string>();
                        oldFromDt = null;
                        oldToDt = null;
                        oldFromDocEntry = 0;
                        oldToDocEntry = 0;
                        oldFromCust = string.Empty;
                        oldToCust = string.Empty;
                        oldFromBranch = string.Empty;
                        oldToBranch = string.Empty;
                        oldFromOutlet = string.Empty;
                        oldToOutlet = string.Empty;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        FormHelper.FinishLoading(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_KEY_DOWN && !pVal.BeforeAction)
                {
                    // Check if ENTER key pressed
                    if (pVal.CharPressed == 13) // 13 = ENTER key
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

                        var oButton = (SAPbouiCOM.Button)oForm.Items.Item("1").Specific;
                        oButton.Item.Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE &&
                pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    //SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    try
                    {
                        FormHelper.StartLoading(oForm, "Loading...", 0, false);
                        
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                        string docEntry = oDBDS_Header.GetValue("DocEntry", 0).Trim();

                        if (oForm.Items.Count > 0)
                        {
                            oForm.Items.Item("LDisp").TextStyle = (int)(SAPbouiCOM.BoFontStyle.fs_Bold | SAPbouiCOM.BoFontStyle.fs_Underline);
                            oForm.Items.Item("LRange").TextStyle = (int)(SAPbouiCOM.BoFontStyle.fs_Bold | SAPbouiCOM.BoFontStyle.fs_Underline);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                NewForm(oForm);
                            }
                            else
                            {
                                GetDataSeriesComboBox(oForm);
                                oForm.Items.Item("CbSeries").Enabled = false;

                                string docNum = oDBDS_Header.GetValue("DocNum", 0)?.Trim();
                                string status = oDBDS_Header.GetValue("Status", 0)?.Trim();
                                ((SAPbouiCOM.EditText)oForm.Items.Item("TDocNum").Specific).Value = docNum;
                                ((SAPbouiCOM.EditText)oForm.Items.Item("TStatus").Specific).Value = status == "O" ? "Open" : "Closed";
                            }
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        FormHelper.FinishLoading(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED)
                {
                    if (pVal.BeforeAction) // before the press
                    {
                        if (pVal.ItemUID == "BtCbAdd")
                        {
                            int answer = Application.SBO_Application.MessageBox(
                                                "Adding this document is irreversible. Do you want to continue?",
                                                2, // two buttons
                                                "Yes", "No"
                                            );
                            // MessageBox returns 1 for first button, 2 for second
                            if (answer == 2)
                            {
                                BubbleEvent = false; // cancel the Add
                                return;
                            }
                        }
                        if (pVal.ItemUID == "1")
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                            if (oForm.Mode != SAPbouiCOM.BoFormMode.fm_FIND_MODE && oForm.Mode != SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                if (oForm.Items.Count > 0)
                                {
                                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                                    SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                                    if (string.IsNullOrEmpty(oDBDS_Header.GetValue("U_T2_Posting_Date", 0)?.Trim()))
                                    {
                                        Application.SBO_Application.StatusBar.SetText(
                                            "Posting date is required.",
                                            SAPbouiCOM.BoMessageTime.bmt_Short,
                                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                                        );
                                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                        {
                                            oForm.Items.Item("1").Visible = false;
                                            oForm.Items.Item("BtCbAdd").Visible = true;
                                        }
                                        BubbleEvent = false;
                                        return;
                                    }
                                    if (oDBDS_Detail.Size <= 0)
                                    {
                                        Application.SBO_Application.StatusBar.SetText(
                                            "Detail generated data is required.",
                                            SAPbouiCOM.BoMessageTime.bmt_Short,
                                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                                        );
                                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                                        {
                                            oForm.Items.Item("1").Visible = false;
                                            oForm.Items.Item("BtCbAdd").Visible = true;
                                        }
                                        BubbleEvent = false;
                                        return;
                                    }
                                }
                            }
                        }
                    }
                    else if (!pVal.BeforeAction)
                    {
                        //Check Box
                        if (new[] { "CkInv", "CkDp", "CkCm" }.Contains(pVal.ItemUID))
                        {
                            DocCheckBoxHandler(pVal, FormUID);
                        }

                        if (new[] { "CkAllDt", "CkAllDoc", "CkAllCust", "CkAllBr", "CkAllOtl" }.Contains(pVal.ItemUID))
                        {
                            CheckAllHandler(pVal, FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtFilter")
                        {
                            BtnFilterHandler(pVal, FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtGen")
                        {
                            BtnGenHandler(pVal, FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtXML")
                        {
                            ExportToXML(FormUID);
                        }

                        //
                        if (pVal.ItemUID == "BtCSV")
                        {
                            ExportToCsv(FormUID);
                        }

                        if (pVal.ItemUID == "BtCbAdd")
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                            oForm.Items.Item("1").Visible = true;
                            oForm.Items.Item("BtCbAdd").Visible = false;
                            oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
                        }

                        if (pVal.ItemUID == "1")
                        {
                            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(pVal.FormUID);
                            try
                            {
                                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                                if (oForm.Items.Count > 0)
                                {
                                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && pVal.ActionSuccess)
                                    {
                                        if (SelectedAddAction == "2")
                                        {
                                            NewForm(oForm);
                                            oForm.Items.Item("BtCbAdd").Visible = true;
                                        }
                                        else
                                        {
                                            oForm.Items.Item("BtCbAdd").Visible = false;
                                            RefreshCurrentForm(oForm, strDocEntry);
                                        }
                                    }
                                    else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                                    {
                                        oForm.Items.Item("BtCSV").Enabled = true;
                                        oForm.Items.Item("BtXML").Enabled = true;
                                    }
                                }
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            finally
                            {
                                FormHelper.FinishLoading(oForm);
                            }
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                    if (pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "TFromDt")
                        {
                            string strDate = oDBDS_Header.GetValue("U_T2_From_Date", 0).Trim();
                            if (DateTime.TryParseExact(strDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate)) oldFromDt = parsedDate;
                        }
                        if (pVal.ItemUID == "TToDt")
                        {
                            string strDate = oDBDS_Header.GetValue("U_T2_To_Date", 0).Trim();
                            if (DateTime.TryParseExact(strDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate)) oldToDt = parsedDate;
                        }
                        if (pVal.ItemUID == "TFromDoc")
                        {
                            string strDocEntry = oDBDS_Header.GetValue("U_T2_From_Doc_Entry", 0).Trim();
                            if (int.TryParse(strDocEntry, out int parsedVal)) oldFromDocEntry = parsedVal;
                        }
                        if (pVal.ItemUID == "TToDoc")
                        {
                            string strDocEntry = oDBDS_Header.GetValue("U_T2_To_Doc_Entry", 0).Trim();
                            if (int.TryParse(strDocEntry, out int parsedVal)) oldToDocEntry = parsedVal;
                        }
                        if (pVal.ItemUID == "TFromCust")
                        {
                            oldFromCust = oDBDS_Header.GetValue("U_T2_From_Cust", 0).Trim();
                        }
                        if (pVal.ItemUID == "TToCust")
                        {
                            oldToCust = oDBDS_Header.GetValue("U_T2_To_Cust", 0).Trim();
                        }
                    }
                    else if (!pVal.BeforeAction)
                    {
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                        {
                            FormHelper.SetEnabled(oForm, new[] { "BtCSV", "BtXML" }, true);
                        }
                        else
                        {
                            FormHelper.SetEnabled(oForm, new[] { "BtCSV", "BtXML" }, false);
                        }

                        if (pVal.ItemUID == "TFromDt" || pVal.ItemUID == "TToDt")
                        {
                            string strFromDate = oDBDS_Header.GetValue("U_T2_From_Date", 0).Trim();
                            string strToDate = oDBDS_Header.GetValue("U_T2_To_Date", 0).Trim();
                            DateTime? newFromDate = null;
                            DateTime? newToDate = null;
                            if (DateTime.TryParseExact(strFromDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate)) newFromDate = parsedDate;
                            if (DateTime.TryParseExact(strToDate, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedToDate)) newToDate = parsedToDate;
                            if (pVal.ItemUID == "TFromDt")
                            {
                                if ((oldFromDt != newFromDate))
                                {
                                    ResetDetail(oForm);
                                }
                            }
                            else
                            {
                                if ((oldToDt != newToDate))
                                {
                                    ResetDetail(oForm);
                                }
                            }
                            // Checkbox logic
                            if (newFromDate == null && newToDate == null)
                                FormHelper.SetValueDS(oForm, "CkDtDS", "Y");
                            else
                                FormHelper.SetValueDS(oForm, "CkDtDS", "N");
                        }

                        if (pVal.ItemUID == "TFromDoc" || pVal.ItemUID == "TToDoc")
                        {
                            string strFromEntry = oDBDS_Header.GetValue("U_T2_From_Doc_Entry", 0).Trim();
                            string strToEntry = oDBDS_Header.GetValue("U_T2_To_Doc_Entry", 0).Trim();
                            int newFromEntry = 0;
                            int newToEntry = 0;
                            if (int.TryParse(strFromEntry, out int parsedFromEntry)) newFromEntry = parsedFromEntry;
                            if (int.TryParse(strToEntry, out int parsedToEntry)) newToEntry = parsedToEntry;
                            if (pVal.ItemUID == "TFromDoc")
                            {
                                if ((oldFromDocEntry != newFromEntry))
                                {
                                    ResetDetail(oForm);
                                }
                            }
                            else
                            {
                                if ((oldToDocEntry != newToEntry))
                                {
                                    ResetDetail(oForm);
                                }
                            }

                            // Checkbox logic
                            if (newFromEntry == 0 && newToEntry == 0)
                                FormHelper.SetValueDS(oForm, "CkDocDS", "Y");
                            else
                                FormHelper.SetValueDS(oForm, "CkDocDS", "N");
                        }

                        if (pVal.ItemUID == "TFromCust" || pVal.ItemUID == "TToCust")
                        {
                            string newFromCust = oDBDS_Header.GetValue("U_T2_From_Cust", 0).Trim();
                            string newToCust = oDBDS_Header.GetValue("U_T2_To_Cust", 0).Trim();
                            if (pVal.ItemUID == "TFromCust")
                            {
                                if ((oldFromCust != newFromCust))
                                {
                                    ResetDetail(oForm);
                                }
                            }
                            else
                            {
                                if ((oldToCust != newToCust))
                                {
                                    ResetDetail(oForm);
                                }
                            }

                            // Checkbox logic
                            if (string.IsNullOrEmpty(newFromCust) && string.IsNullOrEmpty(newToCust))
                                FormHelper.SetValueDS(oForm, "CkCustDS", "Y");
                            else
                                FormHelper.SetValueDS(oForm, "CkCustDS", "N");
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    try
                    {
                        FormHelper.StartLoading(oForm, "Loading...", 0, false);
                        if (oForm.Items.Count > 0)
                        {
                            if (pVal.ItemUID == "MtFind" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false && pVal.ColUID == "Col_10")
                            {
                                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                                var status = oDBDS_Header.GetValue("Status", 0).Trim();
                                if (status != "O") return;
                                SelectFilterHandler(FormUID, pVal.Row);
                            }
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        FormHelper.FinishLoading(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    try
                    {
                        FormHelper.StartLoading(oForm, "Loading", 0, false);
                        if (new[] { "CbFromBr", "CbToBr", "CbFromOtl", "CbToOtl" }.Contains(pVal.ItemUID))
                        {
                            SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                            if (pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "CbFromBr")
                                {
                                    oldFromBranch = oDBDS_Header.GetValue("U_T2_From_Branch", 0).Trim();
                                }
                                if (pVal.ItemUID == "CbToBr")
                                {
                                    oldToBranch = oDBDS_Header.GetValue("U_T2_To_Branch", 0).Trim();
                                }
                                if (pVal.ItemUID == "CbFromOtl")
                                {
                                    oldFromOutlet = oDBDS_Header.GetValue("U_T2_From_Outlet", 0).Trim();
                                }
                                if (pVal.ItemUID == "CbToOtl")
                                {
                                    oldFromOutlet = oDBDS_Header.GetValue("U_T2_To_Outlet", 0).Trim();
                                }
                            }
                            else if (!pVal.BeforeAction)
                            {
                                if (pVal.ItemUID == "CbFromBr" || pVal.ItemUID == "CbToBr")
                                {
                                    string newFromBranch = oDBDS_Header.GetValue("U_T2_From_Branch", 0).Trim();
                                    string newToBranch = oDBDS_Header.GetValue("U_T2_To_Branch", 0).Trim();
                                    if ((oldFromBranch != newFromBranch) || (oldToBranch != newToBranch))
                                    {
                                        ResetDetail(oForm);
                                    }

                                    // Checkbox logic
                                    if (string.IsNullOrEmpty(newFromBranch) && string.IsNullOrEmpty(newToBranch))
                                        FormHelper.SetValueDS(oForm, "CkBrDS", "Y");
                                    else
                                        FormHelper.SetValueDS(oForm, "CkBrDS", "N");
                                }
                                if (pVal.ItemUID == "CbFromOtl" || pVal.ItemUID == "CbToOtl")
                                {
                                    string newFromOutlet = oDBDS_Header.GetValue("U_T2_From_Outlet", 0).Trim();
                                    string newToOutlet = oDBDS_Header.GetValue("U_T2_To_Outlet", 0).Trim();
                                    if ((oldFromOutlet != newFromOutlet) || (oldToOutlet != newToOutlet))
                                    {
                                        ResetDetail(oForm);
                                    }

                                    // Checkbox logic
                                    if (string.IsNullOrEmpty(newFromOutlet) && string.IsNullOrEmpty(newToOutlet))
                                        FormHelper.SetValueDS(oForm, "CkOtlDS", "Y");
                                    else
                                        FormHelper.SetValueDS(oForm, "CkOtlDS", "N");
                                }
                            }

                        }

                        if (pVal.ItemUID == "CbSeries" && !pVal.BeforeAction)
                        {
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                                int seriesId = int.Parse(oDBDS_Header.GetValue("Series", 0)?.Trim());
                                int nextDocNum = QueryHelper.GetLastDocNum(seriesId);
                                oDBDS_Header.SetValue("DocNum", 0, nextDocNum.ToString());
                            }
                        }

                        if (pVal.ItemUID == "BtCbAdd" && !pVal.BeforeAction)
                        {
                            var btCb = (SAPbouiCOM.ButtonCombo)oForm.Items.Item("BtCbAdd").Specific;
                            SelectedAddAction = btCb.Selected.Value;
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        FormHelper.FinishLoading(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE &&
                    pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    try
                    {
                        FormHelper.StartLoading(oForm, "Loading...", 0, false);
                        if (oForm.Items.Count > 0)
                        {
                            AdjustMatrix(oForm);
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        FormHelper.FinishLoading(oForm);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                {
                    CflHandler(pVal, FormUID);
                }
            }
        }

        private void NewForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                oDBDS_Header.Clear();
                oDBDS_Header.InsertRecord(0);
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                oDBDS_Detail.Clear();

                foreach (var key in SelectedCkBox.Keys.ToList())
                {
                    SelectedCkBox[key] = false;
                }
                FindListModel.Clear();
                cbBranchValues = new Dictionary<string, string>();
                cbOutletValues = new Dictionary<string, string>();
                oldFromDt = null;
                oldToDt = null;
                oldFromDocEntry = 0;
                oldToDocEntry = 0;
                oldFromCust = string.Empty;
                oldToCust = string.Empty;
                oldFromBranch = string.Empty;
                oldToBranch = string.Empty;
                oldFromOutlet = string.Empty;
                oldToOutlet = string.Empty;

                ShowFilterGroup(oForm);
                ResetDetail(oForm);

                var seriesId = QueryHelper.GetSeriesIdCoretax();
                var nextDocNum = QueryHelper.GetLastDocNum(seriesId).ToString();
                oDBDS_Header.SetValue("DocNum", 0, nextDocNum);

                // Update DB DataSource
                oDBDS_Header.SetValue("Series", 0, seriesId.ToString());
                oDBDS_Header.SetValue("Status", 0, "O");

                // Set display fields (unbound helper fields)
                GetDataSeriesComboBox(oForm);
                ((SAPbouiCOM.EditText)oForm.Items.Item("TStatus").Specific).Value = "Open";

                FormHelper.SetEnabled(oForm, new[] { "CkInv", "CkDp", "CkCm" }, true);
                FormHelper.SetVisible(oForm, new[] { "CkAllDt", "CkAllDoc", "CkAllCust", "CkAllBr", "CkAllOtl" }, true);
                FormHelper.SetEnabled(oForm, new[] { "TFromDt", "TToDt", "TFromDoc", "TToDoc", "TFromCust", "TToCust", "CbFromBr", "CbToBr", "CbFromOtl", "CbToOtl" }, false);
                if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                    oForm.Items.Item("CbStatus").Visible = false;
                oForm.Items.Item("CbSeries").Enabled = true;
                oForm.Items.Item("TDocNum").Enabled = false;
                oForm.Items.Item("TPostDate").Enabled = true;

                oForm.Items.Item("BtCSV").Enabled = false;
                oForm.Items.Item("BtXML").Enabled = false;

                var btAddItem = oForm.Items.Item("1");
                //btAddItem.Width = 1;
                if (!FormHelper.ItemIsExists(oForm, "BtCbAdd"))
                {
                    // Tambahkan item button combo
                    SAPbouiCOM.Item oItem = oForm.Items.Add("BtCbAdd", SAPbouiCOM.BoFormItemTypes.it_BUTTON_COMBO);
                    oItem.Left = btAddItem.Left;
                    oItem.Top = btAddItem.Top;
                    oItem.Width = btAddItem.Width;
                    oItem.Height = btAddItem.Height;
                    oItem.DisplayDesc = true;

                    // Cast ke ButtonCombo
                    SAPbouiCOM.ButtonCombo oBtnCombo = (SAPbouiCOM.ButtonCombo)oItem.Specific;
                    foreach (var item in addActions)
                    {
                        oBtnCombo.ValidValues.Add(item.Key, item.Value);
                    }
                    oBtnCombo.Select(SelectedAddAction, SAPbouiCOM.BoSearchKey.psk_ByValue);
                }
                else
                {
                    oForm.Items.Item("BtCbAdd").Visible = true;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void AdjustMatrix(SAPbouiCOM.Form oForm)
        {
            try
            {
                // Get form width and height
                int formWidth = oForm.ClientWidth;
                int formHeight = oForm.ClientHeight;

                SAPbouiCOM.Item mtx1 = oForm.Items.Item("MtFind");

                mtx1.Width = formWidth - 20;
                mtx1.Height = Convert.ToInt32(formHeight * 0.25);

                SAPbouiCOM.Item mtx2 = oForm.Items.Item("MtDetail");
                mtx2.Width = formWidth - 20;
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void SelectFilterHandler(string FormUID, int Row)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                if (FindListModel.Any())
                {
                    if (Row == 0)
                    {
                        ToggleSelectAll(oForm);
                    }
                    else
                    {
                        ToggleSelectSingle(oForm, Row);
                    }
                    SAPbouiCOM.Item btGen = oForm.Items.Item("BtGen");
                    if (FindListModel != null && FindListModel.Any((f) => f.Selected))
                    {
                        btGen.Enabled = true;
                    }
                    else
                    {
                        btGen.Enabled = false;
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                FormHelper.ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX_DT");
            }
        }

        private void ToggleSelectAll(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;
                SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Item("DT_FILTER");

                bool selectAll = oMatrix.Columns.Item("Col_10").TitleObject.Caption != "Unselect All";

                for (int i = 0; i < oDT.Rows.Count; i++)
                {
                    // update datatable column, e.g. U_Select
                    oDT.SetValue("Select", i, selectAll ? "Y" : "N");
                }
                foreach (var item in FindListModel)
                {
                    item.Selected = selectAll;
                }

                // reload once
                oMatrix.LoadFromDataSource();

                // update caption
                oMatrix.Columns.Item("Col_10").TitleObject.Caption = selectAll ? "Unselect All" : "Select All";
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        private void ToggleSelectSingle(SAPbouiCOM.Form oForm, int mtRow)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;

                // Get clicked row
                int row = mtRow - 1;

                // Get checkbox value (grid stores it as string "Y"/"N" or "tYES"/"tNO")
                bool isChecked = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_10").Cells.Item(mtRow).Specific).Checked;
                SAPbouiCOM.DataTable  oDT = oForm.DataSources.DataTables.Item("DT_FILTER");
                
                string docEntryVal = oDT.GetValue("DocEntry", row).ToString();
                string objTypeVal = oDT.GetValue("ObjType", row).ToString();
                var tempData = FindListModel.Where((f) => f.DocEntry == docEntryVal && f.ObjType == objTypeVal).FirstOrDefault();
                if (tempData != null)
                {
                    tempData.Selected = isChecked;
                }
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void CflHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                SAPbouiCOM.IChooseFromListEvent oCFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;
                if (SelectedCkBox.Where((ck) => ck.Value).Count() == 1)
                {
                    var selectedCk = SelectedCkBox.Where((ck) => ck.Value).First().Key;
                    if (oCFLEvent.ChooseFromListUID == "CflDocFrom" + selectedCk)
                    {
                        SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                        if (oDataTable != null && oDataTable.Rows.Count > 0)
                        {
                            // Get values from the selected row
                            string docNum = oDataTable.GetValue("DocNum", 0).ToString();
                            string docEntry = oDataTable.GetValue("DocEntry", 0).ToString();
                            string strDocEntry = oDBDS_Header.GetValue("DocEntry", 0).Trim();
                            int currDocEntry = 0;
                            int selectedDocEntry = 0;
                            if (int.TryParse(strDocEntry, out int parsedCurr)) currDocEntry = parsedCurr;
                            if (int.TryParse(docEntry, out int parsedSel)) selectedDocEntry = parsedSel;
                            if (currDocEntry != selectedDocEntry)
                            {
                                // Set to DBDataSource first
                                oDBDS_Header.SetValue("U_T2_From_Doc", 0, docNum);
                                oDBDS_Header.SetValue("U_T2_From_Doc_Entry", 0, docEntry);

                                // Then update edit text (only if it's editable)
                                var oEdit = (SAPbouiCOM.EditText)oForm.Items.Item("TFromDoc").Specific;
                                oEdit.Value = docNum;
                            }
                        }
                    }
                    if (oCFLEvent.ChooseFromListUID == "CflDocTo" + selectedCk)
                    {
                        SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                        if (oDataTable != null && oDataTable.Rows.Count > 0)
                        {
                            // Get values from the selected row
                            string docNum = oDataTable.GetValue("DocNum", 0).ToString();
                            string docEntry = oDataTable.GetValue("DocEntry", 0).ToString();
                            string strDocEntry = oDBDS_Header.GetValue("DocEntry", 0).Trim();
                            int currDocEntry = 0;
                            int selectedDocEntry = 0;
                            if (int.TryParse(strDocEntry, out int parsedCurr)) currDocEntry = parsedCurr;
                            if (int.TryParse(docEntry, out int parsedSel)) selectedDocEntry = parsedSel;
                            if (currDocEntry != selectedDocEntry)
                            {
                                oDBDS_Header.SetValue("U_T2_To_Doc_Entry", 0, docEntry);
                                oDBDS_Header.SetValue("U_T2_To_Doc", 0, docNum);

                                oForm.Items.Item("TToDoc").Update();
                            }
                        }
                    }
                }
                if (oCFLEvent.ChooseFromListUID == "CflCustFrom")
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;
                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    {
                        // Get values from the selected row
                        string code = oDataTable.GetValue("CardCode", 0).ToString();
                        string currCode = oDBDS_Header.GetValue("U_T2_From_Cust", 0).Trim();
                        if (code != currCode)
                        {
                            oDBDS_Header.SetValue("U_T2_From_Cust", 0, code);
                        }
                    }
                }
                if (oCFLEvent.ChooseFromListUID == "CflCustTo")
                {
                    SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;
                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    {
                        // Get values from the selected row
                        string code = oDataTable.GetValue("CardCode", 0).ToString();
                        string currCode = oDBDS_Header.GetValue("U_T2_To_Cust", 0).Trim();
                        if (currCode != code)
                        {
                            oDBDS_Header.SetValue("U_T2_To_Cust", 0, code);
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void SetMtGenerate(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtDetail").Specific;

                oMatrix.LoadFromDataSource();

                oMatrix.Columns.Item("DocEntry").Visible = false;
                oMatrix.Columns.Item("LineNum").Visible = false;
                oMatrix.Columns.Item("TIN").Visible = false;

                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    ((SAPbouiCOM.EditText)oMatrix.Columns.Item("#").Cells.Item(i).Specific).Value = i.ToString();
                    oMatrix.CommonSetting.SetRowBackColor(i, white);
                }

                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                oMatrix.AutoResizeColumns();
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void DocCheckBoxHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            if (oForm == null) return;

            try
            {
                FormHelper.StartLoading(oForm, "Loading...", 0, false);

                ShowFilterGroup(oForm);

                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE || oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                {
                    ResetFilters(oForm);
                    ResetDetail(oForm);

                }
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void CheckAllHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            try
            {
                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                var oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                if (pVal.ItemUID == "CkAllDt")
                {
                    oDBDS_Header.SetValue("U_T2_From_Date", 0, null);
                    oDBDS_Header.SetValue("U_T2_To_Date", 0, null);
                }
                if (pVal.ItemUID == "CkAllDoc")
                {
                    oDBDS_Header.SetValue("U_T2_From_Doc", 0, null);
                    oDBDS_Header.SetValue("U_T2_To_Doc", 0, null);
                }
                if (pVal.ItemUID == "CkAllCust")
                {
                    oDBDS_Header.SetValue("U_T2_From_Cust", 0, null);
                    oDBDS_Header.SetValue("U_T2_To_Cust", 0, null);
                }
                if (pVal.ItemUID == "CkAllBr")
                {
                    oDBDS_Header.SetValue("U_T2_From_Branch", 0, null);
                    oDBDS_Header.SetValue("U_T2_To_Branch", 0, null);
                }
                if (pVal.ItemUID == "CkAllOtl")
                {
                    oDBDS_Header.SetValue("U_T2_From_Outlet", 0, null);
                    oDBDS_Header.SetValue("U_T2_To_Outlet", 0, null);
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void ShowFilterGroup(SAPbouiCOM.Form oForm)
        {
            try
            {
                var oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                string[] dateItems = { "TFromDt", "TToDt", "CkAllDt" };
                string[] docItems = { "TFromDoc", "TToDoc", "CkAllDoc" };
                string[] custItems = { "TFromCust", "TToCust", "CkAllCust" };
                string[] brItems = { "CbFromBr", "CbToBr", "CkAllBr" };
                string[] otlItems = { "CbFromOtl", "CbToOtl", "CkAllOtl" };

                SelectedCkBox["OINV"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkInv").Specific).Checked;
                SelectedCkBox["ODPI"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkDp").Specific).Checked;
                SelectedCkBox["ORIN"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkCm").Specific).Checked;

                if (SelectedCkBox.ContainsValue(true))
                {
                    int countDoc = SelectedCkBox.Count(d => d.Value);

                    // Always enable date filters
                    FormHelper.SetEnabled(oForm, dateItems, true);

                    // Enable doc filter only if exactly 1 doc type
                    FormHelper.SetEnabled(oForm, docItems, countDoc == 1);
                    if (countDoc == 1)
                    {
                        string selectedDoc = SelectedCkBox.First(d => d.Value).Key;

                        if (!FormHelper.HasCfl(oForm, "CflDocFrom" + selectedDoc))
                            FormHelper.SetDocumentCfl(oForm, "CflDocFrom" + selectedDoc, "TFromDoc", selectedDoc);

                        if (!FormHelper.HasCfl(oForm, "CflDocTo" + selectedDoc))
                            FormHelper.SetDocumentCfl(oForm, "CflDocTo" + selectedDoc, "TToDoc", selectedDoc);
                    }
                    else
                    {
                        oDBDS_Header.SetValue("U_T2_From_Doc", 0, null);
                        oDBDS_Header.SetValue("U_T2_To_Doc", 0, null);
                        //FormHelper.ClearEdit(oForm, "TFromDoc");
                    }

                    // Always enable cust, branch, outlet
                    FormHelper.SetEnabled(oForm, custItems, true);
                    FormHelper.SetEnabled(oForm, brItems, true);
                    FormHelper.SetEnabled(oForm, otlItems, true);

                    // Customer CFLs
                    if (!FormHelper.HasCfl(oForm, "CflCustFrom"))
                        FormHelper.SetCustomerCfl(oForm, "CflCustFrom", "TFromCust");
                    if (!FormHelper.HasCfl(oForm, "CflCustTo"))
                        FormHelper.SetCustomerCfl(oForm, "CflCustTo", "TToCust");

                    // Lazy load branch combos
                    var cbFromBr = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbFromBr").Specific;
                    if (cbFromBr.ValidValues.Count < 2)
                    {
                        GetDataComboBox(oForm, "CbFromBr", "OBPL");
                        GetDataComboBox(oForm, "CbToBr", "OBPL");
                    }

                    // Lazy load outlet combos
                    var cbFromOtl = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbFromOtl").Specific;
                    if (cbFromOtl.ValidValues.Count < 2)
                    {
                        GetDataComboBox(oForm, "CbFromOtl", "OPRC");
                        GetDataComboBox(oForm, "CbToOtl", "OPRC");
                    }

                    oForm.Items.Item("BtFilter").Enabled = true;
                }
                else
                {
                    // Disable all
                    FormHelper.SetEnabled(oForm, dateItems, false);
                    FormHelper.SetEnabled(oForm, docItems, false);
                    FormHelper.SetEnabled(oForm, custItems, false);
                    FormHelper.SetEnabled(oForm, brItems, false);
                    FormHelper.SetEnabled(oForm, otlItems, false);

                    oForm.Items.Item("BtFilter").Enabled = false;
                }

                // Sync DS values with filters
                FormHelper.SetValueDS(oForm, "CkDtDS",
                    (!string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_From_Date", 0)?.Trim())
                        || !string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_To_Date", 0)?.Trim())) ? "N" : "Y");
                var fromDoc = oDBDS_Header.GetValue("U_T2_From_Doc", 0)?.Trim();
                var toDoc = oDBDS_Header.GetValue("U_T2_To_Doc", 0)?.Trim();
                FormHelper.SetValueDS(oForm, "CkDocDS",
                    (!string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_From_Doc", 0)?.Trim())
                        || !string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_To_Doc", 0)?.Trim())) ? "N" : "Y");

                FormHelper.SetValueDS(oForm, "CkCustDS",
                    (!string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_From_Cust", 0)?.Trim())
                        || !string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_To_Cust", 0)?.Trim())) ? "N" : "Y");

                FormHelper.SetValueDS(oForm, "CkBrDS",
                    (!string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_From_Branch", 0)?.Trim())
                        || !string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_To_Branch", 0)?.Trim())) ? "N" : "Y");

                FormHelper.SetValueDS(oForm, "CkOtlDS",
                    (!string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_From_Outlet", 0)?.Trim())
                        || !string.IsNullOrWhiteSpace(oDBDS_Header.GetValue("U_T2_To_Outlet", 0)?.Trim())) ? "N" : "Y");
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    "ShowFilterGroup error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

        private void ResetDetail(SAPbouiCOM.Form oForm)
        {
            if (oForm == null) return;

            try
            {
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                // Clear FindListModel
                if (FindListModel?.Count > 0)
                {
                    FindListModel.Clear();
                    FormHelper.ClearMatrix(oForm, "MtFind", "DT_FILTER");
                }

                // Clear Detail
                while (oDBDS_Detail.Size > 0)
                {
                    oDBDS_Detail.RemoveRecord(0);
                }
                oDBDS_Detail.Clear();  // extra reset
                FormHelper.ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX_DT");
                oForm.Items.Item("BtGen").Enabled = false;
            }
            catch (Exception)
            {
                throw;
            }
            
        }

        private void ResetFilters(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                // Dates
                oDBDS_Header.SetValue("U_T2_From_Date",0,null);
                oDBDS_Header.SetValue("U_T2_To_Date",0,null);
                oForm.DataSources.UserDataSources.Item("CkDtDS").Value = "Y";

                // Docs
                oDBDS_Header.SetValue("U_T2_From_Doc", 0, null);
                oDBDS_Header.SetValue("U_T2_To_Doc", 0, null);
                oForm.DataSources.UserDataSources.Item("CkDocDS").Value = "Y";

                // Customers
                oDBDS_Header.SetValue("U_T2_From_Cust", 0, null);
                oDBDS_Header.SetValue("U_T2_To_Cust", 0, null);
                oForm.DataSources.UserDataSources.Item("CkCustDS").Value = "Y";

                // Branches
                oDBDS_Header.SetValue("U_T2_From_Branch", 0, null);
                oDBDS_Header.SetValue("U_T2_To_Branch", 0, null);
                oForm.DataSources.UserDataSources.Item("CkBrDS").Value = "Y";

                // Outlets
                oDBDS_Header.SetValue("U_T2_From_Outlet", 0, null);
                oDBDS_Header.SetValue("U_T2_To_Outlet", 0, null);
                oForm.DataSources.UserDataSources.Item("CkOtlDS").Value = "Y";
            }
            catch (Exception)
            {

                throw;
            }
        }

        private void GetDataComboBox(SAPbouiCOM.Form form, string id, string table)
        {
            SAPbouiCOM.Item comboItem = form.Items.Item(id);
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)comboItem.Specific;

            //// Remove any default values
            //while (oCombo.ValidValues.Count > 0)
            //{
            //    oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            //}
            if (oCombo.ValidValues.Count <= 0)
            {
                // Add dropdown values
                if (table == "OBPL")
                {
                    if (cbBranchValues.Any())
                    {
                        oCombo.ValidValues.Add("", "");
                        foreach (var item in cbBranchValues)
                        {
                            oCombo.ValidValues.Add(item.Key, item.Value);
                        }
                    }
                    else
                    {
                        cbBranchValues = QueryHelper.GetDataCbBranch();
                        if (cbBranchValues.Any())
                        {
                            oCombo.ValidValues.Add("", "");
                            foreach (var item in cbBranchValues)
                            {
                                oCombo.ValidValues.Add(item.Key, item.Value);
                            }
                        }
                    }
                }
                if (table == "OPRC")
                {
                    if (cbOutletValues.Any())
                    {
                        oCombo.ValidValues.Add("", "");
                        foreach (var item in cbOutletValues)
                        {
                            oCombo.ValidValues.Add(item.Key, item.Value);
                        }
                    }
                    else
                    {
                        cbOutletValues = QueryHelper.GetDataCbOutlet();
                        if (cbOutletValues.Any())
                        {
                            oCombo.ValidValues.Add("", "");
                            foreach (var item in cbOutletValues)
                            {
                                oCombo.ValidValues.Add(item.Key, item.Value);
                            }
                        }
                    }
                }
            }
        }

        private void GetDataSeriesComboBox(SAPbouiCOM.Form oForm)
        {
            Dictionary<string, string> result = new Dictionary<string, string>();
            SAPbouiCOM.Item comboItem = oForm.Items.Item("CbSeries");
            
            SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)comboItem.Specific;

            // Remove any default values
            while (oCombo.ValidValues.Count > 0)
            {
                oCombo.ValidValues.Remove(0, SAPbouiCOM.BoSearchKey.psk_Index);
            }

            result = QueryHelper.GetDataSeries(oForm,"T2_CORETAX");
            if (result.Any())
            {
                foreach (var item in result)
                {
                    oCombo.ValidValues.Add(item.Key, item.Value);
                }
            }
        }

        private void RefreshCurrentForm(SAPbouiCOM.Form oForm, string docEntry = "")
        {
            var currMode = oForm.Mode;
            try
            {
                // get DocEntry before refreshing
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                string docEntryStr = oDBDS_Header.GetValue("DocEntry", 0).Trim();

                // switch to Find mode
                oForm.Mode = SAPbouiCOM.BoFormMode.fm_FIND_MODE;

                // set DocEntry value in key field (ItemUID normally = "DocEntry")
                if (!string.IsNullOrEmpty(docEntry))
                {
                    ((SAPbouiCOM.EditText)oForm.Items.Item("DocEntry").Specific).Value = docEntry;
                }
                else
                {
                    if (string.IsNullOrEmpty(docEntryStr))
                        return;
                    ((SAPbouiCOM.EditText)oForm.Items.Item("DocEntry").Specific).Value = docEntryStr;
                }
                
                // press Find (button "1" is Find/OK in SAP system & UDO forms)
                oForm.Items.Item("1").Click(SAPbouiCOM.BoCellClickType.ct_Regular);
            }
            catch (Exception ex)
            {
                oForm.Mode = currMode;
                Application.SBO_Application.StatusBar.SetText(
                    $"Refresh failed: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }


        private void ExportToXML(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            
            SAPbobsCOM.Company oCompany = null;

            try
            {
                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                oCompany = Services.CompanyService.GetCompany();
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                }

                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                
                var status = oDBDS_Header.GetValue("Status", 0)?.Trim();
                var listData = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);

                if (listData == null || !listData.Any())
                    throw new Exception("No data to Export.");

                var listOfTax = FormHelper.BuildTaxInvoiceList(listData);
                if (!listOfTax.Any())
                    throw new Exception("No data to Export.");

                var invoice = new TaxInvoiceBulk
                {
                    TIN = listData.First().TIN,
                    ListOfTaxInvoice = new ListOfTaxInvoice { TaxInvoiceCollection = listOfTax }
                };

                var confirmStr = status == "O" ? "Export will also close the document. " : "";
                int response = Application.SBO_Application.MessageBox(
                    $"{confirmStr}Are you sure you want to export this document?",
                    1, "Yes", "No", "");

                if (response != 1) return; // user pressed No

                FormHelper.StartLoading(oForm, "Exporting data to XML...", 0, false);

                // === business logic sync ===
                var docNum = oDBDS_Header.GetValue("DocNum", 0).Trim();
                var datas = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);
                var docEntry = int.Parse(oDBDS_Header.GetValue("DocEntry", 0).Trim());

                if (status == "O")
                {
                    TransactionService.CloseCoretax(oCompany, docEntry);
                    TransactionService.UpdateStatusInv(oCompany, docNum, datas);
                }
                else
                {
                    TransactionService.UpdateStatusInv(oCompany, docNum, datas);
                }

                if (ExportHelper.ExportXml(invoice))
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    RefreshCurrentForm(oForm);
                    Application.SBO_Application.StatusBar.SetText(
                        "Successfully exported to XML.",
                        SAPbouiCOM.BoMessageTime.bmt_Medium,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);

                }
                else
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    Application.SBO_Application.StatusBar.SetText(
                        "Failed to export data to XML.",
                        SAPbouiCOM.BoMessageTime.bmt_Medium,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                if (oCompany?.InTransaction == true)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                System.Threading.Tasks.Task.Run(() =>
                {
                    System.Threading.Thread.Sleep(200); // let dialog close
                    Application.SBO_Application.StatusBar.SetText(
                        ex.Message,
                        SAPbouiCOM.BoMessageTime.bmt_Short,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                });
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void ExportToCsv(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            
            SAPbobsCOM.Company oCompany = null;

            try
            {
                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                oCompany = Services.CompanyService.GetCompany();
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                }

                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");

                if (string.IsNullOrEmpty(oDBDS_Header.GetValue("U_T2_Posting_Date", 0)?.Trim()))
                {
                    throw new Exception("Posting date must be filled in before exporting data.");
                }

                var status = oDBDS_Header.GetValue("Status", 0)?.Trim();
                var listData = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);

                if (listData == null || !listData.Any())
                    throw new Exception("No data to Export.");

                var listOfTax = FormHelper.BuildTaxInvoiceList(listData);
                if (!listOfTax.Any())
                    throw new Exception("No data to Export.");

                var invoice = new TaxInvoiceBulk
                {
                    TIN = listData.First().TIN,
                    ListOfTaxInvoice = new ListOfTaxInvoice { TaxInvoiceCollection = listOfTax }
                };

                var confirmStr = status == "O" ? "Export will also close the document. " : "";
                int response = Application.SBO_Application.MessageBox(
                    $"{confirmStr}Are you sure you want to export this document?",
                    1, "Yes", "No", "");

                if (response != 1) return; // user pressed No

                FormHelper.StartLoading(oForm, "Exporting data to CSV...", 0, false);

                // === business logic sync ===
                var docNum = oDBDS_Header.GetValue("DocNum", 0).Trim();
                var datas = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);
                var docEntry = int.Parse(oDBDS_Header.GetValue("DocEntry", 0).Trim());

                if (status == "O")
                {
                    TransactionService.CloseCoretax(oCompany, docEntry);
                    TransactionService.UpdateStatusInv(oCompany, docNum, datas);
                }
                else
                {
                    TransactionService.UpdateStatusInv(oCompany, docNum, datas);
                }

                if (ExportHelper.ExportCsv(invoice))
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                    RefreshCurrentForm(oForm);
                    Application.SBO_Application.StatusBar.SetText(
                        "Successfully exported to CSV.",
                        SAPbouiCOM.BoMessageTime.bmt_Medium,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                }
                else
                {
                    if (oCompany.InTransaction)
                        oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                    Application.SBO_Application.StatusBar.SetText(
                        "Failed to export data to CSV.",
                        SAPbouiCOM.BoMessageTime.bmt_Medium,
                        SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                }
            }
            catch (Exception ex)
            {
                if (oCompany?.InTransaction == true)
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);

                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }

        private void BtnGenHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            //if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
            Task.Run(async () =>
            {
                try
                {
                    FormHelper.StartLoading(oForm, "Generating data...", 0, false);

                    SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                    while (oDBDS_Detail.Size > 0)
                    {
                        oDBDS_Detail.RemoveRecord(0);
                    }
                    oDBDS_Detail.Clear();  // extra reset
                    FormHelper.ClearMatrix(oForm, "MtDetail", "DT_DETAIL", "@T2_CORETAX_DT");
                    var filteredHeader = FindListModel.Where((f) => f.Selected).ToList();
                    if (filteredHeader != null && filteredHeader.Any())
                    {
                        var listDetail = await TransactionService.GetDataGenerate(filteredHeader);
                        if (listDetail.Any())
                        {
                            oDBDS_Detail.Clear(); // optional: clear old rows first

                            for (int i = 0; i < listDetail.Count; i++)
                            {
                                oDBDS_Detail.InsertRecord(i);

                                var detail = listDetail[i];

                                oDBDS_Detail.SetValue("U_T2_TIN", i, detail.TIN ?? "");
                                oDBDS_Detail.SetValue("U_T2_DocEntry", i, detail.DocEntry ?? "");
                                oDBDS_Detail.SetValue("U_T2_LineNum", i, detail.LineNum ?? "");
                                DateTime parsedDate;
                                if (DateTime.TryParseExact(detail.InvDate,
                                                           "dd/MM/yyyy HH.mm.ss",   // your format
                                                           CultureInfo.InvariantCulture,
                                                           DateTimeStyles.None,
                                                           out parsedDate))
                                {
                                    oDBDS_Detail.SetValue("U_T2_Inv_Date", i, parsedDate.ToString("yyyyMMdd"));
                                }
                                else
                                {
                                    // fallback if invalid
                                    oDBDS_Detail.SetValue("U_T2_Inv_Date", i, "");
                                }
                                oDBDS_Detail.SetValue("U_T2_No_Doc", i, detail.NoDocument ?? "");
                                oDBDS_Detail.SetValue("U_T2_Object_Type", i, detail.ObjectType ?? "");
                                oDBDS_Detail.SetValue("U_T2_Object_Name", i, detail.ObjectName ?? "");
                                oDBDS_Detail.SetValue("U_T2_BP_Code", i, detail.BPCode ?? "");
                                oDBDS_Detail.SetValue("U_T2_BP_Name", i, detail.BPName ?? "");
                                oDBDS_Detail.SetValue("U_T2_Seller_IDTKU", i, detail.SellerIDTKU ?? "");
                                oDBDS_Detail.SetValue("U_T2_Buyer_Doc", i, detail.BuyerDocument ?? "");
                                oDBDS_Detail.SetValue("U_T2_Nomor_NPWP", i, detail.NomorNPWP ?? "");
                                oDBDS_Detail.SetValue("U_T2_NPWP_Name", i, detail.NPWPName ?? "");
                                oDBDS_Detail.SetValue("U_T2_NPWP_Address", i, detail.NPWPAddress ?? "");
                                oDBDS_Detail.SetValue("U_T2_Buyer_IDTKU", i, detail.BuyerIDTKU ?? "");
                                oDBDS_Detail.SetValue("U_T2_Item_Code", i, detail.ItemCode ?? "");
                                oDBDS_Detail.SetValue("U_T2_Item_Name", i, detail.ItemName ?? "");
                                oDBDS_Detail.SetValue("U_T2_Item_Unit", i, detail.ItemUnit ?? "");

                                // --- decimals ---
                                oDBDS_Detail.SetValue("U_T2_Item_Price", i, detail.ItemPrice.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_Qty", i, detail.Qty.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_Total_Disc", i, detail.TotalDisc.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_Tax_Base", i, detail.TaxBase.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_Other_Tax_Base", i, detail.OtherTaxBase.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_VAT_Rate", i, detail.VATRate.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_Amount_VAT", i, detail.AmountVAT.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_STLG_Rate", i, detail.STLGRate.ToString(CultureInfo.InvariantCulture));
                                oDBDS_Detail.SetValue("U_T2_STLG", i, detail.STLG.ToString(CultureInfo.InvariantCulture));

                                // --- strings ---
                                oDBDS_Detail.SetValue("U_T2_Jenis_Pajak", i, detail.JenisPajak ?? "");
                                oDBDS_Detail.SetValue("U_T2_Ket_Tambahan", i, detail.KetTambahan ?? "");
                                oDBDS_Detail.SetValue("U_T2_Pajak_Pengganti", i, detail.PajakPengganti ?? "");
                                oDBDS_Detail.SetValue("U_T2_Referensi", i, detail.Referensi ?? "");
                                oDBDS_Detail.SetValue("U_T2_Status", i, detail.Status ?? "");
                                oDBDS_Detail.SetValue("U_T2_Kode_Dok_Pendukung", i, detail.KodeDokumenPendukung ?? "");
                                oDBDS_Detail.SetValue("U_T2_Branch_Code", i, detail.BranchCode ?? "");
                                oDBDS_Detail.SetValue("U_T2_Branch_Name", i, detail.BranchName ?? "");
                                oDBDS_Detail.SetValue("U_T2_Outlet_Code", i, detail.OutletCode ?? "");
                                oDBDS_Detail.SetValue("U_T2_Outlet_Name", i, detail.OutletName ?? "");
                                oDBDS_Detail.SetValue("U_T2_Add_Info", i, detail.AddInfo ?? "");
                                oDBDS_Detail.SetValue("U_T2_Buyer_Country", i, detail.BuyerCountry ?? "");
                                oDBDS_Detail.SetValue("U_T2_Buyer_Email", i, detail.BuyerEmail ?? "");
                            }
                        }
                        SetMtGenerate(oForm);
                    }
                }
                catch (Exception e)
                {

                    throw e;
                }
                finally
                {
                    FormHelper.FinishLoading(oForm);
                }
            });
        }

        private void BtnFilterHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
            
            Task.Run(async () => {
                try
                {
                    FormHelper.StartLoading(oForm, "Retrieving data...", 0, false);

                    ResetDetail(oForm);

                    bool allDt = false;
                    bool allDoc = false;
                    bool allCust = false;
                    bool allBranch = false;
                    bool allOutlet = false;
                    string dtFrom = string.Empty;
                    string dtTo = string.Empty;
                    string docFrom = string.Empty;
                    string docTo = string.Empty;
                    string custFrom = string.Empty;
                    string custTo = string.Empty;
                    string branchFrom = string.Empty;
                    string branchTo = string.Empty;
                    string outFrom = string.Empty;
                    string outTo = string.Empty;

                    if (oForm.Items.Item("CkAllDt").Enabled)
                    //if (ItemIsExists(oForm, "CkAllDt") && oForm.Items.Item("CkAllDt").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllDt").Specific;
                        allDt = oCk.Checked;
                    }
                    if (oForm.Items.Item("CkAllDoc").Enabled)
                    //if (ItemIsExists(oForm, "CkAllDoc") && oForm.Items.Item("CkAllDoc").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllDoc").Specific;
                        allDoc = oCk.Checked;
                    }
                    if (oForm.Items.Item("CkAllCust").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllCust").Specific;
                        allCust = oCk.Checked;
                    }
                    if (oForm.Items.Item("CkAllBr").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllBr").Specific;
                        allBranch = oCk.Checked;
                    }
                    if (oForm.Items.Item("CkAllOtl").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllOtl").Specific;
                        allOutlet = oCk.Checked;
                    }

                    if (!allDt)
                    {
                        if (oForm.Items.Item("TFromDt").Enabled)
                        {
                            string oDtFrom = oDBDS_Header.GetValue("U_T2_From_Date",0).Trim();
                            if (DateTime.TryParseExact(oDtFrom, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDtFrom))
                            {
                                dtFrom = parsedDtFrom.ToString("yyyy-MM-dd");
                            }
                        }
                        if (oForm.Items.Item("TToDt").Enabled)
                        {
                            string oDtTo = oDBDS_Header.GetValue("U_T2_To_Date", 0).Trim();
                            if (DateTime.TryParseExact(oDtTo, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDtTo))
                            {
                                dtTo = parsedDtTo.ToString("yyyy-MM-dd");
                            }
                        }
                    }
                    if (!allDoc)
                    {
                        if (oForm.Items.Item("TFromDoc").Enabled)
                        {
                            docFrom = oDBDS_Header.GetValue("U_T2_From_Doc_Entry", 0).Trim();
                        }
                        if (oForm.Items.Item("TToDoc").Enabled)
                        {
                            docTo = oDBDS_Header.GetValue("U_T2_To_Doc_Entry", 0).Trim();
                        }
                    }
                    if (!allCust)
                    {
                        if (oForm.Items.Item("TFromCust").Enabled)
                        {
                            custFrom = oDBDS_Header.GetValue("U_T2_From_Cust", 0).Trim();
                        }
                        if (oForm.Items.Item("TToCust").Enabled)
                        {
                            custTo = oDBDS_Header.GetValue("U_T2_To_Cust", 0).Trim();
                        }
                    }
                    if (!allBranch)
                    {
                        if (oForm.Items.Item("CbFromBr").Enabled)
                        {
                            branchFrom = oDBDS_Header.GetValue("U_T2_From_Branch",0).Trim();
                        }
                        if (oForm.Items.Item("CbToBr").Enabled)
                        {
                            branchTo = oDBDS_Header.GetValue("U_T2_To_Branch", 0).Trim();
                        }
                    }
                    if (!allOutlet)
                    {
                        if (oForm.Items.Item("CbFromOtl").Enabled)
                        {
                            outFrom = oDBDS_Header.GetValue("U_T2_From_Outlet", 0).Trim();
                        }
                        if (oForm.Items.Item("CbToOtl").Enabled)
                        {
                            outTo = oDBDS_Header.GetValue("U_T2_To_Outlet", 0).Trim();
                        }
                    }
                    FindListModel = await TransactionService.GetDataFilter(
                                SelectedCkBox, dtFrom, dtTo, docFrom, docTo, custFrom, custTo,
                                branchFrom, branchTo, outFrom, outTo
                                );
                    SetMtFind(oForm);
                    
                }
                catch (Exception)
                {

                    throw;
                }
                finally
                {
                    FormHelper.FinishLoading(oForm);
                }
            });
        }

        private void SetMtFind(SAPbouiCOM.Form oForm)
        {
            try
            {
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                var status = oDBDS_Header.GetValue("Status", 0).Trim();
                // Create DataTable if not exists
                SAPbouiCOM.DataTable oDT;
                bool selectAll = !FindListModel.Any((f) => !f.Selected);

                if (!FormHelper.DtIsExists(oForm, "DT_FILTER"))
                {
                    oDT = oForm.DataSources.DataTables.Add("DT_FILTER");
                }
                else
                {
                    oDT = oForm.DataSources.DataTables.Item("DT_FILTER");
                }
                // Clear previous rows
                oDT.Clear();

                // Also clear matrix (important)
                oMatrix.Clear();

                // Define all columns (make sure sizes are large enough)
                oDT.Columns.Add("Select", SAPbouiCOM.BoFieldsType.ft_AlphaNumeric, 1);
                oDT.Columns.Add("DocEntry", SAPbouiCOM.BoFieldsType.ft_Integer);
                oDT.Columns.Add("DocNum", SAPbouiCOM.BoFieldsType.ft_Integer);
                oDT.Columns.Add("ObjType", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("ObjName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("BPCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BPName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("PostDate", SAPbouiCOM.BoFieldsType.ft_Date);
                oDT.Columns.Add("BranchCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("BranchName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("OutletCode", SAPbouiCOM.BoFieldsType.ft_Text, 50);
                oDT.Columns.Add("OutletName", SAPbouiCOM.BoFieldsType.ft_Text, 100);
                oDT.Columns.Add("#", SAPbouiCOM.BoFieldsType.ft_Integer);

                oForm.Items.Item("BtGen").Enabled = FindListModel.Any((f) => f.Selected);

                // Fill DataTable from model
                for (int i = 0; i < FindListModel.Count; i++)
                {
                    var row = FindListModel[i];
                    oDT.Rows.Add();

                    // Convert date string to SAP date format
                    if (!string.IsNullOrEmpty(row.PostDate) && DateTime.TryParse(row.PostDate, out var invDate))
                        oDT.SetValue("PostDate", i, invDate.ToString("yyyyMMdd"));
                    else
                        oDT.SetValue("PostDate", i, "");

                    oDT.SetValue("Select", i, row.Selected ? "Y" : "N");
                    oDT.SetValue("DocEntry", i, row.DocEntry ?? "");
                    oDT.SetValue("DocNum", i, row.DocNo ?? "");
                    oDT.SetValue("BPCode", i, row.CardCode ?? "");
                    oDT.SetValue("BPName", i, row.CardName ?? "");
                    oDT.SetValue("ObjType", i, row.ObjType ?? "");
                    oDT.SetValue("ObjName", i, row.ObjName ?? "");
                    oDT.SetValue("BranchCode", i, row.BranchCode ?? "");
                    oDT.SetValue("BranchName", i, row.BranchName ?? "");
                    oDT.SetValue("OutletCode", i, row.OutletCode ?? "");
                    oDT.SetValue("OutletName", i, row.OutletName ?? "");
                    oDT.SetValue("#", i, (i + 1));
                }

                oMatrix.Columns.Item("Col_1").DataBind.Bind("DT_FILTER", "DocNum");
                oMatrix.Columns.Item("Col_1").Width = 80;
                oMatrix.Columns.Item("Col_2").DataBind.Bind("DT_FILTER", "BPCode");
                oMatrix.Columns.Item("Col_2").Width = 80;
                oMatrix.Columns.Item("Col_3").DataBind.Bind("DT_FILTER", "BPName");
                oMatrix.Columns.Item("Col_3").Width = 100;
                oMatrix.Columns.Item("Col_4").DataBind.Bind("DT_FILTER", "ObjName");
                oMatrix.Columns.Item("Col_4").Width = 100;
                oMatrix.Columns.Item("Col_5").DataBind.Bind("DT_FILTER", "PostDate");
                oMatrix.Columns.Item("Col_5").Width = 80;
                oMatrix.Columns.Item("Col_6").DataBind.Bind("DT_FILTER", "BranchCode");
                oMatrix.Columns.Item("Col_6").Width = 80;
                oMatrix.Columns.Item("Col_7").DataBind.Bind("DT_FILTER", "BranchName");
                oMatrix.Columns.Item("Col_7").Width = 100;
                oMatrix.Columns.Item("Col_8").DataBind.Bind("DT_FILTER", "OutletCode");
                oMatrix.Columns.Item("Col_8").Width = 80;
                oMatrix.Columns.Item("Col_9").DataBind.Bind("DT_FILTER", "OutletName");
                oMatrix.Columns.Item("Col_9").Width = 100;
                oMatrix.Columns.Item("Col_10").DataBind.Bind("DT_FILTER", "Select");
                oMatrix.Columns.Item("Col_10").Width = 40;
                oMatrix.Columns.Item("Col_10").TitleObject.Caption = selectAll ? "Unselect All" : "Select All";
                oMatrix.Columns.Item("Col_10").Editable = status == "O";
                oMatrix.Columns.Item("#").DataBind.Bind("DT_FILTER", "#");
                oMatrix.Columns.Item("#").Width = 30;

                // Load data into matrix
                oMatrix.LoadFromDataSource();
                oMatrix.SelectionMode = SAPbouiCOM.BoMatrixSelect.ms_Auto;
                int white = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.White);
                for (int i = 0; i < oMatrix.RowCount; i++)
                {
                    oMatrix.CommonSetting.SetRowBackColor(i + 1, white);
                }
                oMatrix.AutoResizeColumns();
            }
            catch (Exception e)
            {

                throw e;
            }
        }

    }
}
