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
        private SAPbouiCOM.ProgressBar _pb;
        private Dictionary<string, bool> SelectedCkBox = new Dictionary<string, bool>()
            {
                { "OINV", false},
                { "ODPI", false},
                { "ORIN", false},
            };
        private List<FilterDataModel> FindListModel = new List<FilterDataModel>();
        private bool FilterIsShow = false;
        private Dictionary<string, string> cbBranchValues = new Dictionary<string, string>();
        private Dictionary<string, string> cbOutletValues = new Dictionary<string, string>();
        DateTime? oldFromDt = null;
        DateTime? oldToDt = null;

        int oldFromDocEntry = 0;
        int oldToDocEntry = 0;

        string oldFromCust = null;
        string oldToCust = null;

        string oldFromBranch = null;
        string oldToBranch = null;

        string oldFromOutlet = null;
        string oldToOutlet = null;

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
            catch (Exception e)
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
            catch (Exception er)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        private void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if(BusinessObjectInfo.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_LOAD
                && !BusinessObjectInfo.BeforeAction
                && BusinessObjectInfo.ActionSuccess
                && BusinessObjectInfo.FormTypeEx == "EFakturCoretax.MainForm")
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(BusinessObjectInfo.FormUID);
                    try
                    {
                        if (_pb != null) { _pb.Stop();_pb = null; }
                        _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Data loading...", 0, false);
                        oForm.Freeze(true);
                        SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                        
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE && oDBDS_Detail.Size > 0)
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

                            SetMtGenerate(oForm);

                            FormHelper.SetEnabled(oForm, new[] { "BtCSV", "BtXML" }, true);
                        }
                        else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                        {
                            NewForm(oForm);
                            FormHelper.SetEnabled(oForm, new[] { "BtCSV", "BtXML" }, false);
                        }
                        else
                        {
                            FormHelper.SetEnabled(oForm, new[] { "BtCSV", "BtXML" }, false);
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        if (_pb != null) { _pb.Stop(); _pb = null; }
                        oForm.Freeze(false);
                    }
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
                                oForm.Freeze(true);
                                if (_pb != null) { _pb.Stop(); _pb = null; }
                                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                                oForm.Items.Item("TDocNum").Enabled = true;

                                var statusItem = oForm.Items.Item("TStatus");
                                statusItem.Visible = false;

                                FormHelper.RemoveFocus(oForm);
                                if (!FormHelper.ItemIsExists(oForm, "CbStatus"))
                                {
                                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                                    var selectedStatus = oDBDS_Header.GetValue("Status", 0).Trim();

                                    SAPbouiCOM.Item oNewItem = oForm.Items.Add("CbStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                                    oNewItem.Left = statusItem.Left;
                                    oNewItem.Top = statusItem.Top;
                                    oNewItem.Width = statusItem.Width;
                                    oNewItem.Height = statusItem.Height;

                                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                                    oCombo.DataBind.SetBound(true, "@T2_CORETAX", "Status");
                                    // For FIND mode, better use unbound combo
                                    oCombo.ValidValues.Add("", "");
                                    oCombo.ValidValues.Add("O", "Open");
                                    oCombo.ValidValues.Add("C", "Closed");
                                    //oCombo.Select(selectedStatus, SAPbouiCOM.BoSearchKey.psk_ByValue);
                                }
                                else
                                {
                                    oForm.Items.Item("CbStatus").Visible = true;
                                }

                                oForm.Items.Item("TFromDt").Enabled = true;
                                oForm.Items.Item("TToDt").Enabled = true;
                                oForm.Items.Item("CkAllDt").Visible = false;
                                oForm.Items.Item("TFromDoc").Enabled = true;
                                oForm.Items.Item("TToDoc").Enabled = true;
                                oForm.Items.Item("CkAllDoc").Visible = false;
                                oForm.Items.Item("TFromCust").Enabled = true;
                                oForm.Items.Item("TToCust").Enabled = true;
                                oForm.Items.Item("CkAllCust").Visible = false;
                                oForm.Items.Item("CbFromBr").Enabled = true;
                                oForm.Items.Item("CbToBr").Enabled = true;
                                oForm.Items.Item("CkAllBr").Visible = false;
                                oForm.Items.Item("CbFromOtl").Enabled = true;
                                oForm.Items.Item("CbToOtl").Enabled = true;
                                oForm.Items.Item("CkAllOtl").Visible = false;
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            finally
                            {
                                if (_pb != null) { _pb.Stop(); _pb = null; }
                                oForm.Freeze(false);
                            }
                        }
                        else if (pVal.MenuUID == "1282" ) // Add
                        {
                            NewForm(oForm);
                        }
                        else if (pVal.MenuUID == "1284" || pVal.MenuUID == "1286") // Update, Cancel
                        {
                            try
                            {
                                oForm.Freeze(true);
                                if (_pb != null) { _pb.Stop(); _pb = null; }
                                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                                FormHelper.RemoveFocus(oForm);
                                oForm.Items.Item("TDocNum").Enabled = false;
                                oForm.Items.Item("TStatus").Visible = true;

                                if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                                    oForm.Items.Item("CbStatus").Visible = false;

                                oForm.Items.Item("CkAllDt").Visible = true;
                                oForm.Items.Item("CkAllDoc").Visible = true;
                                oForm.Items.Item("CkAllCust").Visible = true;
                                oForm.Items.Item("CkAllBr").Visible = true;
                                oForm.Items.Item("CkAllOtl").Visible = true;
                            }
                            catch (Exception)
                            {

                                throw;
                            }
                            finally
                            {
                                if (_pb != null) { _pb.Stop(); _pb = null; }
                                oForm.Freeze(false);
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
            if (pVal.FormTypeEx == "EFakturCoretax.MainForm")
            {
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

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_CLOSE && pVal.BeforeAction == false)
                {
                    try
                    {
                        if (_pb != null) { _pb.Stop(); _pb = null; }
                        _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                        
                        foreach (var key in SelectedCkBox.Keys.ToList())
                        {
                            SelectedCkBox[key] = false;
                        }
                        FindListModel.Clear();
                        FilterIsShow = false;
                        cbBranchValues = new Dictionary<string, string>();
                        cbOutletValues = new Dictionary<string, string>();
                        oldFromDt = null;
                        oldToDt = null;
                        oldFromDocEntry = 0;
                        oldToDocEntry = 0;
                        oldFromCust = null;
                        oldToCust = null;
                        oldFromBranch = null;
                        oldToBranch = null;
                        oldFromOutlet = null;
                        oldToOutlet = null;
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        if (_pb != null) { _pb.Stop(); _pb = null; }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_ACTIVATE
                && !pVal.BeforeAction && pVal.ActionSuccess)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    try
                    {
                        if (_pb != null) { _pb.Stop(); _pb = null; }
                        _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                        oForm.Freeze(true);
                        if (oForm.Items.Count > 0)
                        {
                            FormHelper.RemoveFocus(oForm);
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE)
                            {
                                oForm.Items.Item("TDocNum").Enabled = false;
                                oForm.Items.Item("TSeries").Enabled = false;
                                oForm.Items.Item("TStatus").Enabled = false;
                                if (!oForm.Items.Item("TStatus").Visible) oForm.Items.Item("TStatus").Visible = true;
                                if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                                {
                                    if (oForm.Items.Item("CbStatus").Visible) oForm.Items.Item("CbStatus").Visible = false;
                                }
                                oForm.Items.Item("CkAllDt").Visible = true;
                                oForm.Items.Item("CkAllDoc").Visible = true;
                                oForm.Items.Item("CkAllCust").Visible = true;
                                oForm.Items.Item("CkAllBr").Visible = true;
                                oForm.Items.Item("CkAllOtl").Visible = true;

                                FilterIsShow = SelectedCkBox.ContainsValue(true);
                                ShowFilterGroup(oForm);
                                SetMtGenerate(oForm);

                                oForm.Items.Item("BtCSV").Enabled = false;
                                oForm.Items.Item("BtXML").Enabled = false;
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE)
                            {
                                oForm.Items.Item("TDocNum").Enabled = true;
                                oForm.Items.Item("TSeries").Enabled = true;
                                oForm.Items.Item("TStatus").Enabled = true;
                                if (oForm.Items.Item("TStatus").Visible) oForm.Items.Item("TStatus").Visible = false;
                                if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                                {
                                    if (!oForm.Items.Item("CbStatus").Visible) oForm.Items.Item("CbStatus").Visible = true;
                                }
                                else
                                {
                                    var statusItem = oForm.Items.Item("TStatus");
                                    SAPbouiCOM.Item oNewItem = oForm.Items.Add("CbStatus", SAPbouiCOM.BoFormItemTypes.it_COMBO_BOX);
                                    oNewItem.Left = statusItem.Left;
                                    oNewItem.Top = statusItem.Top;
                                    oNewItem.Width = statusItem.Width;
                                    oNewItem.Height = statusItem.Height;

                                    SAPbouiCOM.ComboBox oCombo = (SAPbouiCOM.ComboBox)oNewItem.Specific;
                                    oCombo.DataBind.SetBound(true, "@T2_CORETAX", "Status");
                                    // For FIND mode, better use unbound combo
                                    oCombo.ValidValues.Add("", "");
                                    oCombo.ValidValues.Add("O", "Open");
                                    oCombo.ValidValues.Add("C", "Closed");
                                    oCombo.Select(0, SAPbouiCOM.BoSearchKey.psk_Index);
                                }

                                oForm.Items.Item("TFromDt").Enabled = true;
                                oForm.Items.Item("TToDt").Enabled = true;
                                oForm.Items.Item("TFromDoc").Enabled = true;
                                oForm.Items.Item("TToDoc").Enabled = true;
                                oForm.Items.Item("TFromCust").Enabled = true;
                                oForm.Items.Item("TToCust").Enabled = true;
                                oForm.Items.Item("CbFromBr").Enabled = true;
                                oForm.Items.Item("CbToBr").Enabled = true;
                                oForm.Items.Item("CbFromOtl").Enabled = true;
                                oForm.Items.Item("CbToOtl").Enabled = true;

                                oForm.Items.Item("CkAllDt").Visible = false;
                                oForm.Items.Item("CkAllDoc").Visible = false;
                                oForm.Items.Item("CkAllCust").Visible = false;
                                oForm.Items.Item("CkAllBr").Visible = false;
                                oForm.Items.Item("CkAllOtl").Visible = false;

                                oForm.Items.Item("BtCSV").Enabled = false;
                                oForm.Items.Item("BtXML").Enabled = false;
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_OK_MODE)
                            {
                                oForm.Items.Item("TDocNum").Enabled = false;
                                oForm.Items.Item("TSeries").Enabled = false;
                                oForm.Items.Item("TStatus").Enabled = false;
                                if (!oForm.Items.Item("TStatus").Visible) oForm.Items.Item("TStatus").Visible = true;
                                if (FormHelper.ItemIsExists(oForm, "CbStatus"))
                                {
                                    if (oForm.Items.Item("CbStatus").Visible) oForm.Items.Item("CbStatus").Visible = false;
                                }
                                SelectedCkBox["OINV"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkInv").Specific).Checked;
                                SelectedCkBox["ODPI"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkDp").Specific).Checked;
                                SelectedCkBox["ORIN"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkCm").Specific).Checked;

                                oForm.Items.Item("CkAllDt").Visible = true;
                                oForm.Items.Item("CkAllDoc").Visible = true;
                                oForm.Items.Item("CkAllCust").Visible = true;
                                oForm.Items.Item("CkAllBr").Visible = true;
                                oForm.Items.Item("CkAllOtl").Visible = true;

                                FilterIsShow = SelectedCkBox.ContainsValue(true);
                                ShowFilterGroup(oForm);
                                SetMtGenerate(oForm);

                                oForm.Items.Item("BtCSV").Enabled = true;
                                oForm.Items.Item("BtXML").Enabled = true;
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_UPDATE_MODE)
                            {
                                oForm.Items.Item("BtCSV").Enabled = false;
                                oForm.Items.Item("BtXML").Enabled = true;
                            }
                            else if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_EDIT_MODE)
                            {
                                oForm.Items.Item("BtCSV").Enabled = false;
                                oForm.Items.Item("BtXML").Enabled = false;
                            }
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        if (_pb != null) { _pb.Stop(); _pb = null; }
                        oForm.Freeze(false);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_DATA_ADD && !pVal.BeforeAction)
                {
                    try
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "After Add Event",
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                    }
                    catch (Exception ex)
                    {
                        Application.SBO_Application.StatusBar.SetText(
                            "Error refreshing matrix: " + ex.Message,
                            SAPbouiCOM.BoMessageTime.bmt_Short,
                            SAPbouiCOM.BoStatusBarMessageType.smt_Error
                        );
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE &&
                pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                    //SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    try
                    {
                        oForm.Freeze(true);
                        if (_pb != null) { _pb.Stop();_pb = null; }
                        _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                        string docEntry = oDBDS_Header.GetValue("DocEntry", 0).Trim();

                        // Run only if new document (Add mode / no DocEntry yet)
                        if (string.IsNullOrEmpty(docEntry) || docEntry == "0")
                        {
                            var seriesId = QueryHelper.GetSeriesIdCoretax();

                            // Update DB DataSource
                            oDBDS_Header.SetValue("Series", 0, seriesId.ToString());
                            oDBDS_Header.SetValue("Status", 0, "O");


                            // Set display fields (unbound helper fields)
                            if (FormHelper.ItemIsExists(oForm, "TSeries")) ((SAPbouiCOM.EditText)oForm.Items.Item("TSeries").Specific).Value = QueryHelper.GetSeriesName(seriesId);
                            if (FormHelper.ItemIsExists(oForm, "TDocNum")) ((SAPbouiCOM.EditText)oForm.Items.Item("TDocNum").Specific).Value = QueryHelper.GetLastDocNum(seriesId).ToString();
                            if (FormHelper.ItemIsExists(oForm, "TStatus")) ((SAPbouiCOM.EditText)oForm.Items.Item("TStatus").Specific).Value = "Open";
                            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && oForm.Items.Count > 0)
                            {
                                ShowFilterGroup(oForm);
                            }
                        }
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                    finally
                    {
                        if (_pb != null) { _pb.Stop(); _pb = null; }
                        oForm.Freeze(false);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && !pVal.BeforeAction) // after the press
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
                        ExportToXml(FormUID);
                    }

                    //
                    if (pVal.ItemUID == "BtCSV")
                    {
                        ExportToCsv(FormUID);
                    }

                    if (pVal.ItemUID == "BtClose")
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                        string strEntry = oDBDS_Header.GetValue("DocEntry", 0).Trim();
                        if (!string.IsNullOrEmpty(strEntry) && strEntry != "0")
                        {
                            int docEntry = int.Parse(strEntry);
                            // Confirmation dialog
                            int response = Application.SBO_Application.MessageBox(
                                $"Are you sure you want to close this document?",
                                1,
                                "Yes",
                                "No",
                                ""
                            );

                            if (response == 1) // Yes
                            {
                                Task.Run(async () => {

                                    try
                                    {
                                        SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                                        var docNum = int.Parse(oDBDS_Header.GetValue("DocNum", 0).Trim());
                                        var datas = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);
                                        await TransactionService.CloseCoretax(docEntry);
                                        await TransactionService.UpdateStatusInv(docNum,datas);
                                    }
                                    catch (Exception ex)
                                    {
                                        Application.SBO_Application.StatusBar.SetText($"Error closing: {ex.Message}",
                                            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                        throw;
                                    }
                                    Application.SBO_Application.StatusBar.SetText(
                                        "Successfully closed.",
                                        SAPbouiCOM.BoMessageTime.bmt_Short,
                                        SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                    );

                                }).ContinueWith(task => {
                                    oForm.Mode = SAPbouiCOM.BoFormMode.fm_OK_MODE;
                                });
                            }
                        }
                    }

                    if (pVal.ItemUID == "1")
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                        if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_ADD_MODE && !pVal.BeforeAction && pVal.ActionSuccess)
                        {
                            NewForm(oForm);
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_VALIDATE)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                    if (pVal.BeforeAction)
                    {
                        if (pVal.ItemUID == "TFromDt")
                        {
                            string strDate = oDBDS_Header.GetValue("U_T2_From_Date", 0).Trim();
                            if (DateTime.TryParse(strDate, out DateTime parsedDate)) oldFromDt = parsedDate;
                        }
                        if (pVal.ItemUID == "TToDt")
                        {
                            string strDate = oDBDS_Header.GetValue("U_T2_To_Date", 0).Trim();
                            if (DateTime.TryParse(strDate, out DateTime parsedDate)) oldToDt = parsedDate;
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
                        if (pVal.ItemUID == "TFromDt" || pVal.ItemUID == "TToDt")
                        {
                            string strFromDate = oDBDS_Header.GetValue("U_T2_From_Date", 0).Trim();
                            string strToDate = oDBDS_Header.GetValue("U_T2_To_Date", 0).Trim();
                            DateTime? newFromDate = null;
                            DateTime? newToDate = null;
                            if (DateTime.TryParse(strFromDate, out DateTime parsedDate)) newFromDate = parsedDate;
                            if (DateTime.TryParse(strToDate, out DateTime parsedToDate)) newToDate = parsedToDate;
                            if ((oldFromDt != newFromDate) || (oldToDt != newToDate))
                            {
                                ResetDetail(oForm);
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
                            if ((oldFromDocEntry != newFromEntry) || (oldToDocEntry != newToEntry))
                            {
                                ResetDetail(oForm);
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
                            if ((oldFromCust != newFromCust)|| (oldToCust != newToCust))
                            {
                                ResetDetail(oForm);
                            }

                            // Checkbox logic
                            if (string.IsNullOrEmpty(newFromCust) && string.IsNullOrEmpty(newToCust))
                                FormHelper.SetValueDS(oForm, "CkCustDS", "Y");
                            else
                                FormHelper.SetValueDS(oForm, "CkCustDS", "N");
                        }
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST && !pVal.BeforeAction)
                {
                    CflHandler(pVal, FormUID);
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && !pVal.BeforeAction)
                {
                    if (pVal.ItemUID == "MtFind" && pVal.EventType == SAPbouiCOM.BoEventTypes.et_CLICK && pVal.BeforeAction == false && pVal.ColUID == "Col_10")
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                        var status = oDBDS_Header.GetValue("Status", 0).Trim();
                        if (status != "O") return;
                        SelectFilterHandler(FormUID, pVal.Row);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_COMBO_SELECT)
                {
                    if (new[] { "CbFromBr", "CbToBr", "CbFromOtl", "CbToOtl" }.Contains(pVal.ItemUID))
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                        if (pVal.BeforeAction)
                        {
                            if (pVal.ItemUID == "CbFromBr")
                            {
                                oldFromBranch = oDBDS_Header.GetValue("U_T2_From_Branch", 0).Trim();
                            }
                            if (pVal.ItemUID == "CbTiBr")
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

                        //CbHandler(pVal, FormUID);
                    }
                }

                if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_RESIZE &&
                    pVal.BeforeAction == false)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    if (FormHelper.ItemIsExists(oForm,"MtFind"))
                    {
                        AdjustMatrix(oForm);
                    }
                }
            }
        }

        private void NewForm(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop();_pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                oDBDS_Header.Clear();
                oDBDS_Header.InsertRecord(0);
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                oDBDS_Detail.Clear();
                oDBDS_Detail.InsertRecord(0);

                foreach (var key in SelectedCkBox.Keys.ToList())
                {
                    SelectedCkBox[key] = false;
                }
                FindListModel.Clear();
                FilterIsShow = false;
                cbBranchValues = new Dictionary<string, string>();
                cbOutletValues = new Dictionary<string, string>();
                oldFromDt = null;
                oldToDt = null;
                oldFromDocEntry = 0;
                oldToDocEntry = 0;
                oldFromCust = null;
                oldToCust = null;

                ShowFilterGroup(oForm);
                ResetDetail(oForm);

                var seriesId = QueryHelper.GetSeriesIdCoretax();
                var nextDocNum = QueryHelper.GetLastDocNum(seriesId).ToString();
                oDBDS_Header.SetValue("DocNum", 0, nextDocNum);

                // unbound display fields
                if (FormHelper.ItemIsExists(oForm, "TDocNum"))
                    ((SAPbouiCOM.EditText)oForm.Items.Item("TDocNum").Specific).Value = nextDocNum;


                // Update DB DataSource
                oDBDS_Header.SetValue("Series", 0, seriesId.ToString());
                oDBDS_Header.SetValue("Status", 0, "O");
                
                // Set display fields (unbound helper fields)
                if (FormHelper.ItemIsExists(oForm, "TSeries")) ((SAPbouiCOM.EditText)oForm.Items.Item("TSeries").Specific).Value = QueryHelper.GetSeriesName(seriesId);
                if (FormHelper.ItemIsExists(oForm, "TStatus")) ((SAPbouiCOM.EditText)oForm.Items.Item("TStatus").Specific).Value = "Open";
                FormHelper.SetEnabled(oForm, new[] { "BtCSV", "BtXML" }, false);
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void AdjustMatrix(SAPbouiCOM.Form oForm)
        {
            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop();_pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
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
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        //private void CbHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        //{
        //    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
        //    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
        //    SAPbouiCOM.ComboBox oCombo = null;
        //    try
        //    {
        //        if (_pb != null) { _pb.Stop(); _pb = null; }
        //        _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
        //        oForm.Freeze(true);
        //        if (pVal.ItemUID == "CbFromBr")
        //        {
        //            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbFromBr").Specific;

        //            string selValue = oCombo.Value;                    // Key/Value yg dipilih
        //            string selDesc = oCombo.Selected.Description;     // Description yg ditampilkan

        //            if ((coretaxModel.FromBranch ?? "") != selValue)
        //            {
        //                coretaxModel.FromBranch = selValue;
        //                ResetDetail(oForm);
        //            }
        //            if ((coretaxModel.FromBranch ?? "") == "" && (coretaxModel.ToBranch ?? "") == "")
        //                FormHelper.SetValueDS(oForm, "CkBrDS", "Y");
        //            else if ((coretaxModel.FromBranch ?? "") != "")
        //                FormHelper.SetValueDS(oForm, "CkBrDS", "N");
        //        }
        //        else if (pVal.ItemUID == "CbToBr")
        //        {
        //            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbToBr").Specific;

        //            string selValue = oCombo.Value;                    // Key/Value yg dipilih
        //            string selDesc = oCombo.Selected.Description;     // Description yg ditampilkan

        //            if ((coretaxModel.ToBranch ?? "") != selValue)
        //            {
        //                coretaxModel.ToBranch = selValue;
        //                ResetDetail(oForm);
        //            }
        //            if ((coretaxModel.FromBranch ?? "") == "" && (coretaxModel.ToBranch ?? "") == "")
        //                FormHelper.SetValueDS(oForm, "CkBrDS", "Y");
        //            else if ((coretaxModel.ToBranch ?? "") != "")
        //                FormHelper.SetValueDS(oForm, "CkBrDS", "N");
        //        }
        //        else if (pVal.ItemUID == "CbFromOtl")
        //        {
        //            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbFromOtl").Specific;

        //            string selValue = oCombo.Value;                    // Key/Value yg dipilih
        //            string selDesc = oCombo.Selected.Description;     // Description yg ditampilkan

        //            if ((coretaxModel.FromOutlet ?? "") != selValue)
        //            {
        //                coretaxModel.FromOutlet = selValue;
        //                ResetDetail(oForm);
        //            }
        //            if ((coretaxModel.FromOutlet ?? "") == "" && (coretaxModel.ToOutlet ?? "") == "")
        //                FormHelper.SetValueDS(oForm, "CkOtlDS", "Y");
        //            else if ((coretaxModel.FromOutlet ?? "") != "")
        //                FormHelper.SetValueDS(oForm, "CkOtlDS", "N");
        //        }
        //        else if (pVal.ItemUID == "CbToOtl")
        //        {
        //            oCombo = (SAPbouiCOM.ComboBox)oForm.Items.Item("CbToOtl").Specific;

        //            string selValue = oCombo.Value;                    // Key/Value yg dipilih
        //            string selDesc = oCombo.Selected.Description;     // Description yg ditampilkan

        //            if ((coretaxModel.ToOutlet ?? "") != selValue)
        //            {
        //                coretaxModel.ToOutlet = selValue;
        //                ResetDetail(oForm);
        //            }
        //            if ((coretaxModel.FromOutlet ?? "") == "" && (coretaxModel.ToOutlet ?? "") == "")
        //                FormHelper.SetValueDS(oForm, "CkOtlDS", "Y");
        //            else if ((coretaxModel.ToOutlet ?? "") != "")
        //                FormHelper.SetValueDS(oForm, "CkOtlDS", "N");
        //        }
        //    }
        //    catch (Exception)
        //    {

        //        throw;
        //    }
        //    finally
        //    {
        //        if (_pb != null) { _pb.Stop(); _pb = null; }
        //        oForm.Freeze(false);
        //    }
        //}

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
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Selecting data...", 0, false);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;

                bool selectAll = oMatrix.Columns.Item("Col_10").TitleObject.Caption != "Unselect All";

                for (int i = 1; i <= oMatrix.RowCount; i++)
                {
                    ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_10").Cells.Item(i).Specific).Checked = selectAll;
                }
                oMatrix.Columns.Item("Col_10").TitleObject.Caption = selectAll ? "Unselect All" : "Select All";
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void ToggleSelectSingle(SAPbouiCOM.Form oForm, int mtRow)
        {
            try
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Selecting...", 0, false);
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtFind").Specific;

                // Get clicked row
                int row = mtRow - 1;

                // Get checkbox value (grid stores it as string "Y"/"N" or "tYES"/"tNO")
                bool isChecked = ((SAPbouiCOM.CheckBox)oMatrix.Columns.Item("Col_10").Cells.Item(mtRow).Specific).Checked;
                SAPbouiCOM.DataTable oDT;
                if (!FormHelper.DtIsExists(oForm, "DT_FILTER"))
                {
                    oDT = oForm.DataSources.DataTables.Add("DT_FILTER");
                }
                else
                {
                    oDT = oForm.DataSources.DataTables.Item("DT_FILTER");
                }
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
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void CflHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.IChooseFromListEvent oCFLEvent = (SAPbouiCOM.IChooseFromListEvent)pVal;
            if (SelectedCkBox.Where((ck) => ck.Value).Count() == 1)
            {
                var selectedCk = SelectedCkBox.Where((ck) => ck.Value).First().Key;
                if (oCFLEvent.ChooseFromListUID == "CflDocFrom" + selectedCk)
                {
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
                    SAPbouiCOM.DataTable oDataTable = oCFLEvent.SelectedObjects;
                    SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                    if (oDataTable != null && oDataTable.Rows.Count > 0)
                    {
                        // Get values from the selected row
                        string docNum = oDataTable.GetValue("DocNum", 0).ToString();
                        string docEntry = oDataTable.GetValue("DocEntry", 0).ToString();
                        string strDocEntry = oDBDS_Header.GetValue("DocEntry",0).Trim();
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
                    SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                    if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
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
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
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
                SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
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

        private void SetMtGenerate(SAPbouiCOM.Form oForm)
        {
            try
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                oForm.Freeze(true);
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtDetail").Specific;
                
                oMatrix.Clear();

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
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void DocCheckBoxHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);

            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;

            if (oForm == null) return;

            try
            {
                oForm.Freeze(true);
                _pb?.Stop();
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("...", 0, false);

                //// Map between ItemUID, SelectedCkBox key, and model property update
                //var map = new Dictionary<string, Action<bool>>
                //{
                //    { "CkInv", isChecked => { SelectedCkBox["OINV"] = isChecked; coretaxModel.IsARInvoice = isChecked; } },
                //    { "CkDp",  isChecked => { SelectedCkBox["ODPI"] = isChecked; coretaxModel.IsARDownPayment = isChecked; } },
                //    { "CkCm",  isChecked => { SelectedCkBox["ORIN"] = isChecked; coretaxModel.IsARCreditMemo = isChecked; } }
                //};

                //if (map.TryGetValue(pVal.ItemUID, out var updateAction))
                //{
                //    var checkBox = (SAPbouiCOM.CheckBox)oForm.Items.Item(pVal.ItemUID).Specific;
                //    updateAction(checkBox.Checked);
                //}

                SelectedCkBox["OINV"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkInv").Specific).Checked;
                SelectedCkBox["ODPI"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkDp").Specific).Checked;
                SelectedCkBox["ORIN"] = ((SAPbouiCOM.CheckBox)oForm.Items.Item("CkCm").Specific).Checked;

                // Set filter visibility flag
                FilterIsShow = SelectedCkBox.ContainsValue(true);

                ShowFilterGroup(oForm);

                //ResetDetail(oForm);
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void CheckAllHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
            try
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(true);
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("...", 0, false);

                if (pVal.ItemUID == "CkAllDt")
                {
                    FormHelper.ClearEdit(oForm, "TFromDt");
                    FormHelper.ClearEdit(oForm, "TToDt");
                }
                if (pVal.ItemUID == "CkAllDoc")
                {
                    FormHelper.ClearEdit(oForm, "TFromDoc");
                    FormHelper.ClearEdit(oForm, "TToDoc");
                }
                if (pVal.ItemUID == "CkAllCust")
                {
                    FormHelper.ClearEdit(oForm, "TFromCust");
                    FormHelper.ClearEdit(oForm, "TToCust");
                }
                if (pVal.ItemUID == "CkAllBr")
                {
                    FormHelper.ResetSelectCb(oForm, "CbFromBr");
                    FormHelper.ResetSelectCb(oForm, "CbToBr");
                }
                if (pVal.ItemUID == "CkAllOtl")
                {
                    FormHelper.ResetSelectCb(oForm, "CbFromOtl");
                    FormHelper.ResetSelectCb(oForm, "CbToOtl");
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                FormHelper.RemoveFocus(oForm);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void ShowFilterGroup(SAPbouiCOM.Form oForm)
        {
            try
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                oForm.Freeze(true);
                var oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");

                string[] dateItems = { "TFromDt", "TToDt", "CkAllDt" };
                string[] docItems = { "TFromDoc", "TToDoc", "CkAllDoc" };
                string[] custItems = { "TFromCust", "TToCust", "CkAllCust" };
                string[] brItems = { "CbFromBr", "CbToBr", "CkAllBr" };
                string[] otlItems = { "CbFromOtl", "CbToOtl", "CkAllOtl" };

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
                        FormHelper.ClearEdit(oForm, "TFromDoc");
                        FormHelper.ClearEdit(oForm, "TToDoc");
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
            finally
            {
                FormHelper.RemoveFocus(oForm); // avoid stuck focus
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void ResetDetail(SAPbouiCOM.Form oForm)
        {
            if (oForm == null) return;

            try
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Loading...", 0, false);
                oForm.Freeze(true);
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
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void ResetFilters(SAPbouiCOM.Form oForm)
        {
            // Dates
            FormHelper.ClearEdit(oForm, "TFromDt");
            FormHelper.ClearEdit(oForm, "TToDt");
            FormHelper.SetValueDS(oForm, "CkDtDS", "Y");

            // Docs
            FormHelper.ClearEdit(oForm, "TFromDoc");
            FormHelper.ClearEdit(oForm, "TToDoc");
            FormHelper.SetValueDS(oForm, "CkDocDS", "Y");

            // Customers
            FormHelper.ClearEdit(oForm, "TFromCust");
            FormHelper.ClearEdit(oForm, "TToCust");
            FormHelper.SetValueDS(oForm, "CkCustDS", "Y");

            // Branches
            FormHelper.ResetSelectCb(oForm, "CbFromBr");
            FormHelper.ResetSelectCb(oForm, "CbToBr");
            FormHelper.SetValueDS(oForm, "CkBrDS", "Y");

            // Outlets
            FormHelper.ResetSelectCb(oForm, "CbFromOtl");
            FormHelper.ResetSelectCb(oForm, "CbToOtl");
            FormHelper.SetValueDS(oForm, "CkOtlDS", "Y");
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

        private void ExportToXml(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Exporting data to XML...", 0, false);
                
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                var listData = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);
                if (listData != null && listData.Any())
                {
                    // build grouped TaxInvoice with GoodServiceCollection
                    var listOfTax = new List<TaxInvoice>();
                    listOfTax = FormHelper.BuildTaxInvoiceList(listData);

                    // wrap into bulk
                    if (listOfTax.Any())
                    {
                        var TIN = listData.First().TIN;
                        var invoice = new TaxInvoiceBulk
                        {
                            TIN = TIN,
                            ListOfTaxInvoice = new ListOfTaxInvoice
                            {
                                TaxInvoiceCollection = listOfTax
                            }
                        };

                        int response = Application.SBO_Application.MessageBox(
                            $"Export data will be close the document, Are you sure you want to export this document?",
                            1,
                            "Yes",
                            "No",
                            ""
                        );

                        if (response == 1) // Yes
                        {
                            if (ExportHelper.ExportXml(invoice))
                            {
                                Application.SBO_Application.StatusBar.SetText(
                                    "Successfully exported to XML.",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                );
                                //int docEntry = 0;
                                //if (int.TryParse(oDBDS_Header.GetValue("DocEntry", 0)?.Trim(), out int parsedVal)) docEntry = parsedVal;
                                //Task.Run(async () => {

                                //    try
                                //    {
                                //        var docNum = int.Parse(oDBDS_Header.GetValue("DocNum", 0).Trim());
                                //        var datas = ConvertDetail(oForm);
                                //        await TransactionService.CloseCoretax(docEntry);
                                //        await TransactionService.UpdateStatusInv(docNum, datas);
                                //        Application.SBO_Application.StatusBar.SetText(
                                //            "Successfully Exported.",
                                //            SAPbouiCOM.BoMessageTime.bmt_Medium,
                                //            SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                //        );
                                //    }
                                //    catch (Exception ex)
                                //    {
                                //        Application.SBO_Application.StatusBar.SetText($"Error closing: {ex.Message}",
                                //            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //        throw;
                                //    }

                                //});
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("No data to Export.");
                    }

                }
                else
                {
                    throw new Exception("No data to Export.");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void ExportToCsv(string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
            try
            {
                oForm.Freeze(true);
                if (_pb != null) { _pb.Stop(); _pb = null; }
                _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Exporting data to XML...", 0, false);
                SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
                SAPbouiCOM.DBDataSource oDBDS_Detail = oForm.DataSources.DBDataSources.Item("@T2_CORETAX_DT");
                var listData = FormHelper.BuildInvoiceDetailList(oDBDS_Detail);
                if (listData.Any())
                {
                    var listOfTax = new List<TaxInvoice>();
                    listOfTax = FormHelper.BuildTaxInvoiceList(listData);
                    // wrap into bulk
                    if (listOfTax.Any())
                    {
                        var TIN = listData.First().TIN;
                        var invoice = new TaxInvoiceBulk
                        {
                            TIN = TIN,
                            ListOfTaxInvoice = new ListOfTaxInvoice
                            {
                                TaxInvoiceCollection = listOfTax
                            }
                        };

                        int response = Application.SBO_Application.MessageBox(
                            $"Export data will be close the document, Are you sure you want to export this document?",
                            1,
                            "Yes",
                            "No",
                            ""
                        );

                        if (response == 1) // Yes
                        {
                            if (ExportHelper.ExportCsv(invoice))
                            {
                                Application.SBO_Application.StatusBar.SetText(
                                    "Successfully exported to CSV.",
                                    SAPbouiCOM.BoMessageTime.bmt_Short,
                                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                );
                                //int docEntry = 0;
                                //if (int.TryParse(oDBDS_Header.GetValue("DocEntry", 0)?.Trim(), out int parsedVal)) docEntry = parsedVal;
                                //Task.Run(async () => {

                                //    try
                                //    {
                                //        var docNum = int.Parse(oDBDS_Header.GetValue("DocNum", 0).Trim());
                                //        var datas = ConvertDetail(oForm);
                                //        await TransactionService.CloseCoretax(docEntry);
                                //        await TransactionService.UpdateStatusInv(docNum, datas);
                                //        Application.SBO_Application.StatusBar.SetText(
                                //            "Successfully Exported.",
                                //            SAPbouiCOM.BoMessageTime.bmt_Medium,
                                //            SAPbouiCOM.BoStatusBarMessageType.smt_Success
                                //        );
                                //    }
                                //    catch (Exception ex)
                                //    {
                                //        Application.SBO_Application.StatusBar.SetText($"Error closing: {ex.Message}",
                                //            SAPbouiCOM.BoMessageTime.bmt_Long, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                                //        throw;
                                //    }

                                //});
                            }
                        }
                    }
                    else
                    {
                        throw new Exception("No data to Export.");
                    }
                }
                else
                {
                    throw new Exception("No data to Export.");
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    $"Error: {ex.Message}",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
            finally
            {
                if (_pb != null) { _pb.Stop(); _pb = null; }
                oForm.Freeze(false);
            }
        }

        private void BtnGenHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
            Task.Run(async () =>
            {
                try
                {
                    oForm.Freeze(true);
                    if (_pb != null) { _pb.Stop(); _pb = null; }
                    _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Generating data...", 0, false);
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
                    if (_pb != null) { _pb.Stop(); _pb = null; }
                    oForm.Freeze(false);
                }
            });
        }

        private void BtnFilterHandler(SAPbouiCOM.ItemEvent pVal, string FormUID)
        {
            SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
            SAPbouiCOM.DBDataSource oDBDS_Header = oForm.DataSources.DBDataSources.Item("@T2_CORETAX");
            if (oForm.Mode == SAPbouiCOM.BoFormMode.fm_FIND_MODE) return;
            Task.Run(async () => {
                try
                {
                    oForm.Freeze(true);
                    if (_pb != null) { _pb.Stop(); _pb = null; }
                    _pb = Application.SBO_Application.StatusBar.CreateProgressBar("Retrieving data...", 0, false);
                    _pb.Text = "Retrieving data...";
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
                    //if (ItemIsExists(oForm, "CkAllCust") && oForm.Items.Item("CkAllCust").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllCust").Specific;
                        allCust = oCk.Checked;
                    }
                    if (oForm.Items.Item("CkAllBr").Enabled)
                    //if (ItemIsExists(oForm, "CkAllBr") && oForm.Items.Item("CkAllBr").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllBr").Specific;
                        allBranch = oCk.Checked;
                    }
                    if (oForm.Items.Item("CkAllOtl").Enabled)
                    //if (ItemIsExists(oForm, "CkAllOtl") && oForm.Items.Item("CkAllOtl").Enabled)
                    {
                        SAPbouiCOM.CheckBox oCk = (SAPbouiCOM.CheckBox)oForm.Items.Item("CkAllOtl").Specific;
                        allOutlet = oCk.Checked;
                    }

                    if (!allDt)
                    {
                        if (oForm.Items.Item("TFromDt").Enabled)
                        //if (ItemIsExists(oForm, "TFromDt") && oForm.Items.Item("TFromDt").Enabled)
                        {
                            string oDtFrom = oDBDS_Header.GetValue("U_T2_From_Date",0).Trim();
                            if (DateTime.TryParseExact(oDtFrom, "yyyyMMdd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDtFrom))
                            {
                                dtFrom = parsedDtFrom.ToString("yyyy-MM-dd");
                            }
                        }
                        if (oForm.Items.Item("TToDt").Enabled)
                        //if (ItemIsExists(oForm, "TToDt") && oForm.Items.Item("TToDt").Enabled)
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
                        //if (ItemIsExists(oForm, "TFromDoc") && oForm.Items.Item("TFromDoc").Enabled)
                        {
                            docFrom = oDBDS_Header.GetValue("U_T2_From_Doc_Entry", 0).Trim();
                        }
                        if (oForm.Items.Item("TToDoc").Enabled)
                        //if (ItemIsExists(oForm, "TToDoc") && oForm.Items.Item("TToDoc").Enabled)
                        {
                            docTo = oDBDS_Header.GetValue("U_T2_To_Doc_Entry", 0).Trim();
                        }
                    }
                    if (!allCust)
                    {
                        if (oForm.Items.Item("TFromCust").Enabled)
                        //if (ItemIsExists(oForm, "TFromCust") && oForm.Items.Item("TFromCust").Enabled)
                        {
                            custFrom = oDBDS_Header.GetValue("U_T2_From_Cust", 0).Trim();
                        }
                        if (oForm.Items.Item("TToCust").Enabled)
                        //if (ItemIsExists(oForm, "TToCust") && oForm.Items.Item("TToCust").Enabled)
                        {
                            custTo = oDBDS_Header.GetValue("U_T2_To_Cust", 0).Trim();
                        }
                    }
                    if (!allBranch)
                    {
                        if (oForm.Items.Item("CbFromBr").Enabled)
                        //if (ItemIsExists(oForm, "CbFromBr") && oForm.Items.Item("CbFromBr").Enabled)
                        {
                            branchFrom = oDBDS_Header.GetValue("U_T2_From_Branch",0).Trim();
                        }
                        if (oForm.Items.Item("CbToBr").Enabled)
                        //if (ItemIsExists(oForm, "CbToBr") && oForm.Items.Item("CbToBr").Enabled)
                        {
                            branchTo = oDBDS_Header.GetValue("U_T2_To_Branch", 0).Trim();
                        }
                    }
                    if (!allOutlet)
                    {
                        if (oForm.Items.Item("CbFromOtl").Enabled)
                        //if (ItemIsExists(oForm, "CbFromOtl") && oForm.Items.Item("CbFromOtl").Enabled)
                        {
                            outFrom = oDBDS_Header.GetValue("U_T2_From_Outlet", 0).Trim();
                        }
                        if (oForm.Items.Item("CbToOtl").Enabled)
                        //if (ItemIsExists(oForm, "CbToOtl") && oForm.Items.Item("CbToOtl").Enabled)
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
                    if (_pb != null) { _pb.Stop(); _pb = null; }
                    oForm.Freeze(false);
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
