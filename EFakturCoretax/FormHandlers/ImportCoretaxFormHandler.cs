using EFakturCoretax.Helpers;
using EFakturCoretax.Services;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax.FormHandlers
{
    public class ImportCoretaxFormHandler
    {

        #region Event
        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
        }

        public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            try
            {
                if (pVal.FormTypeEx == "EFakturCoretax.ImportCoretaxForm")
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                        if (oForm.Items.Count > 0)
                        {
                            if (FormHelper.ItemIsExists(oForm,"MtData"))
                            {
                                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtData").Specific;
                                oMatrix.AutoResizeColumns();
                            }
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BtBrowse" && !pVal.BeforeAction && pVal.ActionSuccess)
                    {
                        string result = ImportHelper.GetPathFile();
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                        SAPbouiCOM.EditText pathTxt = (SAPbouiCOM.EditText)oForm.Items.Item("TPath").Specific;
                        pathTxt.Value = result;
                        if (!string.IsNullOrEmpty(result))
                        {
                            LoadExcelToMatrix(oForm, result);
                        }
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BtCancel" && !pVal.BeforeAction)
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                        oForm.Close();
                    }

                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED && pVal.ItemUID == "BtGen" && !pVal.BeforeAction)
                    {
                        if (true)
                        {

                        }
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.ActiveForm;
                        SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Item("DT_INVOICE");

                        if (oDT.Rows.Count <= 0) { BubbleEvent = false; return; }

                        int response = Application.SBO_Application.MessageBox(
                                            $"Are you sure you want to generate Faktur Pajak?",
                                            1, "Yes", "No", "");

                        if (response != 1) { BubbleEvent = false; return; }
                        if(GenerateFakturPajak(oForm))
                            Application.SBO_Application.StatusBar.SetText("Successfully generated.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Success);
                        else
                            Application.SBO_Application.StatusBar.SetText("Failed to generate Faktur Pajak.", SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
                    }
                }
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(ex.Message, SAPbouiCOM.BoMessageTime.bmt_Medium, SAPbouiCOM.BoStatusBarMessageType.smt_Error);
            }
        }

        #endregion

        #region Function

        private void LoadExcelToMatrix(SAPbouiCOM.Form oForm, string filePath)
        {
            try
            {
                FormHelper.StartLoading(oForm, "Loading...", 0, false);
                // Step 1: Import Excel
                var invoices = ImportHelper.ImportInvoices(filePath);

                // Step 2: Prepare Matrix and DataTable
                SAPbouiCOM.Matrix oMatrix = (SAPbouiCOM.Matrix)oForm.Items.Item("MtData").Specific;
                SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Item("DT_INVOICE");
                oDT.Rows.Clear();

                // Step 3: Fill DataTable
                int rowIndex = 0;
                foreach (var item in invoices)
                {
                    oDT.Rows.Add();
                    oDT.SetValue("Col_0", rowIndex, item.InvoiceNo);
                    oDT.SetValue("Col_1", rowIndex, item.CoretaxNo);
                    oDT.SetValue("#", rowIndex, rowIndex+1);
                    rowIndex++;
                }

                // Step 4: Refresh Matrix
                oMatrix.LoadFromDataSource();
                oMatrix.AutoResizeColumns();

                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                    $"Loaded {invoices.Count} records from Excel",
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Success
                );
            }
            catch (Exception ex)
            {
                SAPbouiCOM.Framework.Application.SBO_Application.StatusBar.SetText(
                    "Error: " + ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }


        #endregion

        public bool GenerateFakturPajak(SAPbouiCOM.Form oForm)
        {
            SAPbobsCOM.Company oCompany = null;
            try
            {
                FormHelper.StartLoading(oForm, "Generating Faktur Pajak...", 0, false);
                oCompany = CompanyService.GetCompany();
                if (!oCompany.InTransaction)
                {
                    oCompany.StartTransaction();
                }
                SAPbouiCOM.DataTable oDT = oForm.DataSources.DataTables.Item("DT_INVOICE");
                if (oDT.Rows.Count > 0)
                {
                    for (int i = 0; i < oDT.Rows.Count; i++)
                    {
                        string docNum = "";
                        string objCode = "";
                        string tempDocNum = oDT.GetValue("Col_0", i).ToString();
                        string coretaxNo = oDT.GetValue("Col_1", i).ToString();
                        var tempArr = tempDocNum.Split('-');
                        objCode = tempArr[0];
                        docNum = tempArr[1];
                        switch (objCode)
                        {
                            case "IN":
                                TransactionService.UpdateFakturArInvoice(oCompany, docNum, coretaxNo);
                                break;
                            case "DP":
                                TransactionService.UpdateFakturArDownPayment(oCompany, docNum, coretaxNo);
                                break;
                            case "CN":
                                TransactionService.UpdateFakturArCreditMemo(oCompany, docNum, coretaxNo);
                                break;
                            default:
                                break;
                        }
                    }
                }
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_Commit);
                }
                return true;
            }
            catch (Exception)
            {
                if (oCompany.InTransaction)
                {
                    oCompany.EndTransaction(SAPbobsCOM.BoWfTransOpt.wf_RollBack);
                }
                throw;
            }
            finally
            {
                FormHelper.FinishLoading(oForm);
            }
        }
    }
}
