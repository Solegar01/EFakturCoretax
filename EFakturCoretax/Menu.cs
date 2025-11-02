using EFakturCoretax.FormHandlers;
using EFakturCoretax.Helpers;
using EFakturCoretax.Models;
using EFakturCoretax.Services;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace EFakturCoretax
{
    class Menu
    {
        
        string IconFolder = "";
        GenerateFormHandler generateFormHandler;
        ImportCoretaxFormHandler importCoretaxFormHandler;

        public Menu()
        {
            IconFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Icons");
            generateFormHandler = new GenerateFormHandler();
            Application.SBO_Application.MenuEvent += generateFormHandler.SBO_Application_MenuEvent;
            Application.SBO_Application.ItemEvent += generateFormHandler.SBO_Application_ItemEvent;
            Application.SBO_Application.FormDataEvent += generateFormHandler.SBO_Application_FormDataEvent;
            importCoretaxFormHandler = new ImportCoretaxFormHandler();
            Application.SBO_Application.MenuEvent += importCoretaxFormHandler.SBO_Application_MenuEvent;
            Application.SBO_Application.ItemEvent += importCoretaxFormHandler.SBO_Application_ItemEvent;
            Application.SBO_Application.FormDataEvent += importCoretaxFormHandler.SBO_Application_FormDataEvent;
        }

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
            string mainIcon = Path.Combine(IconFolder, "coretax_logo.bmp");
            if (File.Exists(mainIcon))
                oCreationPackage.Image = mainIcon;

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
                Application.SBO_Application.FormDataEvent += SBO_Application_FormDataEvent;
                Application.SBO_Application.ItemEvent += SBO_Application_ItemEvent;
                oMenus = oMenuItem.SubMenus;

                // Create s sub menu
                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EFakturCoretax.GenerateForm";
                oCreationPackage.String = "Generate Coretax";
                oMenus.AddEx(oCreationPackage);

                oCreationPackage.Type = SAPbouiCOM.BoMenuType.mt_STRING;
                oCreationPackage.UniqueID = "EFakturCoretax.ImportCoretaxForm";
                oCreationPackage.String = "Import Coretax";
                oMenus.AddEx(oCreationPackage);
            }
            catch (Exception)
            { //  Menu already exists
                Application.SBO_Application.SetStatusBarMessage("Menu Already Exists", SAPbouiCOM.BoMessageTime.bmt_Short, true);
            }
        }

        public void SBO_Application_MenuEvent(ref SAPbouiCOM.MenuEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;

            try
            {
                if (pVal.BeforeAction && pVal.MenuUID == "EFakturCoretax.GenerateForm")
                {
                    GenerateForm activeForm = new GenerateForm();
                    activeForm.Show();
                    
                }

                if (pVal.BeforeAction && pVal.MenuUID == "EFakturCoretax.ImportCoretaxForm")
                {

                    ImportCoretaxForm activeForm = new ImportCoretaxForm();
                    activeForm.Show();
                    
                }

            }
            catch (Exception ex)
            {
                Application.SBO_Application.MessageBox(ex.ToString(), 1, "Ok", "", "");
            }
        }

        public void SBO_Application_FormDataEvent(ref SAPbouiCOM.BusinessObjectInfo BusinessObjectInfo, out bool BubbleEvent)
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
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        // UDF form not available (not open yet)
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

        public void SBO_Application_ItemEvent(string FormUID, ref SAPbouiCOM.ItemEvent pVal, out bool BubbleEvent)
        {
            BubbleEvent = true;
            var udfFormIds = new[] { "-133", "-65300", "-179" };
            try
            {
                if (udfFormIds.Contains(pVal.FormTypeEx))
                {
                    if (pVal.EventType == SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                    && !pVal.BeforeAction) // After form loaded
                    {
                        SAPbouiCOM.Form oForm = Application.SBO_Application.Forms.Item(FormUID);
                        oForm.Items.Item("U_T2_Exported").Enabled = false;
                        oForm.Items.Item("U_T2_Coretax_No").Enabled = false;
                    }
                }
                
            }
            catch (Exception ex)
            {
                Application.SBO_Application.StatusBar.SetText(
                    ex.Message,
                    SAPbouiCOM.BoMessageTime.bmt_Short,
                    SAPbouiCOM.BoStatusBarMessageType.smt_Error
                );
            }
        }

    }
}
