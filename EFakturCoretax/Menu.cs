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

        decimal oldVatRate = 0;
        private List<FilterDataModel> SelectedReviseDoc = new List<FilterDataModel>();
        string IconFolder = "";

        public Menu()
        {
            IconFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Icons");
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
                var generateFormHandler = new GenerateFormHandler();
                oMenuItem = Application.SBO_Application.Menus.Item("EFakturCoretax");
                Application.SBO_Application.MenuEvent += generateFormHandler.SBO_Application_MenuEvent;
                Application.SBO_Application.ItemEvent += generateFormHandler.SBO_Application_ItemEvent;
                Application.SBO_Application.FormDataEvent += generateFormHandler.SBO_Application_FormDataEvent;
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

    }
}
