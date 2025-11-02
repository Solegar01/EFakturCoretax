using EFakturCoretax.FormHandlers;
using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace EFakturCoretax
{
    [FormAttribute("EFakturCoretax.ImportCoretaxForm", "ImportCoretaxForm.b1f")]
    class ImportCoretaxForm : UserFormBase
    {
        public ImportCoretaxForm()
        {
            
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("LPath").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("TPath").Specific));
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("BtBrowse").Specific));
            this.Matrix0 = ((SAPbouiCOM.Matrix)(this.GetItem("MtData").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("BtGen").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("BtCancel").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.StaticText StaticText0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.Button Button0;
        private SAPbouiCOM.Matrix Matrix0;
        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.Button Button2;
    }
}
