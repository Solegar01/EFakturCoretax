using SAPbouiCOM.Framework;
using System;
using System.Collections.Generic;
using System.Xml;

namespace EFakturCoretax
{
    [FormAttribute("EFakturCoretax.MainForm", "MainForm.b1f")]
    class MainForm : UserFormBase
    {
        public MainForm()
        {
        }

        /// <summary>
        /// Initialize components. Called by framework after form created.
        /// </summary>
        public override void OnInitializeComponent()
        {
            this.Button0 = ((SAPbouiCOM.Button)(this.GetItem("1").Specific));
            this.Button1 = ((SAPbouiCOM.Button)(this.GetItem("2").Specific));
            this.StaticText0 = ((SAPbouiCOM.StaticText)(this.GetItem("LDisp").Specific));
            this.CheckBox0 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkInv").Specific));
            this.CheckBox1 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkDp").Specific));
            this.CheckBox2 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkCm").Specific));
            this.StaticText1 = ((SAPbouiCOM.StaticText)(this.GetItem("LRange").Specific));
            this.StaticText2 = ((SAPbouiCOM.StaticText)(this.GetItem("LFromDt").Specific));
            this.StaticText3 = ((SAPbouiCOM.StaticText)(this.GetItem("LFromDoc").Specific));
            this.StaticText4 = ((SAPbouiCOM.StaticText)(this.GetItem("LFromCust").Specific));
            this.StaticText5 = ((SAPbouiCOM.StaticText)(this.GetItem("LFromBr").Specific));
            this.StaticText6 = ((SAPbouiCOM.StaticText)(this.GetItem("LFromOtl").Specific));
            this.EditText0 = ((SAPbouiCOM.EditText)(this.GetItem("TFromDt").Specific));
            this.StaticText7 = ((SAPbouiCOM.StaticText)(this.GetItem("LToDt").Specific));
            this.StaticText8 = ((SAPbouiCOM.StaticText)(this.GetItem("LDate").Specific));
            this.StaticText9 = ((SAPbouiCOM.StaticText)(this.GetItem("LDoc").Specific));
            this.StaticText10 = ((SAPbouiCOM.StaticText)(this.GetItem("LCust").Specific));
            this.StaticText11 = ((SAPbouiCOM.StaticText)(this.GetItem("LBranch").Specific));
            this.StaticText12 = ((SAPbouiCOM.StaticText)(this.GetItem("LOutlet").Specific));
            this.EditText1 = ((SAPbouiCOM.EditText)(this.GetItem("TFromDoc").Specific));
            this.StaticText13 = ((SAPbouiCOM.StaticText)(this.GetItem("LToDoc").Specific));
            this.EditText2 = ((SAPbouiCOM.EditText)(this.GetItem("TFromCust").Specific));
            this.StaticText14 = ((SAPbouiCOM.StaticText)(this.GetItem("LToCust").Specific));
            this.StaticText15 = ((SAPbouiCOM.StaticText)(this.GetItem("LToBr").Specific));
            this.StaticText16 = ((SAPbouiCOM.StaticText)(this.GetItem("LToOtl").Specific));
            this.EditText5 = ((SAPbouiCOM.EditText)(this.GetItem("TToDt").Specific));
            this.EditText6 = ((SAPbouiCOM.EditText)(this.GetItem("TToDoc").Specific));
            this.EditText7 = ((SAPbouiCOM.EditText)(this.GetItem("TToCust").Specific));
            this.StaticText17 = ((SAPbouiCOM.StaticText)(this.GetItem("LSeries").Specific));
            this.EditText11 = ((SAPbouiCOM.EditText)(this.GetItem("TDocNum").Specific));
            this.StaticText18 = ((SAPbouiCOM.StaticText)(this.GetItem("LStatus").Specific));
            this.StaticText19 = ((SAPbouiCOM.StaticText)(this.GetItem("LPostDt").Specific));
            this.EditText13 = ((SAPbouiCOM.EditText)(this.GetItem("TPostDate").Specific));
            this.CheckBox3 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllDt").Specific));
            this.CheckBox4 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllDoc").Specific));
            this.CheckBox5 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllCust").Specific));
            this.CheckBox6 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllBr").Specific));
            this.CheckBox7 = ((SAPbouiCOM.CheckBox)(this.GetItem("CkAllOtl").Specific));
            this.Button2 = ((SAPbouiCOM.Button)(this.GetItem("BtFilter").Specific));
            this.Matrix1 = ((SAPbouiCOM.Matrix)(this.GetItem("MtFind").Specific));
            this.Button3 = ((SAPbouiCOM.Button)(this.GetItem("BtGen").Specific));
            this.Button5 = ((SAPbouiCOM.Button)(this.GetItem("BtCSV").Specific));
            this.Button6 = ((SAPbouiCOM.Button)(this.GetItem("BtXML").Specific));
            this.Matrix3 = ((SAPbouiCOM.Matrix)(this.GetItem("MtDetail").Specific));
            this.EditText15 = ((SAPbouiCOM.EditText)(this.GetItem("TStatus").Specific));
            this.ComboBox0 = ((SAPbouiCOM.ComboBox)(this.GetItem("CbFromBr").Specific));
            this.ComboBox1 = ((SAPbouiCOM.ComboBox)(this.GetItem("CbFromOtl").Specific));
            this.ComboBox2 = ((SAPbouiCOM.ComboBox)(this.GetItem("CbToBr").Specific));
            this.ComboBox3 = ((SAPbouiCOM.ComboBox)(this.GetItem("CbToOtl").Specific));
            this.ComboBox4 = ((SAPbouiCOM.ComboBox)(this.GetItem("CbSeries").Specific));
            this.EditText3 = ((SAPbouiCOM.EditText)(this.GetItem("DocEntry").Specific));
            this.OnCustomInitialize();

        }

        /// <summary>
        /// Initialize form event. Called by framework before form creation.
        /// </summary>
        public override void OnInitializeFormEvents()
        {
        }

        private SAPbouiCOM.Button Button0;

        private void OnCustomInitialize()
        {

        }

        private SAPbouiCOM.Button Button1;
        private SAPbouiCOM.StaticText StaticText0;
        private SAPbouiCOM.CheckBox CheckBox0;
        private SAPbouiCOM.CheckBox CheckBox1;
        private SAPbouiCOM.CheckBox CheckBox2;
        private SAPbouiCOM.StaticText StaticText1;
        private SAPbouiCOM.StaticText StaticText2;
        private SAPbouiCOM.StaticText StaticText3;
        private SAPbouiCOM.StaticText StaticText4;
        private SAPbouiCOM.StaticText StaticText5;
        private SAPbouiCOM.StaticText StaticText6;
        private SAPbouiCOM.EditText EditText0;
        private SAPbouiCOM.StaticText StaticText7;
        private SAPbouiCOM.StaticText StaticText8;
        private SAPbouiCOM.StaticText StaticText9;
        private SAPbouiCOM.StaticText StaticText10;
        private SAPbouiCOM.StaticText StaticText11;
        private SAPbouiCOM.StaticText StaticText12;
        private SAPbouiCOM.EditText EditText1;
        private SAPbouiCOM.StaticText StaticText13;
        private SAPbouiCOM.EditText EditText2;
        private SAPbouiCOM.StaticText StaticText14;
        private SAPbouiCOM.StaticText StaticText15;
        private SAPbouiCOM.StaticText StaticText16;
        private SAPbouiCOM.EditText EditText5;
        private SAPbouiCOM.EditText EditText6;
        private SAPbouiCOM.EditText EditText7;
        private SAPbouiCOM.StaticText StaticText17;
        private SAPbouiCOM.EditText EditText11;
        private SAPbouiCOM.StaticText StaticText18;
        private SAPbouiCOM.StaticText StaticText19;
        private SAPbouiCOM.EditText EditText13;
        private SAPbouiCOM.CheckBox CheckBox3;
        private SAPbouiCOM.CheckBox CheckBox4;
        private SAPbouiCOM.CheckBox CheckBox5;
        private SAPbouiCOM.CheckBox CheckBox6;
        private SAPbouiCOM.CheckBox CheckBox7;
        private SAPbouiCOM.Button Button2;
        private SAPbouiCOM.Matrix Matrix1;
        private SAPbouiCOM.Button Button3;
        private SAPbouiCOM.Button Button5;
        private SAPbouiCOM.Button Button6;
        private SAPbouiCOM.Matrix Matrix3;
        private SAPbouiCOM.EditText EditText15;
        private SAPbouiCOM.ComboBox ComboBox0;
        private SAPbouiCOM.ComboBox ComboBox1;
        private SAPbouiCOM.ComboBox ComboBox2;
        private SAPbouiCOM.ComboBox ComboBox3;
        private SAPbouiCOM.ComboBox ComboBox4;
        private SAPbouiCOM.EditText EditText3;
    }
}