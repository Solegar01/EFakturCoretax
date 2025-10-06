using System;
using System.Collections.Generic;
using System.Xml.Serialization;

[XmlRoot("TaxInvoiceBulk")]
public class TaxInvoiceBulk
{
    public string TIN { get; set; }

    [XmlElement("ListOfTaxInvoice")]
    public ListOfTaxInvoice ListOfTaxInvoice { get; set; } = new ListOfTaxInvoice();
}

public class ListOfTaxInvoice
{
    [XmlElement("TaxInvoice")]
    public List<TaxInvoice> TaxInvoiceCollection { get; set; } = new List<TaxInvoice>();
}

public class TaxInvoice
{
    public string TaxInvoiceDate { get; set; }
    public string TaxInvoiceOpt { get; set; }
    public string TrxCode { get; set; }
    public string AddInfo { get; set; }
    public string CustomDoc { get; set; }
    public string CustomDocMonthYear { get; set; }
    public string RefDesc { get; set; }
    public string FacilityStamp { get; set; }
    public string SellerIDTKU { get; set; }
    public string BuyerTin { get; set; }
    public string BuyerDocument { get; set; }
    public string BuyerCountry { get; set; }
    public string BuyerDocumentNumber { get; set; }
    public string BuyerName { get; set; }
    public string BuyerAdress { get; set; }
    public string BuyerEmail { get; set; }
    public string BuyerIDTKU { get; set; }

    [XmlElement("ListOfGoodService")]
    public ListOfGoodService ListOfGoodService { get; set; } = new ListOfGoodService();
}

public class ListOfGoodService
{
    [XmlElement("GoodService")]
    public List<GoodService> GoodServiceCollection { get; set; } = new List<GoodService>();
}

public class GoodService
{
    public string Opt { get; set; }
    public string Code { get; set; }
    public string Name { get; set; }
    public string Unit { get; set; }
    public decimal Price { get; set; }
    public decimal Qty { get; set; }
    public decimal TotalDiscount { get; set; }
    public decimal TaxBase { get; set; }
    public decimal OtherTaxBase { get; set; }
    public decimal VATRate { get; set; }
    public decimal VAT { get; set; }
    public decimal STLGRate { get; set; }
    public decimal STLG { get; set; }
}
