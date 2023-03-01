using System.Xml.Serialization;

namespace ParsingXml
{
    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    public class Statement
    {
		[XmlElement(ElementName = "SERIALNO")]
		public string SERIALNO { get; set; }
		[XmlElement(ElementName = "CARDSERIALNO")]
		public string CARDSERIALNO { get; set; }
		[XmlElement(ElementName = "STMPARTCODE")]
		public string STMPARTCODE { get; set; }
		[XmlElement(ElementName = "STMPARTDESC")]
		public string STMPARTDESC { get; set; }
		[XmlElement(ElementName = "MESSAGEOFTHEMONTH")]
		public string MESSAGEOFTHEMONTH { get; set; }
		[XmlElement(ElementName = "STMSTARTDATE")]
		public string STMSTARTDATE { get; set; }
		[XmlElement(ElementName = "STMENDDATE")]
		public string STMENDDATE { get; set; }
		[XmlElement(ElementName = "CONUMBER")]
		public string CONUMBER { get; set; }
		[XmlElement(ElementName = "CREDITCARDNO")]
		public string CREDITCARDNO { get; set; }
		[XmlElement(ElementName = "CUSTOMERNAME")]
		public string CUSTOMERNAME { get; set; }
		[XmlElement(ElementName = "ADDRESS1")]
		public string ADDRESS1 { get; set; }
		[XmlElement(ElementName = "ADDRESS2")]
		public string ADDRESS2 { get; set; }
		[XmlElement(ElementName = "ADDRESS3")]
		public string ADDRESS3 { get; set; }
		[XmlElement(ElementName = "PRODUCTNAME")]
		public string PRODUCTNAME { get; set; }
		[XmlElement(ElementName = "STATEMENTDATE")]
		public string STATEMENTDATE { get; set; }
		[XmlElement(ElementName = "APR")]
		public string APR { get; set; }
		[XmlElement(ElementName = "PAYMENTDUEDATE")]
		public string PAYMENTDUEDATE { get; set; }
		[XmlElement(ElementName = "OPENINGBALANCE")]
		public string OPENINGBALANCE { get; set; }
		[XmlElement(ElementName = "TRANSACTIONDATE")]
		public string TRANSACTIONDATE { get; set; }
		[XmlElement(ElementName = "POSTINGDATE")]
		public string POSTINGDATE { get; set; }
		[XmlElement(ElementName = "DESCRIPTIONREFERENCE")]
		public string DESCRIPTIONREFERENCE { get; set; }
		[XmlElement(ElementName = "TRANSACTIONAMOUNTUSD")]
		public string TRANSACTIONAMOUNTUSD { get; set; }
		[XmlElement(ElementName = "TRANSACTIONAMOUNTPKR")]
		public string TRANSACTIONAMOUNTPKR { get; set; }
		[XmlElement(ElementName = "RUNNINGBALANCEPKR")]
		public string RUNNINGBALANCEPKR { get; set; }
		[XmlElement(ElementName = "CLOSINGBALANCEDATE")]
		public string CLOSINGBALANCEDATE { get; set; }
		[XmlElement(ElementName = "PROFIT_RECOVERED")]
		public string PROFIT_RECOVERED { get; set; }
		[XmlElement(ElementName = "MWM_PAYMENT")]
		public string MWM_PAYMENT { get; set; }
		[XmlElement(ElementName = "ISL_CONV_FLAG")]
		public string ISL_CONV_FLAG { get; set; }
	}
}