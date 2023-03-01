using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace ParsingXml
{

    [System.SerializableAttribute()]
    [System.ComponentModel.DesignerCategoryAttribute("code")]
    [System.Xml.Serialization.XmlTypeAttribute(AnonymousType = true)]
    [System.Xml.Serialization.XmlRootAttribute(Namespace = "", IsNullable = false)]
    public class DtStatement
    {
        [XmlElement(ElementName = "Statement")]
        public List<Statement> Statement { get; set; }
    }
}
