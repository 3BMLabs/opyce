using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace opyce
{
    [XmlRoot("CustomXML", Namespace = Serializer.OpyceNameSpace)]
    public class CustomXML
    {
        [XmlElement("data")]
        public string Key { get; set; }
        public string Value { get; set; }
    }
}
