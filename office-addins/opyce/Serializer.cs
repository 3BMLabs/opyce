using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Serialization;
using Office = Microsoft.Office.Core;

namespace opyce
{
    public static class Serializer
    {
        public const string OpyceNameSpace = "http://schemas.3bm.com/customxml";
        private static string SerializeToXml<T>(T data, string namespaceUri)
        {
            var xmlSerializer = new XmlSerializer(typeof(T), namespaceUri);
            using (var stringWriter = new StringWriter())
            {
                xmlSerializer.Serialize(stringWriter, data);
                return stringWriter.ToString();
            }
        }

        private static T DeserializeFromXml<T>(string xmlContent)
        {
            var xmlSerializer = new XmlSerializer(typeof(T));
            using (var stringReader = new StringReader(xmlContent))
            {
                return (T)xmlSerializer.Deserialize(stringReader);
            }
        }
        public static void AddCustomXmlPart<T>(dynamic documentOrWorkbook, T data, string namespaceUri)
        {
            string xmlContent = SerializeToXml(data, namespaceUri);
            documentOrWorkbook.CustomXMLParts.Add(xmlContent);
        }

        public static T[] GetCustomXmlParts<T>(dynamic documentOrWorkbook, string namespaceUri)
        {
            List<T> arr = new List<T>();
            foreach (Office.CustomXMLPart xmlPart in documentOrWorkbook.CustomXMLParts)
            {
                if (xmlPart.NamespaceURI == namespaceUri)
                {
                    arr.Add(DeserializeFromXml<T>(xmlPart.XML));
                }
            }
            return arr.ToArray();
        }
    }
}
