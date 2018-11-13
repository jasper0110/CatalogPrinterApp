using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace XMLUtil
{
    public static class XMLUtility
    {
        public static void WriteToXml(string xmlPath, params KeyValuePair<string, string>[] kvps)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(xmlPath);
            XmlNode root = doc.DocumentElement;

            XmlNode appSettings = root.SelectSingleNode("appSettings") ?? root.PrependChild(doc.CreateElement("appSettings"));

            

            foreach (var kvp in kvps)
            {
                XmlElement add = appSettings.ChildNodes.OfType<XmlElement>().Where(n => n.Attributes["key"].Value == kvp.Key).FirstOrDefault() ?? doc.CreateElement("add");
                add.SetAttribute("key", kvp.Key);
                add.SetAttribute("value", kvp.Value);
                appSettings.AppendChild(add);
            }


            
            root.AppendChild(appSettings);

            doc.Save(xmlPath);
        }
    }
}
