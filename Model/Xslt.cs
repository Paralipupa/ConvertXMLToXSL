using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Xml.Xsl;

namespace Egrn.Model
{
    class Xslt
    {
        public static string GetXMLString(string namefile)
        {

            if (File.Exists(namefile) == false)
            {
                return string.Empty;
            }

            string textFromFile = string.Empty;
            try
            {

                using (FileStream fstream = File.OpenRead(namefile))
                {
                    byte[] array = new byte[fstream.Length];
                    fstream.Read(array, 0, array.Length);
                    textFromFile = System.Text.Encoding.UTF8.GetString(array);
                }
            }
            catch (Exception ex)
            {
                Log.Instance.Write(ex);
            }
            return textFromFile;
        }

        public static string GetXMLString(string namefile, string xslFile)
        {
            if (File.Exists(namefile) == false)
            {
                return string.Empty;
            }

            string textFromFile = string.Empty;
            try
            {
                XDocument xDoc = GetDocument(namefile);
               if (xDoc != null)
                {
                    textFromFile = TransformDocument(xDoc.ToString(), xslFile);
                }
               
            }
            catch (Exception ex)
            {
                Log.Instance.Write(ex);
            }
            return textFromFile;
        }

        public static XDocument GetDocument(string namefile)
        {
            try
            {
                XDocument xDoc = XDocument.Load(namefile, LoadOptions.SetBaseUri | LoadOptions.SetLineInfo);
                return xDoc;
            }
            catch (Exception ex)
            {
                Log.Instance.Write(ex);
            }

            return null;
        }

        public static string Transform(XDocument xdoc, string xslFile)
        {
            try
            {
                FileInfo fileInf = new FileInfo(xslFile);
                if (fileInf.Exists == false)
                {
                    return xdoc.ToString();
                }

                XslCompiledTransform transform = new XslCompiledTransform();
                XsltSettings settings = new XsltSettings();
                transform.Load(xslFile, settings, null);

                using (XmlWriter writer = XmlWriter.Create("output.xml"))
                {
                    transform.Transform(xdoc.CreateReader(), writer);
                    return writer.ToString();
                }
                 
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                Log.Instance.Write(ex);
            }

            return xdoc.ToString();
        }

        /// <summary>
        /// Xsl преобразование документа
        /// </summary>
        /// <param name="xmlString"></param>
        /// <param name="xslString"></param>
        /// <returns></returns>
        public static string TransformDocument(string xmlString, string xslString)
        {
            Func<string, XmlDocument> GetXmlDocument = (xmlContent) =>
            {
                XmlDocument xmlDocument = new XmlDocument();
                xmlDocument.LoadXml(xmlContent);
                return xmlDocument;
            };

            try
            {
                var xdoc = GetXmlDocument(xmlString);
                string styleXSL = File.ReadAllText(xslString);
                var style = GetXmlDocument(styleXSL);

                XslCompiledTransform transform = new XslCompiledTransform();
                transform.Load(style); 

                StringWriter writer = new StringWriter();
                XmlReader xmlReadB = new XmlTextReader(new StringReader(xdoc.DocumentElement.OuterXml));

                transform.Transform(xmlReadB, null, writer);

                return writer.ToString();
            }
            catch (Exception ex)
            {
                Log.Instance.Write(ex);
                Console.WriteLine(ex.Message);
            }
            return xmlString;
        }
    }
}
