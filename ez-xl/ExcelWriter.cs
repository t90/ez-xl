using System;
using System.Collections.Generic;
using System.IO.Compression;
using System.Linq;
using System.Xml;

namespace ez_xl
{
    public class ExcelWriter : IDisposable 
    {
        private readonly string _odsTemplateFileName;

        public ExcelWriter(string odsTemplateFileName)
        {
            _odsTemplateFileName = odsTemplateFileName;
        }

        public void Write(string outputOdsFile, params Tuple<string, IEnumerable<object>>[] dataSheets)
        {
            using (var result = ZipFile.Open(outputOdsFile, ZipArchiveMode.Create))
            using (var sample = ZipFile.OpenRead(_odsTemplateFileName))
            {

                if(!sample.Entries.Any(e => string.Equals(e.Name, "content.xml", StringComparison.CurrentCultureIgnoreCase)))
                {
                    throw new Exception("Invalid .ods file. If you are using .xlsx file as a template, open excel and do 'Save As' -> 'Computer' -> 'Save as type:'= 'OpenDocument Spreadsheet' -> 'Save'. Use the file with .ods extension.");
                }

                foreach (var entry in sample.Entries)
                {
                    if (string.Equals(entry.Name, "content.xml", StringComparison.CurrentCultureIgnoreCase))
                    {
                        var doc = new XmlDocument();
                        using (var stream = entry.Open())
                        {
                            doc.Load(stream);
                        }

                        var root = doc.DocumentElement;
                        XmlNamespaceManager nsmgr = new XmlNamespaceManager(doc.NameTable);
                        nsmgr.AddNamespace("table", "urn:oasis:names:tc:opendocument:xmlns:table:1.0");
                        nsmgr.AddNamespace("text", "urn:oasis:names:tc:opendocument:xmlns:text:1.0");
                        nsmgr.AddNamespace("office", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");

                        if (root == null)
                        {
                            throw new Exception("Corrupt ODS file. Could not read 'content.xml'. Could not find document root.");
                        }

                        var sheets = root.SelectNodes("//table:table", nsmgr).Cast<XmlNode>().Select(xmlNode => new 
                        {
                            Node = xmlNode,
                            Name = xmlNode.Attributes["name", "urn:oasis:names:tc:opendocument:xmlns:table:1.0"].Value
                        }).ToArray();

                        if (sheets.Length == 1)
                        {
                            WriteSingleSheet(sheets.First().Node, dataSheets.First().Item2);
                        }
                        else
                        {
                            var joinSheetsToData = (from sheet in sheets
                                join data in dataSheets on sheet.Name.ToLower() equals data.Item1.ToLower()
                                select new
                                {
                                    sheet.Node,
                                    Data = data.Item2,
                                }).ToList();
                            joinSheetsToData.ForEach(j => WriteSingleSheet(j.Node, j.Data));
                        }

                        var contentEntry = result.CreateEntry("content.xml", CompressionLevel.Optimal);
                        using (var cnt = contentEntry.Open())
                        {
                            doc.Save(cnt);
                        }
                    }
                    else
                    {
                        using (var outStream = entry.Open())
                        using (var destStream = result.CreateEntry(entry.FullName, CompressionLevel.Optimal).Open())
                        {
                            outStream.CopyTo(destStream);
                        }
                    }
                }
            }

        }

        private void WriteSingleSheet(XmlNode node, IEnumerable<object> dataSet)
        {
            var rows = node.ChildNodes.Cast<XmlNode>().Where(n => n.Name == "table:table-row")
                .Select(xmlNode =>
                {
                    var tableCellNodes = xmlNode.ChildNodes.Cast<XmlNode>().Where(cn => cn.Name == "table:table-cell");
                    var valueNodes = tableCellNodes.Select(tcn => tcn.ChildNodes.Cast<XmlNode>().FirstOrDefault(cn => cn.Name == "text:p")).Where(i => i != null).ToList();
                    return new
                    {
                        RowXmlNode = xmlNode,
                        ValueNodes = valueNodes,
                        Values = valueNodes.Select(vn => vn.InnerText)
                    };
                })
                .ToList();

            var templateRowAndLater = rows.SkipWhile(r => r.Values.All(v => !v.StartsWith("%"))).ToList();

            if (templateRowAndLater.Count < 1)
            {
                throw new Exception("No rows with template information found. You document should have a row with cells marked as ex. '%FieldName' in cell A2. %FieldName value will be replaced with the value stored in you object's property obj.FieldName");
            }

            templateRowAndLater.ForEach(tr => node.RemoveChild(tr.RowXmlNode));

            var templateRow = templateRowAndLater.First();
            var dataNodes = templateRow.ValueNodes.Where(vn => vn.InnerText.StartsWith("%"))
                .Select(vn => new
                {
                    Node = vn,
                    PropertyName = vn.InnerText.Replace("%","").Trim()
                })
                .ToList();

            var type = dataSet.GetType();
            var dataElementType = !type.IsArray ? type.GetGenericArguments()[0] : type.GetElementType();

            var dataProperties = dataElementType.GetProperties();

            var bindingSet = (from dataNode in dataNodes
                join dataProperty in dataProperties on dataNode.PropertyName equals dataProperty.Name
                select new
                {
                    dataNode.Node,
                    Property = dataProperty,
                }).ToList();

            foreach (var item in dataSet)
            {
                foreach (var bindingItem in bindingSet)
                {
                    var value = bindingItem.Property.GetValue(item);
                    var valueText = Convert.ToString(value);
                    if (value != null)
                    {
                        if (IsNumericType(bindingItem.Property.PropertyType))
                        {
                            var valueStoringNode = bindingItem.Node.ParentNode;
                            var valueAttribute = valueStoringNode.Attributes.Cast<XmlAttribute>().FirstOrDefault(a => a.Name == "office:value");
                            var typeAttribute = valueStoringNode.Attributes.Cast<XmlAttribute>().FirstOrDefault(a => a.Name == "office:value-type");
                            if (valueAttribute == null)
                            {
                                valueAttribute = valueStoringNode.OwnerDocument.CreateAttribute("office:value", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
                                valueStoringNode.Attributes.Append(valueAttribute);
                            }
                            if (typeAttribute == null)
                            {
                                typeAttribute = valueStoringNode.OwnerDocument.CreateAttribute("office:value-type", "urn:oasis:names:tc:opendocument:xmlns:office:1.0");
                                valueStoringNode.Attributes.Append(typeAttribute);
                            }
                            typeAttribute.Value = "float";
                            valueAttribute.Value = valueText;
                        }
                    }
                    bindingItem.Node.InnerText = valueText;
                }
                var clonedRow = templateRow.RowXmlNode.CloneNode(true);
                node.AppendChild(clonedRow);
            }
        }

        private static bool IsNumericType(Type type)
        {
            switch (Type.GetTypeCode(type))
            {
                case TypeCode.Byte:
                case TypeCode.SByte:
                case TypeCode.UInt16:
                case TypeCode.UInt32:
                case TypeCode.UInt64:
                case TypeCode.Int16:
                case TypeCode.Int32:
                case TypeCode.Int64:
                case TypeCode.Decimal:
                case TypeCode.Double:
                case TypeCode.Single:
                    return true;
                default:
                    return false;
            }
        }

        public void Dispose()
        {
            
        }
    }
}
