using System;
using System.IO;
using System.Security;
using System.Xml;
using System.Xml.Linq;
using Microsoft.Office.Core;
using Newtonsoft.Json;
using SubmissionCollector.Enums;
using SubmissionCollector.ExcelEventSetters;
using SubmissionCollector.ExcelUtilities.Extensions;
using SubmissionCollector.Models.Package;
using SubmissionCollector.View.Forms;

namespace SubmissionCollector.ExcelWorkspaceFolder
{
    internal class CustomXmlPartManager
    {
        private readonly IWorkbookLogger _logger;
        private const string BexNamespace = "BEX";

        public CustomXmlPartManager(IWorkbookLogger logger)
        {
            _logger = logger;
        }

        internal void WriteJsonToCustomXmlPart(Package package, ThisWorkbook thisWorkbook)
        {
            try
            {
                SaveWorksheetNames(package);
                var xml = MapToXmlFormat(package);

                var xmlParts = thisWorkbook.CustomXMLParts.SelectByNamespace("BEX");
                var enumerator = xmlParts.GetEnumerator();
                enumerator.MoveNext();
                var bexXmlPart = (CustomXMLPart)enumerator.Current;
                bexXmlPart?.Delete();

                thisWorkbook.CustomXMLParts.Add(xml);
            }
            catch (Exception ex)
            {
                _logger.WriteNew(ex);
                const string message = "Custom XML part failed: Contact ARM";
                MessageHelper.Show(message, MessageType.Stop);
            }
        }

        internal Package RestoreFromCustomXmlPart(CustomXMLPart bexCustomXmlPart)
        {
            var json = string.Empty;
            var stringReader = new StringReader(bexCustomXmlPart.XML);

            using (var reader = XmlReader.Create(stringReader))
            {
                reader.MoveToContent();
                while (reader.Read())
                {
                    if (reader.NodeType != XmlNodeType.Element || reader.Name != "json") continue;
                    json = ((XElement) XNode.ReadFrom(reader)).Value;
                }
            }

            Package package;
            using (new ExcelEventDisabler())
            {
                package = JsonConvert.DeserializeObject<Package>(
                    json,
                    new JsonSerializerSettings
                    {
                        TypeNameHandling = TypeNameHandling.Auto,
                        PreserveReferencesHandling = PreserveReferencesHandling.Objects
                    });
            }

            if (package == null) throw new ArgumentException("Stored workbook information not retrievable");
            return package;
        }

        internal void RestoreWorksheetsFromJson(Package package)
        {
            package.Worksheet = package.WorksheetName.GetWorksheetWithNullTrap();
            foreach (var segment in package.Segments)
            {
                var worksheetName = segment.WorksheetManager.WorksheetName;
                segment.WorksheetManager.Worksheet = worksheetName.GetWorksheetWithNullTrap();
            }
        }

        internal static void GetBexCustomXmlPartWithMessage(IWorkbookLogger logger)
        {
            try
            {
                CustomXMLPart bexXmlPart;
                using (new CursorToWait())
                {
                    bexXmlPart = GetBexCustomXmlPart();
                }

                MessageHelper.Show(bexXmlPart != null ? bexXmlPart.XML : "No custom xml part");
            }
            catch (Exception ex)
            {
                logger.WriteNew(ex);
                MessageHelper.Show("Get custom xml part failed", MessageType.Stop);
            }
        }

        internal static CustomXMLPart GetBexCustomXmlPart()
        {
            // ReSharper disable once RedundantCast
            var custom = (CustomXMLParts) Globals.ThisWorkbook.CustomXMLParts;
            var xmlParts = custom.SelectByNamespace(BexNamespace);
            var enumerator = xmlParts.GetEnumerator();
            enumerator.MoveNext();
            return (CustomXMLPart)enumerator.Current;
        }

        private static string MapToXmlFormat(Package package)
        {
            var json = SerializeManager.ConvertToJson(package);

            var encodedJson = SecurityElement.Escape(json);
            var xml =
                "<?xml version=\"1.0\" encoding=\"utf-8\" ?>" +
                $"<package xmlns=\"{BexNamespace}\">" +
                "<json>" +
                encodedJson +
                "</json>" +
                "</package>";
            return xml;
        }

        private static void SaveWorksheetNames(Package package)
        {
            package.WorksheetName = package.Worksheet.Name;
            foreach (var segment in package.Segments)
            {
                segment.WorksheetManager.WorksheetName = segment.WorksheetManager.Worksheet.Name;
            }
        }
    }
}
