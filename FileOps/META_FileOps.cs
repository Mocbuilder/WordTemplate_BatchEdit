using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace WordTemplate_BatchEdit.FileOps
{
    public class META_FileOps
    {
        public static async Task META_SingleFile_GetMetaData(FileInfo file, string dump, string output)
        {
            Dictionary<string, string> metadata = new Dictionary<string, string>();

            if (!File.Exists(output))
            {

            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file.FullName, false))
            {
                var coreProperties = wordDoc.PackageProperties;

                Console.WriteLine("Core Properties:");
                Console.WriteLine($"Title: {coreProperties.Title}");
                Console.WriteLine($"Author: {coreProperties.Creator}");
                Console.WriteLine($"Subject: {coreProperties.Subject}");
                Console.WriteLine($"Description: {coreProperties.Description}");
                Console.WriteLine($"Keywords: {coreProperties.Keywords}");
                Console.WriteLine($"Last Modified By: {coreProperties.LastModifiedBy}");
                Console.WriteLine($"Created Date: {coreProperties.Created}");
                Console.WriteLine($"Modified Date: {coreProperties.Modified}");

                metadata.Add("Title", coreProperties.Title!);
                metadata.Add("Author", coreProperties.Creator!);
                metadata.Add("Subject", coreProperties.Subject!);
                metadata.Add("Description", coreProperties.Description!);
                metadata.Add("Keywords", coreProperties.Keywords!);
                metadata.Add("LastModBy", coreProperties.LastModifiedBy!);
                metadata.Add("CreatedDate", coreProperties.Created?.ToString("o")!);
                metadata.Add("ModifiedDate", coreProperties.Modified?.ToString("o")!);

                var customProperties = wordDoc.ExtendedFilePropertiesPart?.Properties;

                if (customProperties != null)
                {
                    Console.WriteLine("\nCustom Properties:");
                    foreach (var property in customProperties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
                    {
                        Console.WriteLine($"{property.Name}: {property.InnerText}");
                        metadata.Add(property.Name!, property.InnerText!);
                    }
                }
                else
                {
                    Console.WriteLine("\nNo custom properties found.");
                }

                string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText("metadata.json", json);

            }
        }
    }
}
