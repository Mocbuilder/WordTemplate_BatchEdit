using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
            string fileName = $"metadata_{file.Name}_{Random.Shared.Next()}.json";

            //if directory (as in: a valid FOLDER) doesnt exist, you'll create a file to dump it at that location. That way users cant specify an existing file, but who dumps into an already existing file anyway ?
            if (!Directory.Exists(output))
            {
                output = Directory.GetCurrentDirectory();
            }

            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(file.FullName, false))
            {
                var coreProperties = wordDoc.PackageProperties;

                //Dictionary for the printing AND dumping \O/
                var metadata = new Dictionary<string, string?>
                {
                    { "Title", coreProperties.Title },
                    { "Author", coreProperties.Creator },
                    { "Subject", coreProperties.Subject },
                    { "Description", coreProperties.Description },
                    { "Keywords", coreProperties.Keywords },
                    { "LastModBy", coreProperties.LastModifiedBy },
                    { "CreatedDate", coreProperties.Created?.ToString("o") },
                    { "ModifiedDate", coreProperties.Modified?.ToString("o") }
                };

                Console.WriteLine("Core Properties:");
                Console.WriteLine(string.Join(Environment.NewLine, metadata.Select(kv => $"{kv.Key}: {kv.Value}")));

                var customProperties = wordDoc.ExtendedFilePropertiesPart?.Properties;

                if (customProperties.HasChildren == false)
                {
                    Console.WriteLine("\nNo custom properties found.");
                }

                Console.WriteLine("\nCustom Properties:");
                foreach (var property in customProperties.Elements<DocumentFormat.OpenXml.CustomProperties.CustomDocumentProperty>())
                {
                    Console.WriteLine($"{property.Name}: {property.InnerText}");
                    metadata.Add(property.Name!, property.InnerText!);
                }

                string json = JsonSerializer.Serialize(metadata, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(Path.Combine(output, fileName), json);
            }
        }
    }
}
