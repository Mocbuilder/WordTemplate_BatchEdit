using DocumentFormat.OpenXml.Packaging;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;
using System.Configuration;

namespace WordTemplate_BatchEdit
{
    public class FileOps
    {
        public static void DocSRFooter(string path, string search, string replace)
        {
            if (!File.Exists(path)) { Console.WriteLine($"File path '{path}' is invalid"); return; }

            Log.Information($"Editing file: {path}");
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(path, true))
            {
                var mainPart = wordDoc.MainDocumentPart;

                var footerParts = mainPart.FooterParts;

                foreach (var footerPart in footerParts)
                {
                    var footer = footerPart.Footer;

                    foreach (var paragraph in footer.Elements<Paragraph>())
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            Log.Information($"{path} is at run");
                            foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
                            {
                                if (text.Text.Contains(search))
                                {
                                    text.Text = text.Text.Replace(search, replace);
                                }
                            }
                        }
                    }
                }

                mainPart.Document.Save();
                Log.Information($"Footer updated successfully at {path}");
            }
            Log.Information($"Succesfull Edit: {path}");
        }
    }
}
