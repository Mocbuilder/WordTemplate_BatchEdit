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

namespace WordTemplate_BatchEdit.FileOps
{
    public class SR_FileOps
    {
        public static void SR_SingleFileAll(string path, string search, string replace)
        {
            Log.Information("SR_SingleFileAll invoked...");
            SR_SingleFileHeader(path, search, replace);
            Log.Information("SR_SingleFileAll: SR_SingleFileHeader finished...");
            SR_SingleFileBody(path, search, replace);
            Log.Information("SR_SingleFileAll: SR_SingleFileBody finished...");
            SR_SingleFileFooter(path, search, replace);
            Log.Information("SR_SingleFileAll: SR_SingleFileFooter finished...");
            Log.Information("SR_SingleFileAll finished.");
        }

        public static void SR_SingleFileFooter(string path, string search, string replace)
        {
            if (!File.Exists(path)) { Console.WriteLine($"File path '{path}' is invalid"); return; }

            Log.Information($"SR_SingleFileFooter: Editing file: {path}");
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
                            Log.Information($"SR_SingleFileFooter: {path} is at run");
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
                Log.Information($"SR_SingleFileFooter: Footer updated successfully at {path}");
            }
            Log.Information($"SR_SingleFileFooter: Succesfull Edit at {path}");
        }

        public static void SR_SingleFileHeader(string path, string search, string replace)
        {
            if (!File.Exists(path)) { Console.WriteLine($"File path '{path}' is invalid"); return; }

            Log.Information($"SR_SingleFileHeader: Editing file: {path}");
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(path, true))
            {
                var mainPart = wordDoc.MainDocumentPart;

                var headerParts = mainPart.HeaderParts;

                foreach (var headerPart in headerParts)
                {
                    var header = headerPart.Header;

                    foreach (var paragraph in header.Elements<Paragraph>())
                    {
                        foreach (var run in paragraph.Elements<Run>())
                        {
                            Log.Information($"SR_SingleFileHeader: {path} is at run");
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
                Log.Information($"SR_SingleFileHeader: Header updated successfully at {path}");
            }
            Log.Information($"SR_SingleFileHeader: Succesfull Edit at {path}");
        }

        public static void SR_SingleFileBody(string path, string search, string replace)
        {
            if (!File.Exists(path)) { Console.WriteLine($"File path '{path}' is invalid"); return; }

            Log.Information($"SR_SingleFileBody: Editing file: {path}");
            using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(path, true))
            {
                var mainPart = wordDoc.MainDocumentPart;

                if (mainPart != null)
                {
                    var body = mainPart.Document.Body;

                    if (body != null)
                    {
                        foreach (var textElement in body.Descendants<DocumentFormat.OpenXml.Wordprocessing.Text>())
                        { 
                            if (textElement.Text.Contains(search))
                            {
                                textElement.Text = textElement.Text.Replace(search, replace);
                            }
                        }
                        mainPart.Document.Save();
                    }
                }
            }
            Log.Information($"SR_SingleFileBody: Succesfull Edit at {path}");
        }
    }
}
