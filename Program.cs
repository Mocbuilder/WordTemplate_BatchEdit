using System.IO;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;

namespace WordTemplate_BatchEdit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            GetUserInput();
        }

        public static void GetUserInput()
        {
            while (true)
            {
                string userInput = "";
                int userChoice = 0;

                Console.WriteLine("Enter the path to a single doc (1) or path to a CSV with multiple paths (2) or directory with multiple docs (3) ?");

                try
                {
                    userInput = Console.ReadLine();
                    if (userInput == "exit") Environment.Exit(0);
                    userChoice = Convert.ToInt32(userInput);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                }

                switch (userChoice)
                {
                    case 1:
                        Console.WriteLine("Enter path to document:");
                        string docPath = Console.ReadLine();

                        Console.WriteLine("Enter Text that should be replaced:");
                        string toSearch = Console.ReadLine();

                        Console.WriteLine("Enter Text that will replace it:");
                        string toReplace = Console.ReadLine();


                        SearchAndEditDoc(docPath, toSearch, toReplace);
                        break;
                    case 2:
                        Console.WriteLine("Enter path to CSV:");
                        string csvPath = Console.ReadLine();
                        //GetPathsFromCSV(csvPath);
                        break;
                    case 3:
                        Console.WriteLine("Enter directory that contains the docs:");
                        string rootDirPath = Console.ReadLine();

                        string[] files = Directory.GetFiles(rootDirPath, "*.dotx", SearchOption.AllDirectories);

                        for (int i = 0; i < files.Length; i++)
                        {
                            string tempGetName = files[i].Split('\\').Last();

                            if(tempGetName.Contains("MGLG"))
                            {
                                files[i] = "";
                            }
                        }

                        MultiDocEdit(files);

                        break;
                    default:
                        Console.WriteLine("Invalid choice, please try again.");
                        break;
                }
            }
        }

        public static void SearchAndEditDoc(string path, string toSearch, string toReplace)
        {
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
                            foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
                            {
                                if (text.Text.Contains(toSearch))
                                {
                                    text.Text = text.Text.Replace(toSearch, toReplace);
                                }
                            }
                        }
                    }
                }

                mainPart.Document.Save();
                Console.WriteLine($"Footer updated successfully at {path}");
            }

        }

        public static void MultiDocEdit(string[] files)
        {
            Console.WriteLine("Enter Text that should be replaced:");
            string toSearch = Console.ReadLine();

            Console.WriteLine("Enter Text that will replace it:");
            string toReplace = Console.ReadLine();

            foreach (var file in files)
            {
                if (file == "") continue;
                else if (!File.Exists(file)) { Console.WriteLine($"{file} is not an existing directory."); continue; }

                SearchAndEditDoc(file, toSearch, toReplace);
            }

            Console.WriteLine("Succesfully editet all files in the directory.");
        }
    }
}
