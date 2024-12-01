using System.IO;
using System.Reflection.Metadata;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using static System.Net.Mime.MediaTypeNames;
using Serilog;
using Microsoft.VisualBasic;
using System.Configuration;

namespace WordTemplate_BatchEdit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            string logConfig = ConfigurationManager.AppSettings["GenerateLog"];
            bool generateLog = bool.TryParse(logConfig, out bool result) && result;

            string costumLogDir = ConfigurationManager.AppSettings["CostumLogDir"];

            string costumLogDirConfig = ConfigurationManager.AppSettings["UseCostumLogDir"];
            bool useCostumLog = bool.TryParse(costumLogDirConfig, out bool result2) && result2;


            string filePath;

            if(useCostumLog == true) 
            {
                filePath = Path.Combine(costumLogDir, "WordTemplate_BatchEdit", "log-.txt");
            }
            else
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                filePath = Path.Combine(appDataPath, "WordTemplate_BatchEdit", "log-.txt");

            }

            Directory.CreateDirectory(Path.GetDirectoryName(filePath));

            if (generateLog == true)
            {
                Log.Logger = new LoggerConfiguration()
                .WriteTo.File(filePath, rollingInterval: RollingInterval.Hour)
                .CreateLogger();
            }
            else
            {
                Log.Logger = new LoggerConfiguration()
                .CreateLogger();
            }

            try
            {
                Log.Information($"Log from: {DateTime.Now}");
                Log.Information("Application started.");

                CLI.StartCLI(args);
                GetUserInput();
            }
            catch (FileFormatException fileFormatEx)
            {
                Log.Error(fileFormatEx, "An error occured.");
            }
            catch (Exception ex)
            {
                Log.Error(ex, "An error occurred.");
            }
            finally
            {
                Log.CloseAndFlush();
            }
            Console.ReadLine();
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

                        while (string.IsNullOrWhiteSpace(rootDirPath) || !Directory.Exists(rootDirPath))
                        {
                            Console.WriteLine("Invalid directory. Please enter a valid directory path:");
                            rootDirPath = Console.ReadLine();

                            if (rootDirPath.Equals("exit", StringComparison.OrdinalIgnoreCase))
                            {
                                Environment.Exit(0);
                            }
                        }

                        string[] files = Directory.GetFiles(rootDirPath, "*.dotx", SearchOption.AllDirectories);

                        for (int i = 0; i < files.Length; i++)
                        {
                            string tempGetName = files[i].Split('\\').Last();

                            if (tempGetName.Contains("MGLG"))
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
                            Console.WriteLine($"{path} is at run");
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
            Log.Information($"Succesfull Edit: {path}");
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

        public static void LogWrite(string message)
        {
            
        }
    }
}
