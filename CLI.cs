using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.CommandLine;
using Serilog;

namespace WordTemplate_BatchEdit
{
    public class CLI
    {
        public static string userInput;

        public static async void StartCLI(string[] args)
        {
            Logo();

            var rootCommand = new RootCommand("A tool for batch editing .dotx files (Word Document Templates).");

            var singleFileOption = new Option<FileInfo?>(
                name: "--file",
                description: "The file to read and display on the console."
            )
            {
                IsRequired = true,
            };

            var sr_sectionOption = new Option<string>(
                name: "--section",
                description: "The section of the file to process (head, body, footer)."
            )
            {
                IsRequired = true
            };
            sr_sectionOption.AddCompletions("head", "body", "footer");

            var sr_searchOption = new Option<string>(
                name: "--search",
                description: "The text that is going to be searched for."
            )
            {
                IsRequired = true
            };

            var sr_replaceOption = new Option<string>(
                name: "--search",
                description: "The text that is going to be searched for."
            )
            {
                IsRequired = true
            };

            var srSingleFileCommand = new Command("read", "Read and display the file.")
            {
                singleFileOption,
                sr_sectionOption,
                sr_searchOption,
                sr_replaceOption
            };
            rootCommand.AddCommand(srSingleFileCommand);

            srSingleFileCommand.SetHandler(async (singleFile, sectionOption, sr_searchOption, sr_replaceOption) =>
            {
                await sr_SingleFilePartRouting(singleFile!, sectionOption, sr_searchOption, sr_replaceOption);
            }, singleFileOption, sr_sectionOption, sr_searchOption, sr_replaceOption);

            await rootCommand.InvokeAsync(args);
        }

        static async Task sr_SingleFilePartRouting(
            FileInfo file, string sectionOption, string search, string replace)
        {
            switch (sectionOption)
            {
                case "header":
                    FileOps.SR_SingleFileHeader(file.FullName, search, replace);
                    break;
                case "body":
                    FileOps.SR_SingleFileBody(file.FullName, search, replace);
                    break;
                case "footer":
                    FileOps.SR_SingleFileFooter(file.FullName, search, replace);
                    break;
                case "all":
                    FileOps.SR_SingleFileAll(file.FullName, search, replace);
                    break;
                default:
                    FileOps.SR_SingleFileAll(file.FullName, search, replace);
                    break;
            }
        }

        public static void Logo()
        {
            Console.WriteLine(@"
  _____        _  __   __           ____        _       _     ______    _ _ _             
 |  __ \      | | \ \ / /          |  _ \      | |     | |   |  ____|  | (_) |            
 | |  | | ___ | |_ \ V /   ______  | |_) | __ _| |_ ___| |__ | |__   __| |_| |_ ___  _ __ 
 | |  | |/ _ \| __| > <   |______| |  _ < / _` | __/ __| '_ \|  __| / _` | | __/ _ \| '__|
 | |__| | (_) | |_ / . \           | |_) | (_| | || (__| | | | |___| (_| | | || (_) | |   
 |_____/ \___/ \__/_/ \_\          |____/ \__,_|\__\___|_| |_|______\__,_|_|\__\___/|_|   
                                                                                          
                                                                                          
            
            ");
        }
    }
}
