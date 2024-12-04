using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.CommandLine;
using Serilog;
using WordTemplate_BatchEdit.FileOps;
using DocumentFormat.OpenXml.Packaging;
using System.Text.Json;

namespace WordTemplate_BatchEdit
{
    public class CLI
    {
        public static string userInput;

        public static async void InitCLI(string[] args)
        {
            Logo();

            var rootCommand = new RootCommand("A tool for batch editing .dotx files (Word Document Templates).");

            #region Options
            var singleFileOption = new Option<FileInfo?>(
                name: "--file",
                description: "The file to read and display on the console."
            )
            {
                IsRequired = true,
            };

            var sr_sectionOption = new Option<string>(
                name: "--section",
                description: "The section of the file to process (header, body, footer, all)."
            )
            {
                IsRequired = true
            };
            sr_sectionOption.AddCompletions("header", "body", "footer", "all");

            var sr_searchOption = new Option<string>(
                name: "--search",
                description: "The text that is going to be searched for."
            )
            {
                IsRequired = true
            };

            var sr_replaceOption = new Option<string>(
                name: "--replace",
                description: "The text that is going to be searched for."
            )
            {
                IsRequired = true
            };

            var meta_dumpOption = new Option<string>(
                name: "--dump",
                description: "Dump the meta-data as json to the output directory"
            );
            meta_dumpOption.AddCompletions("true", "false");

            var meta_outputPathOption = new Option<string>(
                name: "--output",
                description: "The output directory for the meta-data dump file."
            );
            #endregion Options

            #region Commands
            var sr_SingleFileCommand = new Command("sr", "Search and replace text in a dotx file.")
            {
                singleFileOption,
                sr_sectionOption,
                sr_searchOption,
                sr_replaceOption
            };
            rootCommand.AddCommand(sr_SingleFileCommand);

            sr_SingleFileCommand.SetHandler(async (singleFile, sectionOption, sr_searchOption, sr_replaceOption) =>
            {
                await SR_SingleFile_PartSpecificRouting(singleFile!, sectionOption, sr_searchOption, sr_replaceOption);
            }, singleFileOption, sr_sectionOption, sr_searchOption, sr_replaceOption);


            var meta_SingleFileCommand = new Command("meta", "Get and optionaly dump the meta-data of a .dotx file.")
            {
                singleFileOption,
                meta_dumpOption,
                meta_outputPathOption
            };
            rootCommand.AddCommand(meta_SingleFileCommand);

            meta_SingleFileCommand.SetHandler(async (singleFile, dump, output) =>
            {
                await META_FileOps.META_SingleFile_GetMetaData(singleFile!, dump, output);
            }, singleFileOption, meta_dumpOption, meta_outputPathOption);
            #endregion Commands

            await rootCommand.InvokeAsync(args);
        }

        static async Task SR_SingleFile_PartSpecificRouting(
            FileInfo file, string sectionOption, string search, string replace)
        {
            switch (sectionOption)
            {
                case "header":
                    SR_FileOps.SR_SingleFileHeader(file.FullName, search, replace);
                    break;
                case "body":
                    SR_FileOps.SR_SingleFileBody(file.FullName, search, replace);
                    break;
                case "footer":
                    SR_FileOps.SR_SingleFileFooter(file.FullName, search, replace);
                    break;
                case "all":
                    SR_FileOps.SR_SingleFileAll(file.FullName, search, replace);
                    break;
                default:
                    SR_FileOps.SR_SingleFileAll(file.FullName, search, replace);
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
