﻿using System;
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
            var sr_SingleFileCommand = new Command("read", "Read and display the file.")
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
                await META_SingleFile_GetMetaData(singleFile!, dump, output);
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

        static async Task META_SingleFile_GetMetaData(FileInfo file, string dump, string outputPath)
        {
            Dictionary<string, string> metadata = new Dictionary<string, string>();

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
