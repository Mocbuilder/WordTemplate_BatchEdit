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

        public static void StartCLI(string[] args)
        {
            Logo();

            var rootCommand = new RootCommand("CLI tool for working with .dotx files");
            var srCommand = new Command("sr", "Perform search and replace in a single file")
            {
                new Option<string>("--path", "Path to the .dotx file") { IsRequired = true },
                new Option<string>("--search", "Text to search for") { IsRequired = true },
                new Option<string>("--replace", "Text to replace with") { IsRequired = true }
            };

srCommand.SetHandler<string, string, string>(
                HandleSearchAndReplace,
                srCommand.Options[0],
                srCommand.Options[1],
                srCommand.Options[2]);
            rootCommand.AddCommand(srCommand);
        }

        public static void HandleSearchAndReplace(string path, string search, string replace)
        {
            // Perform the logic here (synchronous)
            FileOps.DocSRFooter(path, search, replace);

            // Return Task.CompletedTask to satisfy the async signature
            //return Task.CompletedTask;
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
