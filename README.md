# Word Template Editor
You can edit either a single file, or every .dotx file in a directory (including subdirectories). 
It replaces a given text with another one in the footer of each document, and saves it again to the original location.

Logs are generated in Appdata\Roaming.
## WARNING: Prerelease
This tool currently has only a prerelease version available for download. Expect bugs and quirks. Also, the CSV Import feature is currently non-existing and will be added in a later version.


## Credit
- Uses the OpenXML SDK from Microsoft for file operations on Office-Files.

- Uses Serilog for Logging.

- Uses System.CommandLine (a prerelease of the official development version from Microsoft.)