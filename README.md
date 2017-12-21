# FileAnalysis

This PowerShell module runs on Windows systems and can be used to produce nice reports in Excel showing the file usage on one or more disks. The original intent was to understand file types, their size and relative age so that third party file archiving rules could be sensibly established before turning on the archiving software.

The module contains two exported functions, as follows:

1. **Get-FileListing** - Generates a file listing of a target path and creates a CSV file with the required output attributes for the Get-FileAnalysis function. The function uses Get-ChildItem and captures the file name, size, extension and modified date of each file in the specfied path. The idea is to use this function on your servers either locally or via a remote session. Once you have a bunch of CSV files, you can pipe them to the Get-FileAnalysis function on your local machine. 

2. **Get-FileAnalysis** - Requires Excel. This function either takes a local path or one or more CSV files as parameters and generates 4 reports in Excel for each target specfied:
  * The size of Files by type and date
  * The size of Files by type and size
  * The number of files by type and date
  * The number of files by type and size

These 4 reports are created in a single Excel Spreadsheet; one spreadsheet per specified path.

Here is an example report:
![Number of Files by Size](/Media/NumberFilesBySize.PNG)

# Instructions

```powershell
1. One time setup
    # Download the repository
    # Unblock the files
    # Extract the files to a module path (e.g. $env:USERPROFILE\Documents\WindowsPowerShell\Modules\FileAnalysis)

2. Import the module.
    Import-Module FileAnalysis

3. Get commands in the module
    Get-Command -Module FileAnalysis

4. Get help for a command
    Get-Help Get-FileAnalysis
```
