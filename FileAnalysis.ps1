# Need a function to create a CSV if using the -CSVFile option in Get-FileAnalysis
Function Get-Files {
    #here's an example for now...
    Get-ChildItem -Path C:\ -Recurse -File | Select-Object FullName,Extension,LastWriteTime,Length | Export-Csv test.csv
}


Function Get-FileAnalysis () {
    <#
    .SYNOPSIS
    This Function analyses file system information and reports on storage usage

    .DESCRIPTION
    This function iterates through a local file system or from a properly formatted
    input CSV file listing and creates the following Excel reports:
        1. The size of Files by type and date
        2. The size of Files by type and size
        3. The number of files by type and date
        4. The number of files by type and size
    
    These reports are first captured as Powershell custom objects which are then exported
    to Excel with a table and chart.

    File types are predetermined (e.g. Document, Program, Database, etc.) and uses file
    extensions as the means of identification.

    .PARAMETER CSVFile
    Specify a CSV with the correct output from a Get-Childitem command. Multiple Files can also be 
    specified from the pipeline. Each file is processed separately with a report for each. See INPUT 
    section in full help to see how to create the input file(s).

    .PARAMETER Path
    Specifies a local path to a drive or folder structure to report on. Only one path can be specified.

    .EXAMPLE
    Get-FileAnalysis -Path C:\
    PS C:\>Get-FileAnalysis -Path E:\Data

    These examples will report on files on the local computer and create a single
    spreadsheet called "FileAnalysis.xlsx" in the users' documents folder.

    .EXAMPLE
    Get-FileAnalysis -CSVFile Server1-DriveD.csv
    PS C:\>Get-Childitem Server*.csv | Get-FileAnalysis
    
    PS C:\>"Server1-DriveD.csv","Server1-DriveE.csv" | Get-FileAnalysis

    These commands will report on input from specified file(s). A separate spread sheet is created for 
    each input file with a matching name.

    .INPUTS
    -Path requires a local drive or folder path

    -CSVfile requires a properly formatted CSV file or files

    To create a CSV file on a remote system, use Get-Childitem. The file must at least contain the following properties:
        -Extension
        -Length
        -LastWriteTime

    For Example: 
    PS C:\>Get-Childitem -Path D:\ -Recurse -File |
        Select-Object FullName,Extension,Length,LastWriteTime |
        Export-Csv Server1-DriveD.csv -NoTypeInformation

    .OUTPUTS
    An Excel spreadsheet with the same name as the input file, or a file called
    "FileAnalysis.xlsx" when using the -Path parameter.

    .NOTES
    This function requires Excel to be locally installed. Tested with Excel 2013.
    #>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$true,
                   ParameterSetName='FileSpecified')]   # Either -CSVfile of -Path can be specified, but not both
        [String]
        $CSVFile,

        [Parameter(Mandatory=$true,
                   Position=0,
                   ValueFromPipeline=$false,
                   ParameterSetName='PathSpecified')]
        [String]
        $Path
    )

    Begin {
        # Begin block runs once only
        # File type variables - used in the Get-FileType function - Yes, these are used - they're in a function called Get-FileType
        $MediaFiles = ".aif .asf .au .avi .p33 .m3u .mid .idi .miv .mov .mp2 .mp3 .mp4 .mpe .peg .mpg .mpeg .qt .rmi .snd .wav .wm .wma .wmv .p3a .p3b"
        $GraphicsFiles = ".3d2 .dmf .3ds .ai .art .bdf .bez .bmf .bmp .byu .cag .cam cdf .cdm .cpt .dcs .dem .dib .dkb .dlg .dwg .dxb .dxf .nff .eps .fac .fbm .fpx .fxd .eom .gif .gry .ham .hrf .iff .ges .img .imj .nst .iv .jas .big .jfi .fif .jpc .peg .jpg .jpeg .lbm .wob .mac .esh .mgf .mic .mng .mod .mrb .sdl .msp .nff .rbs .obj .oct .off .ogl .pbm .pcd .pct .pcx .pgm .pic .ict .ply .pnt .pol .pov .ppm .rop .psd .pub .uad .rad .ras .raw .ray .rgb .rib .rif .rwx .ene .scn .scr .sdl .dml .sgi .sgo .ade .shg .swf .iff .ddd .tga .tif .iff .oly .oly .rif .ect .vid .iff .wrl .x3d .xbm .odl .ydl .met .dc .pnf .jp2 .dgn .dxf .png"
        $BackupFiles = ".arc .arj .bac .bak .bck .bar .bkf .cab .cpt .dms .gl .gz .zip .ha .hpk .hqx .hyp .ish .lha .lzh .lzx .pak .pit .rar .saf .sea .har .shk .sit .sqz .tar .taz .tgz .uc2 .y .z .zip .zoo .old"
        $Documents = ".doc .docx .docm .xls .xlsx .xlsm .xlsb .xlw .dot .dotx .dotm .txt .mpp .oxps .xps .vsdx .vsd .rtf .xml .xlt .xltx .xltm xlam .ppt .pptx .pptm .pot .potx .ppam .ppsm .sldx .sldm .thmx .pdf .vss .vst .wwb .reg"
        $WebFiles = ".url .lnk .htm .html .js .jsp .css .mht .ico"
        $Databases = ".mdb .dbx .db .idx .xml .csv .log .dbd .bdb .dbf .mdx .mdf .ldf .dat"
        $BlockFiles = ".gho .iso .img .vmdk .qic .ghs .otf .mobi .epub"
        $Programs = ".exe .dll .nlm .jar .swf .rpm .c .dl_ .in_ .hlp .hl_ .cn_ .pr_ .ins .ini .hdr .bin .xp_ .inf .msi .ex_ .sys .bat .asp .chm .ps1 .psm1 .ps1xml .psd1 .vbs .cmd .hta .emf .json .pff"
        $Temp = ".tmp .temp"
        $Email = ".mbx .nsf .ns3 .ost .cca .pab .msg .trn .nk2"
        $PST = ".pst"

        $CollectionDate = Get-Date  # Used in Get-FileAge function
        $DateStamp = Get-Date -Format "yyMMddhhmm" 
        $AgeHeader = @('<7 days','7-14 days','14-21 days','21-28 days','28-60 days','60-90 days','90-120 days','120-180 days','180-365 days','1-2 years','2-3 years','3-4 years','4-5 years','5-6 years','6-7 years','>7 years')
        $SizeHeader =@('<1K','1K-8K','8K-16K','16K-32K','32K-64K','64K-128K','128K-256K','256K-512K','512K-1M','1M-2M','2M-4M','4M-8M','8M-16M','16M-32M','32M-50M','>50M')
    }

    Process {
        # Process block runs once per object on the pipeline (i.e. if multiple CSV's are piped to this function)

        # Create four 2 dimensional arrays to hold the output. Specifying long integer to avoid overrun.
        # Arrays must be the same dimensions so we can use the same function to convert the arrays to custom objects (ConvertTo-PSObject) 
        $SizebyDateArr = New-Object 'long[,]' 16,12
        $NumberbyDateArr = New-Object 'long[,]' 16,12
        $SizeBySizeArr = New-Object 'long[,]' 16,12
        $NumberBySizeArr = New-Object 'long[,]' 16,12

        # Determine whether a path was specified or a CSV file, generate the command to use later
        if($Path) {
            $Path = $Path.ToUpper()
            if($Path.EndsWith(":")) {
                $Path = $Path + "\"
            }
            $Command = "Get-Childitem -Path $Path -File -Recurse"
            $OutFile = "FileAnalysis-$DateStamp"
            $TitlePrefix = "Computer $env:COMPUTERNAME - $Path"
        }
        ElseIf ($CSVFile) {
            $Command = "Import-Csv -Path $CSVFile"
            $OutFile = "$(([io.fileinfo]$CSVFile).BaseName)-$DateStamp"  # Using the basename of the input file(s).
            $TitlePrefix = $OutFile # We assume some kind of descriptive filename, like "Server1 - Cdrive"
        }

        # Specify the current users' home folder for the output file(s).
        $OutFile = "$Home\Documents\$OutFile.xlsx"

        # Delete the output file if one already exists. Nasty things happen to the file otherwise.
        # This will be unlikely now that a datestamp is added to the filename
        Try {
            $ErrorActionPreference = "Stop"  #Make all errors terminating
            If (Test-Path $OutFile) {
                Remove-Item $OutFile -Force
                Write-Output "`n$OutFile exists, deleted it"
            }
        }
        Catch {
            Write-Output "`nFound an existing output file $OutFile but can't delete it. Close the file and rename it if you want to keep it."
            #Move onto the next input file if there is one
            Return
        }
        Finally {
            #Allow non-terminating errors again, this is necessary because get-childitem may throw some permissions errors
            $ErrorActionPreference = "Continue"
        }

        # Loop through file names and get their extension and age. Add these values to the appropriate array element
        Write-Output "`nProcessing '$Command'..."
        Invoke-Expression $Command | Foreach-object {
            # Make sure Length Property is an integer - when piping get-childitem the object is file system object an everything works
            # However, when piping from a Csv the object is a custom object and length property is a string. This causes unexpected results in FileSize function 
            # (i.e. string comparison rather than number)
            $Length = [long]$_.Length                         # Convert outside of Get-FileSize function because its also used to add to array here                       

            # Find the correct array indexes for the file - calls 3 helps functions
            $FileType = Get-FileType(($_.Extension).ToLower())
            $FileAge = Get-FileAge($_.LastWriteTime)
            $FileSize = Get-FileSize($Length)

            # Update the input arrays
            $SizebyDateArr[$FileAge,$FileType] += ($Length)   # Add file to size by date array
            $NumberbyDateArr[$FileAge,$FileType] += 1         # Increment files by date array
            $SizeBySizeArr[$FileSize,$FileType] += ($Length)  # Add file to size by size array
            $NumberBySizeArr[$FileSize,$FileType] += 1        # Increment files by size array
        }

        # Now convert the arrays to Powershell Custom Objects mostly for formatting purposes
        $ObjSizebyDate = ConvertTo-PSObject -Header $AgeHeader -DataArray $SizebyDateArr -Format "0.00" -Divider 1mb
        $ObjNumberbyDate = ConvertTo-PSObject -Header $AgeHeader -DataArray $NumberbyDateArr -Format "0"
        $ObjSizebySize = ConvertTo-PSObject -Header $SizeHeader -DataArray $SizebySizeArr -Format "0.00" -Divider 1mb
        $ObjNumberbySize = ConvertTo-PSObject -Header $SizeHeader -DataArray $NumberbySizeArr -Format "0"

        # Here are the 4 custom objects
        #$ObjSizebyDate | Out-GridView
        #$ObjNumberbyDate | Out-GridView
        #$ObjSizebySize | Out-GridView
        #$ObjNumberbySize | Out-GridView

        # Export the custom objects to Excel and display a chart for each one
        New-ExcelWithChart -Data $ObjSizebyDate -XLabel "File Size (MB)" -YLabel "Last Modified Date" -ChartTitle "$TitlePrefix : File Size By Modified Date" -WorksheetName "File Size (by date)" -ExcelFileName $OutFile
        New-ExcelWithChart -Data $ObjSizebySize -Xlabel "File Size (MB)" -YLabel "File Size" -ChartTitle "$TitlePrefix : File Size By Size" -WorksheetName "File Size (by size)" -ExcelFileName $OutFile
        New-ExcelWithChart -Data $ObjNumberbyDate -XLabel "Number of Files" -YLabel "Last Modified Date" -ChartTitle "$TitlePrefix : Number of Files By Modified Date" -WorksheetName "Number of files (by date)" -ExcelFileName $OutFile
        New-ExcelWithChart -Data $ObjNumberbySize -XLabel "Number of Files" -YLabel "File Size" -ChartTitle "$TitlePrefix : Number of Files By Size" -WorksheetName "Number of files (by size)" -ExcelFileName $OutFile 
    }

    End {
        Write-Output "Done"
        Invoke-Item $OutFile  #Show the last spreadsheet created. If -Path is used, there will only be one.
    }
}
Function Get-FileType ($ext) {
    # Helper Function for Get-FileAnalysis - returns an integer based on the type of file (acts as index for the input arrays)
    # Uses variables from parent scope, but we don't change any of them, so its OK that they are local and not global (no copy-on-write)
    # (However, PSScriptAnalyzer marks them as not used because they are never referenced in the parent function)
    if ($MediaFiles.Contains($ext)) {0}
    elseif ($GraphicsFiles.Contains($ext)) {1}
    elseif ($BackupFiles.Contains($ext)) {2}
    elseif ($Documents.Contains($ext)) {3}
    elseif ($WebFiles.Contains($ext)) {4}
    elseif ($Databases.Contains($ext)) {5}
    elseif ($BlockFiles.Contains($ext)) {6}
    elseif ($Programs.Contains($ext)) {7}
    elseif ($Temp.Contains($ext)) {8}
    elseif ($Email.Contains($ext)) {9}
    elseif ($PST.Contains($ext)) {10}
    else {11}
}
Function Get-FileAge ($FileAge) {
    # Helper Function for Get-FileAnalysis - returns an integer based on age of file (acts as index for input arrays)
    $DaysOld = New-TimeSpan -start $FileAge -End $CollectionDate | Select-Object -ExpandProperty Days
    If ($DaysOld -lt 7) {0}
    Elseif ($DaysOld -lt 14) {1}
    Elseif ($DaysOld -lt 21) {2}
    Elseif ($DaysOld -lt 28) {3}
    Elseif ($DaysOld -lt 60) {4}
    Elseif ($DaysOld -lt 90) {5}
    Elseif ($DaysOld -lt 120) {6}
    Elseif ($DaysOld -lt 180) {7}
    Elseif ($DaysOld -lt 365) {8}
    Elseif ($DaysOld -lt 730) {9}
    Elseif ($DaysOld -lt 1095) {10}
    Elseif ($DaysOld -lt 1460) {11}
    Elseif ($DaysOld -lt 1825) {12}
    Elseif ($DaysOld -lt 2190) {13}
    Elseif ($DaysOld -lt 2555) {14}
    Else {15}
}
Function Get-FileSize ($FileSize) {
    # Helper Function for Get-FileAnalysis - returns an integer based on size of file (acts as index for input arrays)
    # Note - the filesize parameter (in kilobytes) needs to be a number, not a string. Convert to long before calling.
    If ($FileSize -lt 1024) {0}
    ElseIf ($FileSize -lt 8192) {1}
    ElseIf ($FileSize -lt 16384) {2}
    ElseIf ($FileSize -lt 32768) {3}
    ElseIf ($FileSize -lt 65536) {4}
    ElseIf ($FileSize -lt 131072) {5}
    ElseIf ($FileSize -lt 262144) {6}
    ElseIf ($FileSize -lt 524288) {7}
    ElseIf ($FileSize -lt 1048576) {8}
    ElseIf ($FileSize -lt 2097152) {9}
    ElseIf ($FileSize -lt 4194304) {10}
    ElseIf ($FileSize -lt 8338608) {11}
    ElseIf ($FileSize -lt 16777216) {12}
    ElseIf ($FileSize -lt 33554432) {13}
    ElseIf ($FileSize -lt 57108864) {14}
    Else {15}
}
Function ConvertTo-PSObject() {
    # Helper Function for Get-FileAnalysis - converts an array into a custom PowerShell Object. Expects 4 arguments, a header array with
    # 16 elements, which becomes the first row of the object, a two dimensional array 16 X 12 in size containing the data, the format
    # of the output string, and an optional denominator to obtain the correct size formatting.
    Param (
        [String[]]$Header,      # 1 dimensional array
        [Long[,]]$DataArray,    # 2 dimensional array with long datatype
        [String]$Format,        # ToString format e.g. "0" and "0.00"
        [Int]$Divider           # Optional denominator value to produce desired size conversation - e.g. 1mb
    )

    If(-not($Divider)) {
        $Divider = 1
    }
    For ($i = 0; $i -le 15; $i++) {
        [pscustomobject]@{
            'Size (MB)' = $Header[$i]
            'Media Files' = ($DataArray[$i,0] / $Divider).ToString($Format)
            'Graphics Files' = ($DataArray[$i,1] / $Divider).ToString($Format)
            'Backup Files' = ($DataArray[$i,2] / $Divider).ToString($Format)
            'Documents' = ($DataArray[$i,3] / $Divider).ToString($Format)
            'WebFiles' = ($DataArray[$i,4] / $Divider).ToString($Format)
            'Databases' = ($DataArray[$i,5] / $Divider).ToString($Format)
            'BlockFiles' = ($DataArray[$i,6] / $Divider).ToString($Format)
            'Programs' = ($DataArray[$i,7] / $Divider).ToString($Format)
            'Temp' = ($DataArray[$i,8] / $Divider).ToString($Format)
            'Email' = ($DataArray[$i,9] / $Divider).ToString($Format)
            'Others' = ($DataArray[$i,11] / $Divider).ToString($Format)
            'PST' = ($DataArray[$i,10] / $Divider).ToString($Format)
        }
    }
}
Function New-ExcelWithChart () {
    # This function opens a new or existing spreadsheet and adds a new worksheet to it.
    # It would be more efficient to keep the Excel session open and keep appending until
    # done, but I wanted to make this funtion self contained for other uses.
    #
    # I looked at Import-Excel module, but found it limiting with charts. https://github.com/dfinke/ImportExcel
    Param
    ( 
        [Parameter(Mandatory=$true,Position=0)]
        $Data,

        [Parameter(Mandatory=$true,Position=1)]
        [String]$XLabel,

        [Parameter(Mandatory=$true,Position=2)]
        [String]$YLabel,

        [Parameter(Mandatory=$true,Position=3)]
        [String]$ChartTitle,

        [Parameter(Mandatory=$true,Position=4)]
        [String]$WorksheetName,

        [Parameter(Mandatory=$true,Position=5)]
        $ExcelFileName
    )
    
    #Open an Excel session
    $xl = new-object -ComObject Excel.Application

    # Some Constants - Not using these
    # $xlConditionValues=[Microsoft.Office.Interop.Excel.XLConditionValueTypes]
    # $xlTheme=[Microsoft.Office.Interop.Excel.XLThemeColor]
    # $xlChart=[Microsoft.Office.Interop.Excel.XLChartType]
    # $xlIconSet=[Microsoft.Office.Interop.Excel.XLIconSet]
    # $xlDirection=[Microsoft.Office.Interop.Excel.XLDirection]

    If(Test-Path -Path $ExcelFileName) {
        #Excel file exists, so open it. Assumes its an Excel file (No error checking)
        Write-Output "$ExcelFileName already exists, appending report '$ChartTitle'"
        $wb = $xl.Workbooks.Open($ExcelFileName)
        $ws = $xl.Worksheets.add()
    }
    Else {
        #Excel file does not exist, so open a new one
        Write-Output "$ExcelFileName does not exist, creating report '$ChartTitle'"
        $wb = $xl.workbooks.add()
        $ws = $wb.activesheet
    }

    # Some preliminary settings, name the worksheet, hide Excel.
    $xl.Visible = $False
    $xl.DisplayAlerts = $False  # You need this, otherwise you are prompted to overwrite. Save() method is better than SaveAs() (no prompt), but need SaveAs() for a new file
    $ws.Name = $WorksheetName

    #Populate the data onto worksheet, This method rocks because the range is already selected for the chart later
    $Data | ConvertTo-CSV -NoTypeInformation -Delimiter "`t" | c:\windows\system32\clip.exe
    $ws.Range("A1").Select | Out-Null
    $ws.Paste()

    #Create Chart 
    #$chart=$xl.Charts.Add()                  # Create a chart in a new worksheet
    #$chart.Name = "Chart-$WorksheetName"     # If creating a chart in a new worksheet, name the worksheet
    $chart=$ws.Shapes.AddChart().Chart        # or, embed a chart into the current worksheet
    
    #Format the Chart
    $chart.ChartType = 55  #$xlChart::xl3DColumnStacked doesn't work. If you pipe this to | Select -ExpandProperty value__ you get 55, but this doesn't work either.
    $chart.ChartStyle = 34 #Favourite colours, background etc.
    $chart.ChartColor = 2
    $chart.Perspective = 20
    $chart.Rotation = 30
    $chart.HeightPercent = 100

    #Modify the chart title and Axis labels
    $chart.HasTitle = $true
    $chart.ChartTitle.Text = $ChartTitle
    # X axis - xlvalue=2
    $chart.Axes(2).HasTitle = $true
    $chart.Axes(2).AxisTitle.Text = $XLabel
    $chart.Axes(2).AxisTitle.Font.Name = "Ariel"
    $chart.Axes(2).AxisTitle.Font.Size = 12
    # Y axis - xlCategory=1
    $chart.Axes(1).HasTitle = $true
    $chart.Axes(1).AxisTitle.Text = $YLabel
    $chart.Axes(1).AxisTitle.Font.Name = "Ariel"
    $chart.Axes(1).AxisTitle.Font.Size = 12

    # Only need these commands if formatting an embedded chart
    $ws.Shapes.Item("Chart 1").Left = 650
    $ws.Shapes.Item("Chart 1").Top = 5
    $ws.Shapes.Item("Chart 1").Width = 700
    $ws.Shapes.Item("Chart 1").Height = 500
    $ws.Shapes.Item("Chart 1").Chart.ChartArea.RoundedCorners = $true

    # Save and Close the workbook
    $xl.ActiveWorkbook.SaveAs($ExcelFileName)  #I guess I could use Save() for existing and SaveAs(...) for new, but this works with alerts off, see above
    # $xl.ActiveWorkbook.Close() | Out-Null
    $xl.Quit()
}

# --- End ---