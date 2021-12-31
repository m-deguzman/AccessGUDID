
Clear-Host
Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0
Get-Date
## ---------- Prompt to Query Loop ------------------------------------------
$About = @"

This script can be used to get a list of medical devices from FDA's GUDID.
(GUDID) Global Unique Device Identification Database contains
key device identification information submitted to the FDA. 

It will build a local SQLite database using open source tools.
Once the GUDID SQLite DB is built, you can query for devices.
Query result is saved in csv format and imported to Excel.

"@
Write-Host $About -ForegroundColor White -BackgroundColor Blue
PAUSE

if(-not $PSScriptRoot) {Write-Host "Script Root Dir: $PSScriptRoot is not defined" -ForegroundColor White -BackgroundColor Red; break} 
Set-Location $PSScriptRoot

## ---- sqlite utility ------------------
$sqlite_uri = "https://sqlite.org/2021/sqlite-tools-win32-x86-3360000.zip"
$sqlite3 = Join-Path -Path $PSScriptRoot -ChildPath "sqlite3.exe"

## ---- sqlite core dll ------------------
$version = '1.0.112'
$sqlite_core_uri = "https://www.nuget.org/api/v2/package/System.Data.SQLite.Core/$version"
$sqlite_core  =  Join-Path -Path $PSScriptRoot -ChildPath "System.Data.SQLite.dll"
$sqlite_interop  =  Join-Path -Path $PSScriptRoot -ChildPath "SQLite.Interop.dll"

## --------- Required Files Check -------------------
$gudid_full_date = "AccessGUDID_Delimited_Full_Release_" +  (Get-Date -Format "yyyyMM") + "01"
$gudid_file_zip = $gudid_full_date + ".zip"
$gudid_uri = "https://accessgudid.nlm.nih.gov/release_files/download/" + $gudid_file_zip
$gudid_db_name = $gudid_full_date + ".sqlite.db"
$gudid_db_full_path = Join-Path -Path $PSScriptRoot -ChildPath $gudid_db_name
If (-not(Test-Path $sqlite3)) {Write-Host "$sqlite3 not found." -ForegroundColor Red -BackgroundColor Black}
If (-not(Test-Path $sqlite_interop)) {Write-Host "$sqlite_interop found." -ForegroundColor Red -BackgroundColor Black}
If (-not(Test-Path $sqlite_core)) {Write-Host "$sqlite_core not found." -ForegroundColor Red -BackgroundColor Black}
If (-not(Test-Path $gudid_db_full_path)) {Write-Host "$gudid_db_full_path not found" -ForegroundColor Red -BackgroundColor Black}
$gudid_db_full_path = $gudid_db_full_path -replace "\\", "/"


## --------- Intro -------------------
Write-Host "*************************************************************************************" -ForegroundColor White -BackgroundColor Green
Get-Date
Set-Location $PSScriptRoot
Write-Host ""


## --------- Start -------------------
$timer = [System.Diagnostics.Stopwatch]::StartNew()
If ( (-not(Test-Path $sqlite3)) -OR (-not(Test-Path $gudid_db_full_path)) -OR (-not(Test-Path $sqlite_core)) -OR (-not(Test-Path $sqlite_interop)) ) {
Write-Host "This tool will perform the following..." -ForegroundColor White -BackgroundColor Black
Write-Host "- Download, unzip and extract required files from:" -ForegroundColor Red -BackgroundColor Black
If (-not(Test-Path $sqlite3)) {Write-Host $sqlite_uri -ForegroundColor White -BackgroundColor Black}
If ((-not(Test-Path $sqlite_core)) -OR (-not(Test-Path $sqlite_interop)) ) {Write-Host $sqlite_core_uri -ForegroundColor White -BackgroundColor Black}
If (-not(Test-Path $gudid_db_full_path)) {
Write-Host $gudid_uri -ForegroundColor White -BackgroundColor Black
Write-Host "-Build Access_GUDID Database File." -ForegroundColor Red -BackgroundColor Black
}
Write-Host "-Verify table count from Access_GUDID DB." -ForegroundColor Red -BackgroundColor Black
Write-Host "Press Ctrl+C to cancel." -ForegroundColor White -BackgroundColor Red
Write-Host "Press Enter to proceed." -ForegroundColor White -BackgroundColor Green
PAUSE
$timer.Stop()
Write-Host "User Response Received (HH:MM:SS.ms) - " $timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
$timer.Start()

}
else {
Write-Host "Required $sqlite3 found."
Write-Host "Required $sqlite_interop found."
Write-Host "Required $sqlite_core found."
$display_string = $gudid_db_full_path -replace "/", "\"
Write-Host "Required $display_string found."
}

## --------- Download sqlite -------------------
If (-not (Test-Path $sqlite3) ) {
Write-Host "**************" -ForegroundColor White -BackgroundColor Green
Get-Date
$sqlite_download_timer = [System.Diagnostics.Stopwatch]::StartNew()

$sqlite_uri = $sqlite_uri.Trim()
$sqlite_zip = Join-Path -Path $PSScriptRoot -ChildPath "sqlite.zip"
$sqlite_dir = Join-Path -Path $PSScriptRoot -ChildPath "sqlite"
$sqlite_exe = Join-Path -Path $sqlite_dir -ChildPath "sqlite-tools-win32-x86-3360000\sqlite3.exe"
Write-Host "Downloading file: $sqlite_uri." -ForegroundColor White -BackgroundColor Black
Write-Host "Destination: " $sqlite_zip -ForegroundColor Black -BackgroundColor White
Write-Host "Please wait..."
    try
    {   
    $Response = Invoke-WebRequest -Uri $sqlite_uri -OutFile $sqlite_zip 
    $StatusCode = $Response.StatusCode
    Write-Host "Download Successful: $sqlite_uri"  -ForegroundColor White -BackgroundColor Green


        If (-not (Test-Path $sqlite_zip) ) {
        Write-Host "$sqlite_zip <---- Required File not found." -ForegroundColor White -BackgroundColor Red 
        PAUSE
        BREAK
        }
        else {  
        If (-not (Test-Path $sqlite_dir)) {New-Item -ItemType directory $sqlite_dir | Out-Null}
        Expand-Archive -Path $sqlite_zip -DestinationPath $sqlite_dir -Force
        copy-item $sqlite_exe $PSScriptRoot -Force
        remove-item $sqlite_zip -recurse
        remove-item $sqlite_dir -recurse
        }      
    }
    catch
    {
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Failed to download:  $sqlite_uri" -ForegroundColor White -BackgroundColor Red
    break
    }
    $StatusCode
$sqlite_download_timer.Stop()
Get-Date
Write-Host "Download Time (HH:MM:SS.ms) - " $sqlite_download_timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
Write-Host "**************" -ForegroundColor White -BackgroundColor Green
}

## --------- Download DLL -------------------
If ((-not(Test-Path $sqlite_core)) -OR (-not(Test-Path $sqlite_interop)) ) {
$file = "system.data.sqlite.core.$version"
$sqlite_core_dll = $file + "/lib/netstandard2.0/System.Data.SQLite.dll"
$sqlite_interop_dll = $file + "/runtimes/win-x64/native/netstandard2.0/SQLite.Interop.dll"
$temp_download_dir =  Join-Path -Path $PSScriptRoot -ChildPath "temp_download"

Get-Date
$sqlite_dll_download_timer = [System.Diagnostics.Stopwatch]::StartNew()


If (-not (Test-Path $temp_download_dir) ) {New-Item -ItemType directory $temp_download_dir | Out-Null}
Set-Location $temp_download_dir

$dl = @{
	uri = $sqlite_core_uri
	outfile = "$file.zip"
}

Write-Host "Downloading file: $sqlite_core_uri." -ForegroundColor White -BackgroundColor Black
Write-Host "Destination: $temp_download_dir" -ForegroundColor Black -BackgroundColor White
try
{
    $Response = Invoke-WebRequest @dl 
    $StatusCode = $Response.StatusCode
    Write-Host "Download Successful: $sqlite_core_uri " -ForegroundColor White -BackgroundColor Green 
}
catch
{
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Failed to download:  $sqlite_core_uri " -ForegroundColor White -BackgroundColor Red
    break
}
$StatusCode

If (-not (Test-Path "$file.zip") ) {
Write-Host "$file zip <---- Required File not found." -ForegroundColor White -BackgroundColor Red 
PAUSE
BREAK
} else {
If (Test-Path "$file.zip") {
Expand-Archive $dl.outfile -Force
copy-item $sqlite_core_dll $PSScriptRoot -Force
copy-item $sqlite_interop_dll $PSScriptRoot -Force
               }
    }
Set-Location $PSScriptRoot
If (Test-Path $file) {remove-item $file -recurse}
If (Test-Path $temp_download_dir) {remove-item $temp_download_dir -recurse}

$sqlite_dll_download_timer.Stop()
Get-Date
Write-Host "Download Time (HH:MM:SS.ms) - " $sqlite_dll_download_timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
Write-Host "**************" -ForegroundColor White -BackgroundColor Green
}


## --------- Download GUDID -------------------
If (-not (Test-Path $gudid_db_full_path) ) {
$gudid_download_timer = [System.Diagnostics.Stopwatch]::StartNew()
Get-Date
$gudid_file_zip_path = Join-Path -Path $PSScriptRoot -ChildPath $gudid_file_zip
$gudid_file_dir = Join-Path -Path $PSScriptRoot -ChildPath $gudid_full_date
$data_file_name = "AccessGUDID_Delimited_Full_Release_" +  (Get-Date -Format "yyyyMM") + "01.data"
$gudid_db = Join-Path -Path $PSScriptRoot -ChildPath $data_file_name
$gudid_file_dir_unix_path = $gudid_file_dir  -replace "\\", "/"
Write-Host "Downloading file: $gudid_uri." -ForegroundColor White -BackgroundColor Black
Write-Host "Destination: " $gudid_file_zip_path -ForegroundColor Black -BackgroundColor White
Write-Host "Please wait..."

    try
    {
    $Response = Invoke-WebRequest -Uri $gudid_uri -OutFile $gudid_file_zip_path 
    $StatusCode = $Response.StatusCode
    Write-Host "Download Succesful: $gudid_uri " -ForegroundColor White -BackgroundColor Green  
	$gudid_download_timer.Stop()
	Get-Date
        Write-Host "Download Time (HH:MM:SS.ms) - " $gudid_download_timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
	Write-Host "**************" -ForegroundColor White -BackgroundColor Green
	Remove-Variable gudid_download_timer

        If (-not (Test-Path $gudid_file_zip_path) ) {
            Write-Host "$gudid_file_zip_path <---- Required Date File not found." -ForegroundColor White -BackgroundColor Red 
            PAUSE
            BREAK
            }
        else {
## --------- Build GUDID.db -------------------
            If (-not (Test-Path $gudid_file_dir)) {New-Item -ItemType directory $gudid_file_dir | Out-Null}
            Expand-Archive -Path $gudid_file_zip_path -DestinationPath $gudid_file_dir -Force
            remove-item $gudid_file_zip_path -recurse


            Set-Location $PSScriptRoot
            $sqlite_exe = Join-Path -Path $PSScriptRoot -ChildPath "sqlite3.exe"
            $gudid_db = Join-Path -Path $PSScriptRoot -ChildPath "gudid.db"

            If (-not (Test-Path $sqlite3) ){Write-Host "$sqlite3 <---- Exe File not found." -ForegroundColor White -BackgroundColor Red;BREAK}
 

## --------- GUDID DB FULL Build Parameters -------------------

$gudid_build_timer = [System.Diagnostics.Stopwatch]::StartNew()

$full_build_parameters = @"
.mode csv
.separator |
.import $gudid_file_dir_unix_path/contacts.txt  contacts
.import $gudid_file_dir_unix_path/device.txt  device
.import $gudid_file_dir_unix_path/deviceSizes.txt  deviceSizes
.import $gudid_file_dir_unix_path/environmentalConditions.txt  environmentalConditions
.import $gudid_file_dir_unix_path/gmdnTerms.txt  gmdnTerms
.import $gudid_file_dir_unix_path/identifiers.txt  identifiers
.import $gudid_file_dir_unix_path/productCodes.txt  productCodes
.import $gudid_file_dir_unix_path/sterilizationMethodTypes.txt  sterilizationMethodTypes
.import $gudid_file_dir_unix_path/premarketSubmissions.txt premarketSubmissions
CREATE INDEX di on device(PrimaryDI);
CREATE INDEX cn on device(catalogNumber);
CREATE INDEX mn on device(versionModelNumber);
CREATE INDEX idi on identifiers(PrimaryDI);
CREATE INDEX pdi on productCodes(PrimaryDI);
CREATE INDEX gdi on gmdnTerms(PrimaryDI);
CREATE UNIQUE INDEX upc on productCodes(PrimaryDI,productCode);
CREATE VIEW v_fda as select productCode, Lower(productCodeName), count(*) from productCodes group by productCode order by Lower(productCodeName);
CREATE VIEW v_gmdn as select gmdnPTName, count(*) from gmdnTerms group by gmdnPTName order by gmdnPTName;
.tables
"@

Get-Date
Write-host $build_parameters -ForegroundColor White -BackgroundColor Blue
Write-host "GUDID Database Build will take 3-5 minutes."
Write-Host "$gudid_db_full_path DB Build in progress. Please wait....."
#$build_parameters | .\sqlite3 $gudid_db_full_path
$full_build_parameters | .\sqlite3 $gudid_db_full_path

$gudid_build_timer.Stop()
Get-Date
Write-Host "Build Time (HH:MM:SS.ms) - " $gudid_build_timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
Write-Host "**************" -ForegroundColor White -BackgroundColor Green

If (-not (Test-Path $gudid_db_full_path) ){Write-Host "$gudid_db_full_path <---- File not found." -ForegroundColor White -BackgroundColor Red;BREAK}   
Write-Host "$gudid_db_full_path build completed." -ForegroundColor White -BackgroundColor Blue
remove-item $gudid_file_dir -recurse


            }
    }
##--------------------End of Build ----------------------------------------
    catch
    {
    $StatusCode = $_.Exception.Response.StatusCode.value__
    Write-Host "Failed to download: " $gudid_uri
    break
    }
    $StatusCode
}


$timer.Stop()
Get-Date
Write-Host "Total Elapsed Time (HH:MM:SS.ms) - " $timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
Write-Host "**************" -ForegroundColor White -BackgroundColor Green
Remove-Variable timer

Get-Date
$timer = [System.Diagnostics.Stopwatch]::StartNew()

Write-Host ""
Write-Host "Verifying row count."  -ForegroundColor White -BackgroundColor Blue
Set-Location $PSScriptRoot

$timestamp = Get-Date -Format "dddd_yyyy_MM_dd_HH_mm_ss"
$output = "row_count_" + $timestamp + ".sql"
$output = Join-Path -Path $PSScriptRoot -ChildPath $output
$output = $output -replace "\\", "/"

$query_parameters = @"
.output $output
WITH RECURSIVE
tbl(name) AS (Select name FROM sqlite_master WHERE type IN ("table"))
SELECT  ' select " No. of rows from ' || name || ' table -  ", printf ("%,d",count (*))  from ' || name || ';' FROM tbl;
.output
.mode tab
.header off
.echo off
.read $output
"@

$query_parameters | .\sqlite3 $gudid_db_full_path
Remove-Item $output


$timer.Stop()
Get-Date
Write-Host "Verification Time (HH:MM:SS.ms) - " $timer.Elapsed -ForegroundColor White -BackgroundColor Magenta
Write-Host "**************" -ForegroundColor White -BackgroundColor Green
Remove-Variable timer 

Add-Type -Path $sqlite_core
$con = New-Object -TypeName System.Data.SQLite.SQLiteConnection
$con.ConnectionString = "Data Source=$gudid_db_full_path"
$con.Open()
$sql = $con.CreateCommand()

$timestamp = Get-Date -Format "dddd_yyyy_MM_dd_HH_mm_ss"
$query_result_csv = "GUDID_DB_Query_Result_" + $timestamp + ".csv"
$query_result_xl = "GUDID_DB_Query_Result_" + $timestamp + ".xlsx"
$query_result_csv = Join-Path -Path $PSScriptRoot -ChildPath $query_result_csv
$query_result_xl = Join-Path -Path $PSScriptRoot -ChildPath $query_result_xl


Write-Host "$gudid_db_full_path is ready to receive your query."  -ForegroundColor White -BackgroundColor Green


## ---------- Prompt to Query Loop ------------------------------------------
$session_timer = [System.Diagnostics.Stopwatch]::StartNew()

$searchString = ""
Do
{ 
Write-host ""
Get-Date
Write-host "Enter device to search (.e.g stent)" -ForegroundColor White -BackgroundColor Magenta
	
$searchString = Read-Host 
    
if ((-not ($searchString)) -OR ($searchString -eq "exit")  -OR ($searchString.Length -lt 4)) {[console]::beep(1000,100); BREAK}
$searchString = $searchString -replace " ", "%"

$timer = [System.Diagnostics.Stopwatch]::StartNew()

$sql.CommandText = @"
Select companyName, PrimaryDI, catalogNumber, versionModelNumber, brandName, deviceDescription
FROM device where PrimaryDI like '%$searchString%'
or deviceDescription like '%$searchString%'
or brandName like '%$searchString%'
or companyName like '%$searchString%'
or catalogNumber like '%$searchString%'
or versionModelNumber like '%$searchString%' 	
"@

    
    Write-Host "SQLite query:"
    Write-Host $sql.CommandText 

	$adapter = New-Object -TypeName System.Data.SQLite.SQLiteDataAdapter $sql
	$data = New-Object System.Data.DataSet
	[void]$adapter.Fill($data)
	
	$timer.Stop()
    Write-Host "Query Time (HH:MM:SS.ms) - " $timer.Elapsed -ForegroundColor White -BackgroundColor Magenta

	If ($data.Tables.Rows) {
        [console]::beep(400,500)
        Write-Host "Displaying result. Please wait...."
        $data.tables.Rows | Out-Gridview -Title "$searchString"
        Write-Host $data.tables.Rows.Count " - records retrieved."  
        $data.tables.Rows | Export-Csv -Path $query_result_csv -Append
        Write-Host "Result saved to $query_result_csv"
		Write-host "Press ENTER to end search." -ForegroundColor Black -BackgroundColor Yellow
		
  
	}
	Else {
        [console]::beep(200,1000)
		Write-Host "NO entry found for:  $searchString" -ForegroundColor Red -BackgroundColor Black
	}


} 
While ($searchString)
## ---------- Prompt to Query Loop ------------------------------------------

$sql.Dispose()
$con.Close()

	$session_timer.Stop()
    Write-Host "Session Time (HH:MM:SS.ms) - " $session_timer.Elapsed -ForegroundColor White -BackgroundColor Red
    Get-Date
	Write-Host "**************" -ForegroundColor White -BackgroundColor Green

If (Test-Path $query_result_csv){

Write-Host "Result file: $query_result_csv found." -ForegroundColor White -BackgroundColor Magenta

$inputCSV = $query_result_csv
$outputXLSX = $query_result_xl
$excel = New-Object -ComObject excel.application 
$workbook = $excel.Workbooks.Add(1)
$worksheet = $workbook.worksheets.Item(1)
$TxtConnector = ("TEXT;" + $inputCSV)
$Connector = $worksheet.QueryTables.add($TxtConnector,$worksheet.Range("A1"))
$query = $worksheet.QueryTables.item($Connector.name)
$query.TextFileOtherDelimiter = $Excel.Application.International(5)
$query.TextFileParseType  = 1
$query.TextFileColumnDataTypes = ,2 * $worksheet.Cells.Columns.Count
$query.AdjustColumnWidth = 1
$query.Refresh() | Out-Null
$query.Delete() 
$workbook.SaveAs($outputXLSX,51) 

   If (-not(Test-Path $query_result_xl)) {
          $excel.Quit()     
          Write-Host "Result file: $query_result_csv can be imported in Excel." -ForegroundColor White -BackgroundColor Magenta      
          }  
    else  {
            Remove-Item $query_result_csv
            Write-Host "Prepping $query_result_xl. Please wait..." -ForegroundColor White -BackgroundColor Magenta	
            $worksheet.Cells.Item(1,1).EntireRow.Delete() | Out-Null
            $UsedRange = $worksheet.UsedRange
            $RowCount = $UsedRange.Rows.count
            $RowCount = $RowCount - 1
            $RowCount = '{0:N0}' -f $RowCount
            $ColCount = $UsedRange.Columns.count
            Write-Host "Total Columns: " $ColCount -ForegroundColor White -BackgroundColor Green
            Write-Host "Total Rows Imported - " $RowCount -ForegroundColor White -BackgroundColor Green

            Write-Host "Setting split pane."
            $workbook.Application.ActiveWindow.SplitColumn = 1
            $workbook.Application.ActiveWindow.SplitRow = 1
            $workbook.Application.ActiveWindow.FreezePanes = $true

            Write-Host "Setting font and highlight."
            for($i = 1; $i -lt $ColCount+1; $i++){ 
            $worksheet.Cells.Item(1,$i).Font.Bold = $True
            $worksheet.Cells.Item(1,$i).Interior.ColorIndex = 6
            $worksheet.Cells.Item(1,$i).columnwidth = 30  
            }
 
            Write-Host "Displaying spreadsheet."
            $workbook.Save()
            $excel.visible = $true
                             
          }  
   
}

Get-Variable -Exclude PWD,*Preference | Remove-Variable -EA 0