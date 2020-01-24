 #############################################################################
# SCRIPT - POWERSHELL
# NAME: Generate Index Workbooks - v1.0.1.ps1
# 
# DATE:  01/24/2020
# 
# COMMENT:  This script will collect index information from all the tables in the specified database
#           and output them to an Excel workbook / worksheets for analysis
#
#
#           This uses a cmdlet from the ImportExcel module
#           https://github.com/dfinke/ImportExcel
#			This also uses sp_SQLSkills_helpindex - 
#
# VERSION HISTORY
# 1.0.1 - Fixes issues with tables with a single index
# 
#
# TO ADD
# -
# -
# #############################################################################

# Template path
$templatePath = "C:\Scripts\Index Analysis Generation Template.xlsx"
# Server to connect to
$serverName = "SERVERNAME\INSTANCENAME"
# Database to connect to
$databaseName = "DATABASENAME"
# Output file path
$strPath = "C:\Scripts\Index Analysis - " + $databaseName + ".xlsx"


# SQL command to get a list of all table names in the database
$tablenames = "SELECT TABLE_NAME FROM " + $databaseName + ".INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = `'BASE TABLE`' AND TABLE_NAME <> `'dtproperties`' ORDER BY TABLE_NAME"

# Create a new Excel workbook
$ExcelObject = New-Object -ComObject Excel.Application
$ExcelObject.visible = $false
$ExcelObject.DisplayAlerts = $false


# if the Excel workbook exists already, delete it.
if (Test-Path $strPath) 
{
  Remove-Item $strPath
}

# create a counter for selecting worksheets
$counter = 1

# Get a list of all the tables in the database
$Tables = Invoke-SqlCmd -ServerInstance $serverName  -Database $databaseName -Query $tablenames

# for each table in the database, create a new worksheet with the same name as the table name
foreach($table in $tables.TABLE_NAME){
Copy-ExcelWorkSheet -SourceWorkbook $templatePath -SourceWorkSheet Template -DestinationWorkSheet $table -DestinationWorkbook $strPath
}


#Open the document, using the COM object model method 
$ActiveWorkbook = $ExcelObject.WorkBooks.Open($strPath)

# for each worksheet, activate it, pull a list of the index information, and paste it appropriately onto the right sheet
foreach($table in $tables.TABLE_NAME){
$ActiveWorksheet = $ActiveWorkbook.Worksheets.Item($counter)
$ActiveWorksheet.Activate()

$index_name = "DECLARE @Tmp 
TABLE (		
		index_id			int,
		is_disabled         bit,
		index_name			nvarchar(MAX),
		index_description   nvarchar(MAX),
		index_keys			nvarchar(MAX),
		inc_columns			nvarchar(max),
		filter_definition	nvarchar(max),
		columns_in_tree		nvarchar(MAX),
		columns_in_leaf		nvarchar(max))

INSERT @Tmp EXEC sp_SQLSkills_helpindex " + $table + "; " + 

"SELECT  index_name, index_description, index_keys, inc_columns from @Tmp;"

$Result = Invoke-SqlCmd -ServerInstance $serverName -Database $databaseName -Query $index_name

$counter2 = 1

# if there are no indexes, $Result will be $null, so we need to check to make sure it isn't null
# if $Result is not null, and $Result.Rows.Count -eq 0, then that means we have a single index on the table
if($Result -and ($Result.Rows.Count -eq 0))
# if there is only one index, the results will be returned in a single dimension array
{
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 1) = $Result[0] 
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 2) = $Result[1]
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 3) = $Result[2]
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 4) = $Result[3]
        $counter2 = $counter2 + 1
}


# if there is more than one index, the results will be a multi-dimensional array
if ($Result.Rows.Count -gt 1)
{
    foreach($Results in $Result)
    {
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 1) = $Result[$counter2 - 1].index_name 
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 2) = $Result[$counter2 - 1].index_description
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 3) = $Result[$counter2 - 1].index_keys
        $ActiveWorkSheet.Cells.Item($counter2 + 1, 4) = $Result[$counter2 - 1].inc_columns
        $counter2 = $counter2 + 1
    }
}

$ActiveWorkSheet.Cells.Item($counter2 + 2, 1) = "Proposed Changes:"
$ActiveWorkSheet.Cells.Item($counter2 + 2, 1).Font.Bold = $True
$ActiveWorksheet.Columns.Autofit()

$counter = $counter + 1
}

# Save workbook, close workbook, and quit Excel for good hygiene
$ActiveWorkbook.Save()
$ActiveWorkbook.Close()
$ExcelObject.quit()

