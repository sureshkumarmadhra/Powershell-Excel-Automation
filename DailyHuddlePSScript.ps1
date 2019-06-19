
#$parentFolderPath = "C:\Users\smadhra\DXC Production\UKI NCR Hub - Huddle Reports\Templates\"


$parentFolderPath ="C:\Users\svindhyachal.000\Desktop\Templates\"

#$parentFolderPath ="C:\UKHuddle\Templates\"


$checkFileStatus = $true

#$parentFolderPath = "C:\UKI NCR Hub - Huddle Reports\Templates\"

$LogFileName = "Script_"

# Create Log File Name for Current D
$currentdate = Get-Date -format "dd_MMM_yyyy"
$CurrentLogFileName = $parentFolderPath + $LogFileName + $currentdate + ".log"

# Delete Last Date Log File
$lastdate = (Get-Date).AddDays(-1).ToString('dd_MMM_yyyy')
$lastDateLogFileName = $parentFolderPath + $LogFileName + $lastdate + ".log"

if (Test-Path -path $lastDateLogFileName){
Remove-Item $lastDateLogFileName
}


if (Test-Path -path $CurrentLogFileName){
Remove-Item $CurrentLogFileName
}


#Write Log Function

function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, 
                   ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path=   $CurrentLogFileName, 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
         
        # If the file already exists and NoClobber was specified, do not write to the log. 
        if ((Test-Path $Path) -AND $NoClobber) { 
            Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
            Return 
            } 
 
        # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
        elseif (!(Test-Path $Path)) { 
            Write-Verbose "Creating [$Path] File" 
            $NewLogFile = New-Item $Path -Force -ItemType File 
            } 
 
        else { 
            # Nothing to see here yet. 
            } 
 
        # Format Date for our Log File 
        $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
        # Write message to error, warning, or verbose pipeline and specify $LevelText 
        switch ($Level) { 
            'Error' { 
                Write-Error $Message 
                $LevelText = 'ERROR:' 
                } 
            'Warn' { 
                Write-Warning $Message 
                $LevelText = 'WARNING:' 
                } 
            'Info' { 
                Write-Verbose $Message 
                $LevelText = 'INFO:' 
                } 
            } 
         
        # Write log entry to $Path 
        "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
    } 
    End 
    { 
    } 
}





write-Log -Message ("Location of UKI NCR HUB Report Parent Folder -->" + $parentFolderPath) -Level Info
Write-Log -Message ("Location of Script Execution Log File -->" + $CurrentLogFileName) -Level Info

$RawDataParentFolderPath = $parentFolderPath + "RAW Data\"


Write-Log -Message ("Location of Raw Data File's Location -->" + $RawDataParentFolderPath) -Level Info

$RawDataParentFileList = get-childitem $RawDataParentFolderPath 
$TemplateFileList = get-childitem $parentFolderPath 

$NWROutStandingFileDir = $RawDataParentFileList | Where-Object { $_.Name -like "*NWR*Out*" }
$NWROutStandingFileDir

if ($NWROutStandingFileDir -ne ""){
    foreach ($file in $NWROutStandingFileDir){
    $nwrOutstandingReportFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("Network Rail Daily Raw Outstanding Incident Data Excel File [NWR Outstanding.csv] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}

$NWRResolvedFileDir = $RawDataParentFileList | Where-Object { $_.Name -like "*NWR*Reso*" }

if ($NWRResolvedFileDir -ne ""){
    foreach ($file in $NWRResolvedFileDir){
    $nwrResolvedReportFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("Network Rail Daily Raw Resolved Incident Data Excel File [NWR Resolved.csv] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}


$NWRTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*NWR*Templ*" }

if ($NWRTemplateParentDir -ne ""){
    foreach ($file in $NWRTemplateParentDir){
    $nwrTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("Network Rail Template Excel File [NWR Template.xls] is unavailable at Location -->" + $parentFolderPath) -Level Info
}


if ($nwrOutstandingReportFileName -eq ""){

Write-output("NWR Outstanding File Missing, Please check")

}

if ($NWRResolvedFileDir -eq ""){

Write-output("NWR Resolved File Missing, Please check")

}


$CPARemdyIncidentParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*CPA Global MTD Metrics*" }

if ($CPARemdyIncidentParentDir -ne ""){
    foreach ($file in $CPARemdyIncidentParentDir){
    $cpaRemedyDailyReportFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("CPA Daily Remedy Raw Incident Data Excel File [CPA Global MTD Metrics.xls] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}



$CPARemdyTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*CPA*Templ*" }

if ($CPARemdyTemplateParentDir -ne ""){
    foreach ($file in $CPARemdyTemplateParentDir){
    
    $cpaRemedyTemplateFileName = $file.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("CPA Remedy Template Excel File [CPA Template.xls] is unavailable at Location -->" + $parentFolderPath) -Level Info
}



$CPAJIRARawParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*DXC*" }

if ($CPAJIRARawParentDir -ne ""){
    foreach ($file in $CPAJIRARawParentDir){
    $cpaJIRADailyReportFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("CPA Daily Raw Excel JIRA File [DXC Current Month Metric*.csv] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}


$CPAJIRATemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*Jira*" }
$CPAJIRATemplateParentDir
if ($CPAJIRATemplateParentDir -ne ""){
    foreach ($file in $CPAJIRATemplateParentDir){
    $cpaJIRATemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("CPA JIRA Template Excel File [CPA JIRA template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}



$XchangingIncidentRawParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*xChanging*" }

if ($XchangingIncidentRawParentDir -ne ""){
    foreach ($file in $XchangingIncidentRawParentDir){
    $xchangingReportFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("Xchanging Daily Raw Incident Data Excel File [xChanging Tickets*.xlsx] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}


$XchangingIncidentTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*Xchanging*Templ*" }

if ($XchangingIncidentTemplateParentDir -ne ""){
    foreach ($file in $XchangingIncidentTemplateParentDir){
    $xchangingTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("XChanging Template Excel File [Xchanging Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}



$qbeIncidentRawParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*inc*sla*" }

if ($qbeIncidentRawParentDir -ne ""){
    foreach ($file in $qbeIncidentRawParentDir){
    $qbeIncidentDailyReportFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("QBE Daily Raw Incident Data Excel File [incident_sla.xlsx] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}

$qbeSRRawParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*task*" }

if ($qbeSRRawParentDir -ne ""){
    foreach ($file in $qbeSRRawParentDir){
    $qbeSRDailyReportFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("QBE Daily Raw SR Data Excel File [task.xlsx] is unavailable at Location -->" + $RawDataParentFolderPath) -Level Info
}

$qbeIncidentTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*QBE Template*" }

if ($qbeIncidentTemplateParentDir -ne ""){
    foreach ($file in $qbeIncidentTemplateParentDir){
    $qbeTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("QBE Incident Template Excel File [QBE Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}


$qbeSRTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*QBE_SR*" }

if ($qbeSRTemplateParentDir -ne ""){
    foreach ($file in $qbeSRTemplateParentDir){
    $qbeSRTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("QBE SR Template Excel File [QBE_SR Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}
$ukiNCRHUBTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "NCR_UKI HUB Huddle_Master 1.xlsm" }

if ($ukiNCRHUBTemplateParentDir -ne ""){
    foreach ($file in $ukiNCRHUBTemplateParentDir){
    $ukNCRHUBExcelFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("UKNCR HUB HUDDLE Excel File [NCR_UKI HUB Huddle.xlsm] is unavailable at Location -->" + $parentFolderPath) -Level Info
}




$belronTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*Belron Template*" }

if ($belronTemplateParentDir -ne ""){
    foreach ($file in $belronTemplateParentDir){
    $belronTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("Belron Incident Template Excel File [Belron Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}


$belronResolvedParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*Belron_Resolved*" }

if ($belronResolvedParentDir -ne ""){
    foreach ($file in $belronResolvedParentDir){
    $belronDailyIncidentResolvedFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("Belron Resolved Incident Excel File [Beloron_Resolved.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}




$belronOutStandingParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*Belron_Outstanding*" }

if ($belronOutStandingParentDir -ne ""){
    foreach ($file in $belronOutStandingParentDir){
    $belronDailyIncidentOutStandingFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("Belron OutStanding Incident Excel File [Belron_Outstanding.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}





$ExovaTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*Exova Template.xls*" }

if ($ExovaTemplateParentDir -ne ""){
    foreach ($file in $ExovaTemplateParentDir){
    $ExovaTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("Exova Incident Template Excel File [Exova Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}


 $ExovaResolvedParentDir_ = $RawDataParentFileList | Where-Object { $_.Name -like "*Exova_Resolved.xls*" }
$ExovaResolvedParentDir_
if ($ExovaResolvedParentDir_ -ne ""){
    foreach ($file in $ExovaResolvedParentDir_){
    $ExovaDailyIncidentResolvedFileName = $File.Name
    $ExovaDailyIncidentResolvedFileName
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("Exova Resolved Incident Excel File [Exova_Resolved.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}
$ExovaDailyIncidentResolvedFileName



$ExovaOutStandingParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*Exova_Outstanding.xls*" }

if ($ExovaOutStandingParentDir -ne ""){
    foreach ($file in $ExovaOutStandingParentDir){
    $ExovaDailyIncidentOutStandingFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("Exova OutStanding Incident Excel File [Exova_Outstanding.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}



$ngridTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*Nationa Grid Template_US*" }

if ($ngridTemplateParentDir -ne ""){
    foreach ($file in $ngridTemplateParentDir){
    $ngridTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("NATIONAL GRID US Incident Template Excel File [ngrid Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}


 $ngridResolvedParentDir_ = $RawDataParentFileList | Where-Object { $_.Name -like "*NG_US_Remedy_Dump_Resolved.csv*" }

if ($ngridResolvedParentDir_ -ne ""){
    foreach ($file in $ngridResolvedParentDir_){
    $ngridDailyIncidentResolvedFileName = $File.Name
    $ngridDailyIncidentResolvedFileName
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("NATIONAL GRID US Resolved Incident Excel File [NG_US_Remedy_Dump_Resolved.csv] is unavailable at Location -->" + $parentFolderPath) -Level Info
}
$ngridDailyIncidentResolvedFileName



$ngridOutStandingParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*NG_US_Remedy_Dump_Outstanding*" }

if ($ngridOutStandingParentDir -ne ""){
    foreach ($file in $ngridOutStandingParentDir){
    $ngridDailyIncidentOutStandingFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("NATIONAL GRID US  OutStanding Incident Excel File [NG_US_Remedy_Dump_Outstanding.csv] is unavailable at Location -->" + $parentFolderPath) -Level Info
}

$ngridukTemplateParentDir = $TemplateFileList | Where-Object { $_.Name -like "*Nationa Grid Template_UK*" }

if ($ngridukTemplateParentDir -ne ""){
    foreach ($file in $ngridukTemplateParentDir){
    $ngridukTemplateFileName = $File.Name
    }
}
else
{
$checkFileStatus=$false
Write-Log -Message ("NATIONAL GRID UK Incident Template Excel File [Nationa Grid Template.xlsx] is unavailable at Location -->" + $parentFolderPath) -Level Info
}


 $ngridukResolvedParentDir_ = $RawDataParentFileList | Where-Object { $_.Name -like "*NG_UK_Remedy_Dump_Resolved.csv*" }

if ($ngridukResolvedParentDir_ -ne ""){
    foreach ($file in $ngridukResolvedParentDir_){
    $ngridukDailyIncidentResolvedFileName = $File.Name
    $ngridukDailyIncidentResolvedFileName
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("NATIONAL GRID UK Resolved Incident Excel File [NG_UK_Remedy_Dump_Resolved.csv] is unavailable at Location -->" + $parentFolderPath) -Level Info
}
$ngridukDailyIncidentResolvedFileName



$ngridukOutStandingParentDir = $RawDataParentFileList | Where-Object { $_.Name -like "*NG_UK_Remedy_Dump_Outstanding*" }

if ($ngridukOutStandingParentDir -ne ""){
    foreach ($file in $ngridukOutStandingParentDir){
    $ngridukDailyIncidentOutStandingFileName = $File.Name
    }
}

else
{
$checkFileStatus=$false
Write-Log -Message ("NATIONAL GRID UK OutStanding Incident Excel File [NG_UK_Remedy_Dump_Outstanding.csv] is unavailable at Location -->" + $parentFolderPath) -Level Info
}




if ($checkFileStatus -eq $false){
Write-Log -Message ("Unable to Run the Script due to unavailable files") -Level Info
}

else{

 

Write-Log -Message ("Going to Start the Script Execution") -Level Info

Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING CPA DAILY INCIDENT REMEDY RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read CPA Daily Incident Remedy  Report Data") -Level Info

#Declare File Name, Sheet Name for CPA Daily Incident Remedy Reports

$cpaRemedytemplateFilePath =  $parentFolderPath + $cpaRemedyTemplateFileName
$cpaRemedyDailyReportFilePath =  $RawDataParentFolderPath + $cpaRemedyDailyReportFileName
$cpaRemedyTemplateRawDataWorkSheetName = "CPA Raw data"
$cpaRemedyTemplateHuddleSheetName = "CPA Data 4 Huddle"
$TempworksheetName = "Temp"

Write-Log -Message ("CPA REMEDY Ticket File Path --> " + $cpaRemedyDailyReportFilePath) -Level Info
Write-Log -Message ("CPA REMEDY Template File --> " + $cpaRemedytemplateFilePath) -Level Info

#Declare Excel Object for CPA Remedy Incident Data
Write-Log -Message ("Going to Create Excel Object for CPA Daily Incident Remedy  Workbook") -Level Info
$cpaRemedyRAWdataExcelFileObject = New-Object -ComObject excel.application
$cpaRemedyRAWdataExcelFileObject.Visible = $true
$cpaRemedyRAWdataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening CPA Daily Incident Remedy File : " + $cpaRemedyDailyReportFilePath) -Level Info
$cpaRemedyRAWdataExcelWorkbook = $cpaRemedyRAWdataExcelFileObject.Workbooks.Open($cpaRemedyDailyReportFilePath)
$cpaRemedyRAWdataExcelWorkbook.activate()
Write-Log -Message ("Selecting CPA Daily Incident Remedy Worksheet : [" + $cpaRemedyIncidentRawDataWorkSheetName + "]") -Level Info
$cpaRemedyrawdataExcelWorksheet = $cpaRemedyRAWdataExcelWorkbook.Worksheets.Item(1)
$cpaRemedyIncidentRawDataWorkSheetName = $cpaRemedyrawdataExcelWorksheet.Name
$cpaRemedyrawdataExcelWorksheet.Activate()
$cpaRemedyrawdataExcelWorksheetRange = $cpaRemedyrawdataExcelWorksheet.Range("A:N").CurrentRegion
Write-Log -Message ("Copying Cells from A to T in Worksheet : [" + $cpaRemedyrawdataExcelWorksheet.Name + "] from " + $cpaRemedyDailyReportFilePath  + " File." ) -Level Info
$cpaRemedyrawdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to T in Worksheet : [" + $cpaRemedyrawdataExcelWorksheet.Name + "] from " + $cpaRemedyDailyReportFilePath  + " File.") -Level Info

# Copy the CPA Daily Remedy Raw Data to CPA Template File

Write-Log -Message ("Going to Open CPA Remedy Template File") -Level Info
$cpaRemedyTemplateExcelFileObject = New-Object -ComObject excel.application
$cpaRemedyTemplateExcelFileObject.Visible = $true
$cpaRemedyTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening CPA Remedy Template File --> " + $cpaRemedytemplateFilePath) -Level Info
$cpaRemedyTemplateExcelWorkbook = $cpaRemedyTemplateExcelFileObject.Workbooks.Open($cpaRemedytemplateFilePath)
$cpaRemedyTemplateExcelWorkbook.activate()



# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $cpaRemedyTemplateExcelWorkbook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $cpaRemedytemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



Write-Log -Message ("Selecting CPA Template Remedy Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $cpaRemedytemplateFilePath) -Level Info
$cpaRemedyTempExcelWorksheet = $cpaRemedyTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$cpaRemedyTempExcelWorksheet.Activate()
$cpaRemedyTempExcelWorksheetRange = $cpaRemedyTempExcelWorksheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $cpaRemedytemplateFilePath) -Level Info

$cpaRemedyTempExcelWorksheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $cpaRemedytemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $cpaRemedytemplateFilePath) -Level Info
$cpaRemedyTempExcelWorksheet = $cpaRemedyTemplateExcelWorkbook.Worksheets.Add()
$cpaRemedyTempExcelWorksheet.Name = $TempworksheetName
$cpaRemedyTempExcelWorksheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $cpaRemedytemplateFilePath) -Level Info

}



Write-Log -Message ("Selecting '" + $cpaRemedyTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $cpaRemedytemplateFilePath) -Level Info

$cpaTemplateIncidentRawExcelWorksheet = $cpaRemedyTemplateExcelWorkbook.Worksheets.Item($cpaRemedyTemplateRawDataWorkSheetName)
$cpaTemplateIncidentRawExcelWorksheet.Activate()
#Delete Old Records from Sheet "CPA RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To T in '" + $cpaRemedyTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $cpaRemedytemplateFilePath) -Level Info
$cpaTemplateIncidentRawExcelWorksheetRange = $cpaTemplateIncidentRawExcelWorksheet.Range("A:T")
$cpaTemplateIncidentRawExcelWorksheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To T in '" + $cpaRemedyTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $cpaRemedyTemplateRawDataWorkSheetName) -Level Info
$cpaTemplateIncidentRawExcelWorksheetRange = $cpaTemplateIncidentRawExcelWorksheetRange.Range("A1")
Write-Log -Message ("Copying CPA Remedy Incident Raw Data from Worksheet [" + $cpaRemedyIncidentRawDataWorkSheetName + "] of File --> " + $cpaRemedyDailyReportFilePath + " to Worksheet [" + $cpaRemedyTemplateRawDataWorkSheetName + " of File --> " +  $cpaRemedytemplateFilePath) -Level Info
$cpaTemplateIncidentRawExcelWorksheet.Paste($cpaTemplateIncidentRawExcelWorksheetRange)
Write-Log -Message ("Copied CPA Remedy Incident Raw Data from Worksheet [" + $cpaRemedyIncidentRawDataWorkSheetName + "] of File --> " + $cpaRemedyDailyReportFilePath + " to Worksheet [" + $cpaRemedyTemplateRawDataWorkSheetName + "] of File --> " +  $cpaRemedytemplateFilePath) -Level Info
#Save the CPA Jira Template File
#$cpaJIRATemplateExcelWorkbook.Save()
#Write-Log -Message ("Closing CPA JIRA Daily Report Data File --> " + $cpaJIRADailyReportFilePath) -Level Info
#$cpaJIRArawdataExcelFileObject.Quit()
#Write-Log -Message ("Closed CPA JIRA Daily Report Data File --> " + $cpaJIRADailyReportFilePath) -Level Info

#Apply Filter the CPA Daily Incident Data in "CPA Data 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of CPA Daily Remedy Incident Data in Worksheet [" + $cpaRemedyTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$cpaRemedyTemplateHurdleExcelWorksheet = $cpaRemedyTemplateExcelWorkbook.Worksheets.Item($cpaRemedyTemplateHuddleSheetName)
$cpaRemedyTemplateHurdleExcelWorksheet.Activate()
$cpaRemedyTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$cpaRemedyTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$cpaRemedyTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $cpaRemedyTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$cpaRemedyTempExcelWorksheet.Activate()

$cpaRemedyTempExcelWorksheet.Range("A1").PasteSpecial(-4163)
$cpaRemedyTempExcelWorksheetRange = $cpaRemedyTempExcelWorksheet.UsedRange

$cpaRemedyIncidentCount = $cpaRemedyTemplateExcelFileObject.WorksheetFunction.CountIf($cpaRemedyTempExcelWorksheetRange.Range("A1:" + "A" + $cpaRemedyTempExcelWorksheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of CPA Daily Remedy Incident Count is : [" + $cpaRemedyIncidentCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $cpaRemedytemplateFilePath) -Level Info
$cpaRemedyTemplateExcelWorkbook.Close($true)
#$cpaRemedyTemplateExcelWorkbook.Save()
$cpaRemedyTemplateExcelFileObject.Quit()
$cpaRemedyRAWdataExcelFileObject.Quit()
$cpaRemedyRAWdataExcelFileObject = $null
$cpaRemedyTemplateExcelFileObject = $null
$cpaRemedyTempExcelWorksheet = $null
$cpaRemedyTemplateHurdleExcelWorksheet = $null
$cpaRemedyRAWdataExcelWorkbook = $null
$cpaRemedyTemplateExcelWorkbook = $null

Write-Log -Message ("**************** PROCESSED CPA DAILY REMEDY INCIDENT RAW DATA *************************************") -Level Info




#Declare File Name, Sheet Name for CPA Daily JIRA Reports

Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING CPA DAILY JIRA TICKET RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read CPA Daily JIRA Ticket Report Data") -Level Info




#$cpaJIRATempRawDataWorkSheetName = "Current Month DXC Tickets (JIRA"
#$cpaJIRATempRawDataWorkSheetName ="DXC Current Month Metric (JIRA "
$cpaJIRATemplateRawDataWorkSheetName = "JIRA_RAW"
$cpaJIRATemplateHuddleSheetName = "CPA Jeera dump 4 Huddle"
$TempworksheetName = "Temp"
$cpaJIRAtemplateFilePath =  $parentFolderPath + $cpaJIRATemplateFileName
$cpaJIRADailyReportFilePath =  $RawDataParentFolderPath + $cpaJIRADailyReportFileName

Write-Log -Message ("CPA Daily JIRA Ticket Data Path --> " + $cpaJIRADailyReportFilePath) -Level Info
Write-Log -Message ("CPA JIRA Template File --> " + $cpaJIRAtemplateFilePath) -Level Info



#Declare Excel Object for CPA Remedy Incident Data
Write-Log -Message ("Going to Create Excel Object for CPA Daily JIRA Ticket  Workbook") -Level Info

$cpaJIRArawdataExcelFileObject = New-Object -ComObject excel.application
$cpaJIRArawdataExcelFileObject.Visible = $true
$cpaJIRArawdataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening CPA JIRA Daily Ticket Data" + $cpaJIRADailyReportFilePath) -Level Info
$cpaJIRArawdataExcelWorkbook = $cpaJIRArawdataExcelFileObject.Workbooks.Open($cpaJIRADailyReportFilePath)
$cpaJIRArawdataExcelWorkbook.activate()
Write-Log -Message ("Selecting CPA Daily JIRA Ticket Worksheet : [" + $cpaJIRATempRawDataWorkSheetName + "]") -Level Info

$cpaJIRArawdataExcelWorksheet = $cpaJIRArawdataExcelWorkbook.Worksheets.Item(1)
$cpaJIRATempRawDataWorkSheetName = $cpaJIRArawdataExcelWorksheet.Name

$cpaJIRArawdataExcelWorksheet.Activate()
$cpaJIRArawdataExcelWorksheetRange = $cpaJIRArawdataExcelWorksheet.Range("A:N").CurrentRegion
Write-Log -Message ("Copying Cells from A to N in Worksheet [" + $cpaJIRATempRawDataWorkSheetName + "] from " + $cpaJIRADailyReportFilePath  + " File." ) -Level Info
$cpaJIRArawdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to N in Worksheet [" + $cpaJIRATempRawDataWorkSheetName + "] from " + $cpaJIRADailyReportFilePath  + " File.") -Level Info

# Copy the CPA Daily Raw Data to CPA Template File

Write-Log -Message ("Going to Open CPA JIRA Template File") -Level Info
$cpaJIRATemplateExcelFileObject = New-Object -ComObject excel.application
$cpaJIRATemplateExcelFileObject.Visible = $true
$cpaJIRATemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening CPA JIRA Template File --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRATemplateExcelWorkbook = $cpaJIRATemplateExcelFileObject.Workbooks.Open($cpaJIRAtemplateFilePath)
$cpaJIRATemplateExcelWorkbook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $cpaJIRATemplateExcelWorkbook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $cpaJIRAtemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



Write-Log -Message ("Selecting CPA Template Remedy Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRATempdataExcelWorksheet = $cpaJIRATemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$cpaJIRATempdataExcelWorksheet.Activate()
$cpaJIRATempdataExcelWorksheetRange = $cpaJIRATempdataExcelWorksheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $cpaJIRAtemplateFilePath) -Level Info

$cpaJIRATempdataExcelWorksheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $cpaJIRAtemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRATempdataExcelWorksheet = $cpaJIRATemplateExcelWorkbook.Worksheets.Add()
$cpaJIRATempdataExcelWorksheet.Name = $TempworksheetName
$cpaJIRATempdataExcelWorksheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $cpaJIRAtemplateFilePath) -Level Info

}


Write-Log -Message ("Selecting '" + $cpaJIRATemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $cpaJIRAtemplateFilePath) -Level Info

$cpaJIRAdataExcelWorksheet = $cpaJIRATemplateExcelWorkbook.Worksheets.Item($cpaJIRATemplateRawDataWorkSheetName)
$cpaJIRAdataExcelWorksheet.Activate()
#Delete Old Records from Sheet "CPA RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To N in '" + $cpaJIRATemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRAdataExcelWorksheetDeleteRange = $cpaJIRAdataExcelWorksheet.Range("A:L")
$cpaJIRAdataExcelWorksheetDeleteRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To N in '" + $cpaJIRATemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRAdataExcelWorksheetRange = $cpaJIRAdataExcelWorksheet.Range("A1")
Write-Log -Message ("Copying Raw Data from Worksheet [" + $cpaJIRArawdataExcelWorksheet.Name + "] of File --> " + $cpaJIRADailyReportFilePath + " to Worksheet [" + $cpaJIRATemplateRawDataWorkSheetName + " of File --> " +  $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRAdataExcelWorksheet.Paste($cpaJIRAdataExcelWorksheetRange)
Write-Log -Message ("Copied Raw Data from Worksheet [" + $cpaJIRArawdataExcelWorksheet.Name + "] of File --> " + $cpaJIRADailyReportFilePath + " to Worksheet [" + $cpaJIRATemplateRawDataWorkSheetName + "] of File --> " +  $cpaJIRAtemplateFilePath) -Level Info
#Save the CPA Jira Template File



#Apply Filter the CPA Daily Incident Data in "CPA Data 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of CPA Daily JIRA Ticket Data in Worksheet [" + $cpaJIRATemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$cpaJIRAdataHurdleExcelWorksheet = $cpaJIRATemplateExcelWorkbook.Worksheets.Item($cpaJIRATemplateHuddleSheetName)
$cpaJIRAdataHurdleExcelWorksheet.Activate()
$cpaJIRAdataHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$cpaJIRAdataHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $cpaJIRATemplateHuddleSheetName  + "]") -Level Info
$cpaJIRAdataHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $cpaJIRATemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$cpaJIRATempdataExcelWorksheet.Activate()
$cpaJIRATempdataExcelWorksheet.Range("A1").PasteSpecial(-4163)
$cpaJIRATempdataRange = $cpaJIRATempdataExcelWorksheet.UsedRange

$cpaJIRATicketCount = $cpaJIRATemplateExcelFileObject.WorksheetFunction.CountIf($cpaJIRATempdataRange.Range("A1:" + "A" + $cpaJIRATempdataRange.Rows.Count), "<>") - 1

Write-Log -Message ("Total Number of Jira Tickets in CPA Daily JIRA Ticket are : [" + $cpaJIRATicketCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $cpaJIRAtemplateFilePath) -Level Info
#$cpaJIRATemplateExcelWorkbook.Save()
$cpaJIRATemplateExcelWorkbook.close($true)
$cpaJIRATemplateExcelFileObject.Quit()
$cpaJIRArawdataExcelFileObject.Quit()
Write-Log -Message ("**************** PROCESSED CPA JIRA DAILY INCIDENT RAW DATA *************************************") -Level Info

#Declare File Name, Sheet Name for NWR Daily JIRA Reports

Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING NETWORK RAIL TICKET RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read Network Rail Daily Ticket Data") -Level Info


$TempworksheetName = "Temp"
$nwrTemplateConsolidatedWorkSheetName = "NWR Consoldiate"
$nwrTemplateHuddleSheetName = "NWR Data 4 Huddle"
$nwrtemplateFilePath =  $parentFolderPath + $nwrTemplateFileName
$nwrOutStandingReportFilePath =  $RawDataParentFolderPath + $nwrOutstandingReportFileName
$nwrResolvedReportFilePath =  $RawDataParentFolderPath + $nwrResolvedReportFileName


Write-Log -Message ("NWR Incident Path --> " + $nwrIncidentParentFolderPath) -Level Info
Write-Log -Message ("NWR OutStanding Incident  Report Data Path --> " + $nwrOutStandingReportFilePath) -Level Info
Write-Log -Message ("NWR Resolved Incident Report Data Path --> " + $nwrResolvedReportFilePath) -Level Info
Write-Log -Message ("NWR Template File --> " + $nwrResolvedReportFilePath) -Level Info

#Declare File Name, Sheet Name for NWR Outstanding Incident Reports

Write-Log -Message ("**************** PROCESSING NETWORK RAIL OUTSTANDING INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read NWR OutStanding Incident Report Data") -Level Info

#Declare Excel Object for CPA Remedy Incident Data

Write-Log -Message ("Going to Create Excel Object for NWR Daily Incident Workbook") -Level Info
$nwrOutstandingdataExcelFileObject = New-Object -ComObject excel.application
$nwrOutstandingdataExcelFileObject.Visible = $true
$nwrOutstandingdataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NWR Outstanding Incident Report Data" + $nwrOutStandingReportFilePath) -Level Info
$nwrOutStandingdataExcelWorkbook = $nwrOutstandingdataExcelFileObject.Workbooks.Open($nwrOutStandingReportFilePath)
$nwrOutStandingdataExcelWorkbook.activate()
$nwrOutStandingdataExcelWorksheet = $nwrOutStandingdataExcelWorkbook.Worksheets.Item(1)
$nwrOutstandingReportSheetName = $nwrOutStandingdataExcelWorksheet.Name
$nwrOutStandingdataExcelWorksheet.Activate()
$nwrOutStandingdataExcelWorksheetRange = $nwrOutStandingdataExcelWorksheet.Range("A:AH")


#$LastRow = $nwrOutStandingdataExcelWorksheetRange.Cells($nwrOutStandingdataExcelWorksheetRange.Rows.Count, "A").End(-4162).Row


Write-Log -Message ("Copying Cells from A to AH in " + $nwrOutStandingdataExcelWorksheet.Name + " from " + $nwrOutStandingReportFilePath  + " File." ) -Level Info
$nwrOutStandingdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to AH in " + $nwrOutStandingdataExcelWorksheet.Name + " from " + $nwrOutStandingReportFilePath  + " File.") -Level Info


# Copy the NWR Outstanding Raw Data to NWR Template File

Write-Log -Message ("Going to Open NWR Template File") -Level Info
$nwrTemplateExcelFileObject = New-Object -ComObject excel.application
$nwrTemplateExcelFileObject.Visible = $true
$nwrTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NWR Template File --> " + $nwrtemplateFilePath) -Level Info
$nwrTemplateExcelWorkbook = $nwrTemplateExcelFileObject.Workbooks.Open($nwrtemplateFilePath)
$nwrTemplateExcelWorkbook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $nwrTemplateExcelWorkbook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $nwrtemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $nwrtemplateFilePath) -Level Info
$nwrTempdataExcelWorksheet = $nwrTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$nwrTempdataExcelWorksheet.Activate()
$nwrTempdataExcelWorksheetRange = $nwrTempdataExcelWorksheet.Range("A:AH")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $nwrtemplateFilePath) -Level Info

$nwrTempdataExcelWorksheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $nwrtemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $nwrtemplateFilePath) -Level Info
$nwrTempdataExcelWorksheet = $nwrTemplateExcelWorkbook.Worksheets.Add()
$nwrTempdataExcelWorksheet.Name = $TempworksheetName
$nwrTempdataExcelWorksheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $nwrtemplateFilePath) -Level Info

}


Write-Log -Message ("Selecting '" + $nwrTempdataExcelWorksheet.Name + "' WorkSheet in file --> " + $nwrtemplateFilePath) -Level Info

$nwrOutStandingdataTemplateExcelWorksheet = $nwrTemplateExcelWorkbook.Worksheets.Item($nwrTemplateConsolidatedWorkSheetName)
$nwrOutStandingdataTemplateExcelWorksheet.Activate()
#Delete Old Records from Sheet "NWR Consolidate OutStanding RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To AH in '" + $nwrTemplateConsolidatedWorkSheetName + "' WorkSheet from file --> " + $nwrtemplateFilePath) -Level Info
$nwrOutStandingdataExcelWorksheetDeleteRange = $nwrOutStandingdataTemplateExcelWorksheet.Range("A:AH")
$nwrOutStandingdataExcelWorksheetDeleteRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To AH in '" + $nwrTemplateConsolidatedWorkSheetName + "' WorkSheet from file --> " + $nwrtemplateFilePath) -Level Info
$nwrOutStandingdataTemplateExcelWorksheetRange = $nwrOutStandingdataTemplateExcelWorksheet.Range("A1")
Write-Log -Message ("Copying Raw Data from Worksheet [" + $nwrOutStandingdataExcelWorksheet.Name  + "] of File --> " + $nwrOutStandingReportFilePath + " to Worksheet [" + $nwrTemplateConsolidatedWorkSheetName + " of File --> " +  $nwrtemplateFilePath) -Level Info
$nwrOutStandingdataTemplateExcelWorksheet.Paste($nwrOutStandingdataTemplateExcelWorksheetRange)
$nwrOutStandingdataTemplateExcelWorksheetRange = $nwrOutStandingdataTemplateExcelWorksheet.UsedRange

$lastNWROutStandingTicketCount = $nwrTemplateExcelFileObject.WorksheetFunction.CountIf($nwrOutStandingdataTemplateExcelWorksheetRange.Range("A1:" + "A" + $nwrOutStandingdataTemplateExcelWorksheetRange.Rows.Count), "<>") - 1
Write-Log -Message ("Total Number of Network Rail Outstanding Incidents are [" + $lastNWROutStandingTicketCount +"]") -Level Info

Write-Log -Message ("Copied Raw Data from Worksheet [" + $nwrOutStandingdataTemplateExcelWorksheet.Name + "] of File --> " + $nwrOutStandingReportFilePath + " to Worksheet [" + $nwrTemplateConsolidatedWorkSheetName + "] of File --> " +  $nwrtemplateFilePath) -Level Info

#Save the NWR OutStanding Workbook File
Write-Log -Message ("Closing NWR OutStanding Incident Report Data File --> " + $nwrOutStandingReportFilePath) -Level Info
$nwrOutstandingdataExcelFileObject.Quit()
Write-Log -Message ("Closed NWR OutStanding Incident Report Data File --> " + $nwrOutStandingReportFilePath) -Level Info


Write-Log -Message ("**************** PROCESSING NWR RESOLVED INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read NWR Resolved Incident Report Data") -Level Info

$nwrResolveddataExcelFileObject = New-Object -ComObject excel.application
$nwrResolveddataExcelFileObject.Visible = $true
$nwrResolveddataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NWR Resolved Incident Report Data :" + $nwrResolvedReportFilePath) -Level Info
$nwrResolveddataExcelWorkbook = $nwrResolveddataExcelFileObject.Workbooks.Open($nwrResolvedReportFilePath)
$nwrResolveddataExcelWorkbook.activate()
$nwrResolveddataExcelWorksheet = $nwrResolveddataExcelWorkbook.Worksheets.Item(1)
$nwrResolveddataExcelWorksheetName = $nwrResolveddataExcelWorksheet.Name
$nwrResolveddataExcelWorksheet.Activate()
$nwrResolveddataExcelWorksheetRange = $nwrResolveddataExcelWorksheet.Range("A:AH")
Write-Log -Message ("Copying Cells from A to AH in " + $nwrResolveddataExcelWorksheetName  + " from " + $nwrResolvedReportFilePath  + " File." ) -Level Info
$nwrResolveddataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to AH in " + $nwrResolveddataExcelWorksheetName + " from " + $nwrResolvedReportFilePath  + " File.") -Level Info

$nwrResolveddataTemplateExcelWorksheet = $nwrTemplateExcelWorkbook.Worksheets.Item($nwrTemplateConsolidatedWorkSheetName)
$nwrResolveddataTemplateExcelWorksheet.Activate()
#Delete Old Records from Sheet "NWR Resolved Date RAW DAta"
$OutStandingLastRow = $lastNWROutStandingTicketCount
$CopyRowPostition = $OutStandingLastRow + 2
$nwrResolveddataTemplateExcelWorksheetRange = $nwrResolveddataTemplateExcelWorksheet.Range("A$CopyRowPostition")

Write-Log -Message ("Copying Raw Data from Worksheet [" + $nwrResolveddataExcelWorksheetName   + "] of File --> " + $nwrResolvedReportFilePath + " to Worksheet [" + $nwrTemplateConsolidatedWorkSheetName + " of File --> " +  $nwrtemplateFilePath) -Level Info
$nwrResolveddataTemplateExcelWorksheet.Paste($nwrResolveddataTemplateExcelWorksheetRange)
Write-Log -Message ("Copied Raw Data from Worksheet [" + $nwrResolveddataExcelWorksheetName + "] of File --> " + $nwrResolvedReportFilePath + " to Worksheet [" + $nwrTemplateConsolidatedWorkSheetName + "] of File --> " +  $nwrtemplateFilePath) -Level Info


$nwrResolvedHeaderdataRange = $nwrResolveddataTemplateExcelWorksheet.Cells.Item($CopyRowPostition,1).EntireRow
$nwrResolvedHeaderdataRange.Delete()

Write-Log -Message ("Closing NWR Resolved Incident Report Data File --> " + $nwrResolvedReportFilePath) -Level Info
$nwrResolveddataExcelFileObject.Quit()
Write-Log -Message ("Closed NWR Resolved Incident Report Data File --> " + $nwrResolvedReportFilePath) -Level Info
$nwrJIRAdataHurdleExcelWorksheet = $nwrTemplateExcelWorkbook.Worksheets.Item($nwrTemplateHuddleSheetName)
$nwrJIRAdataHurdleExcelWorksheet.Activate()
$nwrJIRAdataHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$nwrJIRAdataHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $nwrTemplateHuddleSheetName  + "]") -Level Info
$nwrJIRAdataHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $nwrTemplateHuddleSheetName.Name  + "] to Worksheet [" + $cpaJIRATempdataExcelWorksheet + "]") -Level Info
$nwrTempdataExcelWorksheet.Activate()

$nwrTempdataExcelWorksheet.Range("A1").PasteSpecial(-4163)
#$cpaJIRATempdataExcelWorksheet.Range("A1").PasteSpecial(13)
#xlPasteAllUsingSourceTheme
$nwrTempdataRange = $nwrTempdataExcelWorksheet.UsedRange

$TotalNWRIncidentCount = $nwrTemplateExcelFileObject.WorksheetFunction.CountIf($nwrTempdataRange.Range("A1:" + "A" + $nwrTempdataRange.Rows.Count), "<>") - 1
Write-Log -Message ("Total Number of Incident in NWR Outstand and Resolved are : [" + $TotalNWRIncidentCount   + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $nwrtemplateFilePath) -Level Info
#$nwrTemplateExcelWorkbook.Save()
$nwrTemplateExcelWorkbook.close($true)
$nwrTemplateExcelFileObject.Quit()
Write-Log -Message ("**************** PROCESSED NWR OUTSTANDING AND RESOLVED DAILY INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

#Declare File Name, Sheet Name for Xchanging Daily Incident Reports

Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING XChanging OPEN TICKET RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read XChanging Daily Ticket Data") -Level Info


$xchangingOpenIncidentSheetName = "Open Incidents"
$xchangingResolvedIncidentSheetName = "Resolved and closed"
$xchangingConsolidateIncidentSheetName = "Copied Incident Dump"
$TempworksheetName = "Temp"
$xchangingTemplateHuddleSheetName = "Xchanging data 4 Huddle"
$xchangingtemplateFilePath =  $parentFolderPath + $xchangingTemplateFileName
$xchangingIncidentParentFolderPath =  $RawDataParentFolderPath + $xchangingReportFileName


#Declare File Name, Sheet Name for XChanging Open Incident Reports

Write-Log -Message ("**************** PROCESSING XCHANGING OPEN INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read XChanging Open Incident Report Data") -Level Info

#Declare File Name, Sheet Name for xchanging Outstanding Incident Reports

Write-Log -Message ("Going to Create Excel Object for XChanging Open Incident Workbook") -Level Info
$xchangingIncidentdataExcelFileObject = New-Object -ComObject excel.application
$xchangingIncidentdataExcelFileObject.Visible = $true
$xchangingIncidentdataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening xchanging Incident Report Data" + $xchangingReportFileName) -Level Info
$xchangingIncidentdataExcelWorkbook = $xchangingIncidentdataExcelFileObject.Workbooks.Open($xchangingIncidentParentFolderPath)
$xchangingIncidentdataExcelWorkbook.activate()
$xchangingIncidentdataExcelWorksheet = $xchangingIncidentdataExcelWorkbook.Worksheets.Item($xchangingOpenIncidentSheetName)
$xchangingIncidentdataExcelWorksheet.Activate()
$xchangingIncidentdataExcelWorksheetRange = $xchangingIncidentdataExcelWorksheet.Range("A:J")
$xchangingOpenIncidentHeaderdataRange = $xchangingIncidentdataExcelWorksheet.Cells.Item(1,1).EntireRow
$xchangingOpenIncidentHeaderdataRange.Delete()

Write-Log -Message ("Copying Cells from A to J in " + $xchangingIncidentdataExcelWorksheet.Name + " from " + $xchangingIncidentParentFolderPath  + " File." ) -Level Info
$xchangingIncidentdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to J in " + $xchangingIncidentdataExcelWorksheet.Name + " from " + $xchangingIncidentParentFolderPath  + " File.") -Level Info

Write-Log -Message ("Going to Open xchanging Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for XChanging Template Hurdle Workbook") -Level Info
$xchangingTemplateExcelFileObject = New-Object -ComObject excel.application
$xchangingTemplateExcelFileObject.Visible = $true
$xchangingTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening xchanging Template File --> " + $xchangingtemplateFilePath) -Level Info
$xchangingTemplateExcelWorkbook = $xchangingTemplateExcelFileObject.Workbooks.Open($xchangingtemplateFilePath)
$xchangingTemplateExcelWorkbook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $xchangingTemplateExcelWorkbook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $xchangingtemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }




if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $xchangingtemplateFilePath) -Level Info
$xchangingTempdataExcelWorksheet = $xchangingTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$xchangingTempdataExcelWorksheet.Activate()
$xchangingTempdataExcelWorksheetRange = $xchangingTempdataExcelWorksheet.Range("A:AJ")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $xchangingtemplateFilePath) -Level Info

$xchangingTempdataExcelWorksheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $xchangingtemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $xchangingtemplateFilePath) -Level Info
$xchangingTempdataExcelWorksheet = $xchangingTemplateExcelWorkbook.Worksheets.Add()
$xchangingTempdataExcelWorksheet.Name = $TempworksheetName
$xchangingTempdataExcelWorksheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $xchangingtemplateFilePath) -Level Info

}



Write-Log -Message ("Selecting '" + $xchangingTempdataExcelWorksheet.Name + "' WorkSheet in file --> " + $xchangingtemplateFilePath) -Level Info

$xchangingConsolidateIncidentWorkSheet = $xchangingTemplateExcelWorkbook.Worksheets.Item($xchangingConsolidateIncidentSheetName)
$xchangingConsolidateIncidentWorkSheet.Activate()
#Delete Old Records from Sheet "xchanging Consolidate OutStanding RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To J in '" + $xchangingConsolidateIncidentSheetName + "' WorkSheet from file --> " + $xchangingtemplateFilePath) -Level Info
$xchangingConsolidateIncidentWorksheetDeleteRange = $xchangingConsolidateIncidentWorkSheet.Range("A:J")
$xchangingConsolidateIncidentWorksheetDeleteRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To J in '" + $xchangingConsolidateIncidentSheetName + "' WorkSheet from file --> " + $xchangingtemplateFilePath) -Level Info
$xchangingConsolidateIncidentWorkSheetRange = $xchangingConsolidateIncidentWorkSheet.Range("A1")
Write-Log -Message ("Copying Raw Data from Worksheet [" + $xchangingIncidentdataExcelWorksheet.Name  + "] of File --> " + $xchangingIncidentParentFolderPath + " to Worksheet [" + $xchangingConsolidateIncidentSheetName + "] of File --> " +  $xchangingtemplateFilePath) -Level Info

$xchangingConsolidateIncidentWorkSheet.Paste($xchangingConsolidateIncidentWorkSheetRange)

$xchangingConsolidateIncidentWorkSheetRange = $xchangingConsolidateIncidentWorkSheet.UsedRange


$xchanginglastOpenIncidentRow = $xchangingTemplateExcelFileObject.WorksheetFunction.CountIf($xchangingConsolidateIncidentWorkSheet.Range("A1:" + "A" + $xchangingConsolidateIncidentWorkSheetRange.Rows.Count), "<>") - 1

#$xchangingOpenIncidentTotalCount = $xchangingConsolidateIncidentWorkSheetRange.Rows.Count
Write-Log -Message ("Total Number of Open Incidents are [" + $xchanginglastOpenIncidentRow +"]") -Level Info


Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING XChanging RESOLVED TICKET RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read XChanging Resolved & Closed Daily Ticket Data") -Level Info


Write-Log -Message ("Going to Read Resolved Incidents in Worksheet [" + $xchangingResolvedIncidentSheetName + "] from " + $xchangingIncidentParentFolderPath  + " File." ) -Level Info

$xchangingIncidentdataExcelWorkbook.activate()
$xchangingResolvedIncidentdataExcelWorksheet = $xchangingIncidentdataExcelWorkbook.Worksheets.Item($xchangingResolvedIncidentSheetName)
$xchangingResolvedIncidentdataExcelWorksheet.Activate()

$xchangingResolvedIncidentHeaderdataRange = $xchangingResolvedIncidentdataExcelWorksheet.Cells.Item(2,1).EntireRow
$xchangingResolvedIncidentHeaderdataRange.Delete()
$xchangingResolvedIncidentHeaderdataRange = $xchangingResolvedIncidentdataExcelWorksheet.Cells.Item(1,1).EntireRow
$xchangingResolvedIncidentHeaderdataRange.Delete()



$xchangingResolvedIncidentdataExcelWorksheetRange = $xchangingResolvedIncidentdataExcelWorksheet.Range("A:J")

Write-Log -Message ("Copying Cells from A to J in [" + $xchangingResolvedIncidentdataExcelWorksheet.Name + "] from " + $xchangingIncidentParentFolderPath  + " File." ) -Level Info
$xchangingResolvedIncidentdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to J in [" + $xchangingResolvedIncidentdataExcelWorksheet.Name + "] from " + $xchangingIncidentParentFolderPath  + " File.") -Level Info

Write-Log -Message ("Selecting '" + $xchangingConsolidateIncidentSheetName + "' WorkSheet in file --> " + $xchangingtemplateFilePath) -Level Info

$xchangingTemplateExcelWorkbook.activate()

$xchangingConsolidateIncidentWorkSheet.Activate()
#Delete Old Records from Sheet "xchanging Consolidate OutStanding RAW DAta"
$xchangingFirstResolvedRow = $xchanginglastOpenIncidentRow + 2
$xchangingConsolidateIncidentWorkSheetRange = $xchangingConsolidateIncidentWorkSheet.Range("A$xchangingFirstResolvedRow")
Write-Log -Message ("Copying Raw Data from Worksheet [" + $xchangingResolvedIncidentdataExcelWorksheet.Name  + "] of File --> " + $xchangingIncidentParentFolderPath + " to Worksheet [" + $xchangingConsolidateIncidentSheetName + " of File --> " +  $xchangingtemplateFilePath) -Level Info
$xchangingConsolidateIncidentWorkSheet.Paste($xchangingConsolidateIncidentWorkSheetRange)
$xchangingConsolidateIncidentWorkSheetRange = $xchangingConsolidateIncidentWorkSheet.UsedRange
$xchangingTotalIncidentCount = $xchangingTemplateExcelFileObject.WorksheetFunction.CountIf($xchangingConsolidateIncidentWorkSheet.Range("A1:" + "A" + $xchangingConsolidateIncidentWorkSheetRange.Rows.Count), "<>") - 1

#$xchangingOpenIncidentTotalCount = $xchangingConsolidateIncidentWorkSheetRange.Rows.Count
Write-Log -Message ("Total Number of Incidents are [" + $xchangingTotalIncidentCount +"]") -Level Info


Write-Log -Message ("Closing xChanging Daily Incident Report Data File --> " + $xchangingIncidentParentFolderPath) -Level Info
$xchangingIncidentdataExcelFileObject.Quit()
Write-Log -Message ("Closed xChanging Daily Incident Report Data File --> " + $xchangingIncidentParentFolderPath) -Level Info

$xChangingdataHurdleExcelWorksheet = $xchangingTemplateExcelWorkbook.Worksheets.Item($xchangingTemplateHuddleSheetName)
$xChangingdataHurdleExcelWorksheet.Activate()
$xChangingdataHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$xChangingdataHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $xchangingTemplateHuddleSheetName  + "]") -Level Info
$xChangingdataHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $xchangingTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$xchangingTempdataExcelWorksheet.Activate()

$xchangingTempdataExcelWorksheet.Range("A1").PasteSpecial(-4163)
#$cpaJIRATempdataExcelWorksheet.Range("A1").PasteSpecial(13)
#xlPasteAllUsingSourceTheme
$xchangingTempdataExcelWorksheetRange = $xchangingTempdataExcelWorksheet.UsedRange

$TotalXchangingIncidentCount = $xchangingTemplateExcelFileObject.WorksheetFunction.CountIf($xchangingTempdataExcelWorksheetRange.Range("A1:" + "A" + $xchangingTempdataExcelWorksheetRange.Rows.Count), "<>") - 1
Write-Log -Message ("Total Number of Incident in XChangined Open and Resolved/ Closed Worksheet are : [" + $TotalXchangingIncidentCount   + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $xchangingtemplateFilePath) -Level Info
#$xchangingTemplateExcelWorkbook.SAVE()
$xchangingTemplateExcelWorkbook.close($true)
$xchangingTemplateExcelFileObject.Quit()
Write-Log -Message ("**************** PROCESSED XCHANGING OPEN and RESOLVED DAILY INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING QBE DAILY INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read QBE Daily Incident Data") -Level Info

#Declare File Name, Sheet Name for CPA Daily Incident Remedy Reports

$qbeTemplateRawDataWorkSheetName = "Raw Data"
$qbeTemplateHuddleSheetName = "QBE Data 4 Huddle"
$TempworksheetName = "Temp"
$qbetemplateFilePath =  $parentFolderPath + $qbeTemplateFileName
$qbeDailyReportFilePath =  $RawDataParentFolderPath + $qbeIncidentDailyReportFileName

Write-Log -Message ("QBE Ticket File Path --> " + $qbeDailyReportFilePath) -Level Info
Write-Log -Message ("QBE Template File Path --> " + $qbetemplateFilePath) -Level Info

#Declare Excel Object for QBE Incident Data
Write-Log -Message ("Going to Create Excel Object for QBE Daily Incident Workbook") -Level Info
$qbeRAWdataExcelFileObject = New-Object -ComObject excel.application
$qbeRAWdataExcelFileObject.Visible = $true
$qbeRAWdataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening QBE Daily Incident File : " + $qbeDailyReportFilePath) -Level Info
$qbeRAWdataExcelWorkbook = $qbeRAWdataExcelFileObject.Workbooks.Open($qbeDailyReportFilePath)
$qbeRAWdataExcelWorkbook.activate()
Write-Log -Message ("Selecting QBE Daily Incident Worksheet") -Level Info

$qberawdataExcelWorksheet = $qbeRAWdataExcelWorkbook.Worksheets.Item(1)
$qbeIncidentRawDataWorkSheetName = $qberawdataExcelWorksheet.Name

Write-Log -Message ("Selected QBE Daily Incident Worksheet : [" + $qbeIncidentRawDataWorkSheetName + "]") -Level Info

$qberawdataExcelWorksheet.Activate()
$qberawdataExcelWorksheetRange = $qberawdataExcelWorksheet.Range("A:X").CurrentRegion
Write-Log -Message ("Copying Cells from A to X in Worksheet : [" + $qbeIncidentRawDataWorkSheetName + "] from " + $qbeDailyReportFilePath  + " File." ) -Level Info
$qberawdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to X in Worksheet : [" + $qbeIncidentRawDataWorkSheetName + "] from " + $qbeDailyReportFilePath  + " File.") -Level Info

# Copy the QBE Daily Raw Data to QBE Template File

Write-Log -Message ("Going to Open QBE Template File") -Level Info
$qbeTemplateExcelFileObject = New-Object -ComObject excel.application
$qbeTemplateExcelFileObject.Visible = $true
$qbeTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening QBE Template File --> " + $qbetemplateFilePath) -Level Info
$qbeTemplateExcelWorkbook = $qbeTemplateExcelFileObject.Workbooks.Open($qbetemplateFilePath)
$qbeTemplateExcelWorkbook.activate()



# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $qbeTemplateExcelWorkbook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $qbetemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }

Write-Log -Message ("Selecting QBE Template Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{

Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $qbetemplateFilePath) -Level Info
$qbeTempExcelWorksheet = $qbeTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$qbeTempExcelWorksheet.Activate()
$qbeTempExcelWorksheetRange = $qbeTempExcelWorksheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $qbetemplateFilePath) -Level Info

$qbeTempExcelWorksheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $qbetemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $qbetemplateFilePath) -Level Info
$qbeTempExcelWorksheet = $qbeTemplateExcelWorkbook.Worksheets.Add()
$qbeTempExcelWorksheet.Name = $TempworksheetName
$qbeTempExcelWorksheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $qbetemplateFilePath) -Level Info

}



Write-Log -Message ("Selecting '" + $qbeTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $qbetemplateFilePath) -Level Info

$qbeTemplateIncidentRawExcelWorksheet = $qbeTemplateExcelWorkbook.Worksheets.Item($qbeTemplateRawDataWorkSheetName)
$qbeTemplateIncidentRawExcelWorksheet.Activate()
#Delete Old Records from Sheet "QBE RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To X in '" + $qbeTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $qbetemplateFilePath) -Level Info
$qbeTemplateIncidentRawExcelWorksheetRange = $qbeTemplateIncidentRawExcelWorksheet.Range("A:X")
$qbeTemplateIncidentRawExcelWorksheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To X in '" + $qbeTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $qbetemplateFilePath) -Level Info
$qbeTemplateIncidentRawExcelWorksheetRange = $qbeTemplateIncidentRawExcelWorksheetRange.Range("A1")


Write-Log -Message ("Copying QBE Incident Raw Data from Worksheet [" + $qbeIncidentRawDataWorkSheetName + "] of File --> " + $qbeDailyReportFilePath + " to Worksheet [" + $qbeTemplateRawDataWorkSheetName + " of File --> " +  $qbetemplateFilePath) -Level Info
$qbeTemplateIncidentRawExcelWorksheet.Paste($qbeTemplateIncidentRawExcelWorksheetRange)
Write-Log -Message ("Copied QBE Incident Raw Data from Worksheet [" + $qbeIncidentRawDataWorkSheetName + "] of File --> " + $qbeDailyReportFilePath + " to Worksheet [" + $qbeTemplateRawDataWorkSheetName + "] of File --> " +  $qbetemplateFilePath) -Level Info
#Save the CPA Jira Template File
#$cpaJIRATemplateExcelWorkbook.Save()
#Write-Log -Message ("Closing CPA JIRA Daily Report Data File --> " + $cpaJIRADailyReportFilePath) -Level Info
#$cpaJIRArawdataExcelFileObject.Quit()
#Write-Log -Message ("Closed CPA JIRA Daily Report Data File --> " + $cpaJIRADailyReportFilePath) -Level Info

#Apply Filter the CPA Daily Incident Data in "CPA Data 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of QBE Daily Incident Data in Worksheet [" + $qbeTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$qbeTemplateHurdleExcelWorksheet = $qbeTemplateExcelWorkbook.Worksheets.Item($qbeTemplateHuddleSheetName)
$qbeTemplateHurdleExcelWorksheet.Activate()
$qbeTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$qbeTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$qbeTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $qbeTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$qbeTempExcelWorksheet.Activate()

$qbeTempExcelWorksheet.Range("A1").PasteSpecial(-4163)
$qbeTempExcelWorksheetRange = $qbeTempExcelWorksheet.UsedRange

$qbeIncidentCount = $qbeTemplateExcelFileObject.WorksheetFunction.CountIf($qbeTempExcelWorksheetRange.Range("A1:" + "A" + $qbeTempExcelWorksheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of QBE Daily Incident Count is : [" + $cpaRemedyIncidentCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $qbetemplateFilePath) -Level Info
#$qbeTemplateExcelWorkbook.Save()
$qbeTemplateExcelWorkbook.close($true)

$qbeTemplateExcelFileObject.Quit()
$qbeRAWdataExcelFileObject.Quit()
Write-Log -Message ("**************** PROCESSED QBE DAILY INCIDENT RAW DATA *************************************") -Level Info
Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING QBE SERVICE REQUEST RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read QBE Daily Service Request Data") -Level Info

#Declare File Name, Sheet Name for QBE SR Raw data

$qbeSRTemplateRawDataWorkSheetName = "Raw Data"
$qbeSRTemplateHuddleSheetName = "QBE Data 4 Huddle"
$TempworksheetName = "Temp"
$qbeSRtemplateFilePath =  $parentFolderPath + $qbeSRTemplateFileName
$qbeSRDailyReportFilePath =  $RawDataParentFolderPath + $qbeSRDailyReportFileName

Write-Log -Message ("QBE Ticket File Path --> " + $qbeSRDailyReportFilePath) -Level Info
Write-Log -Message ("QBE Template File Path --> " + $qbeSRtemplateFilePath) -Level Info

#Declare Excel Object for QBE Incident Data
Write-Log -Message ("Going to Create Excel Object for QBE Daily SR Workbook") -Level Info
$qbeSRRAWdataExcelFileObject = New-Object -ComObject excel.application
$qbeSRRAWdataExcelFileObject.Visible = $true
$qbeSRRAWdataExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening QBE Daily SR File : " + $qbeSRDailyReportFilePath) -Level Info
$qbeSRRAWdataExcelWorkbook = $qbeSRRAWdataExcelFileObject.Workbooks.Open($qbeSRDailyReportFilePath)
$qbeSRRAWdataExcelWorkbook.activate()

Write-Log -Message ("Selecting QBE Daily SR Worksheet") -Level Info
$qbeSRrawdataExcelWorksheet = $qbeSRRAWdataExcelWorkbook.Worksheets.Item(1)

$qbeSRRawDataWorkSheetName = $qbeSRrawdataExcelWorksheet.Name
Write-Log -Message ("Selected QBE Daily SR Worksheet : [" + $qbeSRRawDataWorkSheetName + "]") -Level Info

$qbeSRrawdataExcelWorksheet.Activate()
$qbeSRdataExcelWorksheetRange = $qbeSRrawdataExcelWorksheet.Range("A:K").CurrentRegion
Write-Log -Message ("Copying Cells from A to K in Worksheet : [" + $qbeSRRawDataWorkSheetName + "] from " + $qbeSRDailyReportFilePath  + " File." ) -Level Info
$qbeSRdataExcelWorksheetRange.copy()
Write-Log -Message ("Copied Cells from A to K in Worksheet : [" + $qbeSRRawDataWorkSheetName + "] from " + $qbeSRDailyReportFilePath  + " File.") -Level Info

# Copy the QBE Daily SR Data to QBE Template File

Write-Log -Message ("Going to Open QBE SR Template File") -Level Info
$qbeSRTemplateExcelFileObject = New-Object -ComObject excel.application
$qbeSRTemplateExcelFileObject.Visible = $true
$qbeSRTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening QBE SR Template File --> " + $qbeSRtemplateFilePath) -Level Info
$qbeSRTemplateExcelWorkbook = $qbeSRTemplateExcelFileObject.Workbooks.Open($qbeSRtemplateFilePath)
$qbeSRTemplateExcelWorkbook.activate()



# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $qbeSRTemplateExcelWorkbook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $qbeSRtemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }

Write-Log -Message ("Selecting QBE SR Temp Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{

Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $qbeSRtemplateFilePath) -Level Info
$qbeSRTempExcelWorksheet = $qbeSRTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$qbeSRTempExcelWorksheet.Activate()
$qbeSRTempExcelWorksheetRange = $qbeSRTempExcelWorksheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $qbeSRtemplateFilePath) -Level Info

$qbeSRTempExcelWorksheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $qbeSRtemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $qbeSRtemplateFilePath) -Level Info
$qbeSRTempExcelWorksheet = $qbeSRTemplateExcelWorkbook.Worksheets.Add()
$qbeSRTempExcelWorksheet.Name = $TempworksheetName
$qbeSRTempExcelWorksheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $qbeSRtemplateFilePath) -Level Info

}



Write-Log -Message ("Selecting '" + $qbeSRTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $qbeSRtemplateFilePath) -Level Info

$qbeTemplateSRRawExcelWorksheet = $qbeSRTemplateExcelWorkbook.Worksheets.Item($qbeSRTemplateRawDataWorkSheetName)
$qbeTemplateSRRawExcelWorksheet.Activate()
#Delete Old Records from Sheet "QBE SR RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To K in '" + $qbeSRTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $qbeSRtemplateFilePath) -Level Info
$qbeTemplateSRRawExcelWorksheetRange = $qbeTemplateSRRawExcelWorksheet.Range("A:K")
$qbeTemplateSRRawExcelWorksheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To K in '" + $qbeSRTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $qbeSRtemplateFilePath) -Level Info
$qbeTemplateSRRawExcelWorksheetRange = $qbeTemplateSRRawExcelWorksheetRange.Range("A1")


Write-Log -Message ("Copying QBE SR Raw Data from Worksheet [" + $qbeSRRawDataWorkSheetName + "] of File --> " + $qbeSRDailyReportFilePath + " to Worksheet [" + $qbeSRTemplateRawDataWorkSheetName + " of File --> " +  $qbeSRtemplateFilePath) -Level Info
$qbeTemplateSRRawExcelWorksheet.Paste($qbeTemplateSRRawExcelWorksheetRange)
Write-Log -Message ("Copied QBE SR Raw Data from Worksheet [" + $qbeSRRawDataWorkSheetName + "] of File --> " + $qbeSRDailyReportFilePath + " to Worksheet [" + $qbeSRTemplateRawDataWorkSheetName + "] of File --> " +  $qbeSRtemplateFilePath) -Level Info
#Save the CPA Jira Template File
#$cpaJIRATemplateExcelWorkbook.Save()
#Write-Log -Message ("Closing CPA JIRA Daily Report Data File --> " + $cpaJIRADailyReportFilePath) -Level Info
#$cpaJIRArawdataExcelFileObject.Quit()
#Write-Log -Message ("Closed CPA JIRA Daily Report Data File --> " + $cpaJIRADailyReportFilePath) -Level Info

#Apply Filter the CPA Daily Incident Data in "CPA Data 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of QBE Daily SR Data in Worksheet [" + $qbeSRTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$qbeSRTemplateHurdleExcelWorksheet = $qbeSRTemplateExcelWorkbook.Worksheets.Item($qbeSRTemplateHuddleSheetName)
$qbeSRTemplateHurdleExcelWorksheet.Activate()
$qbeSRTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$qbeSRTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$qbeSRTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $qbeSRTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$qbeSRTempExcelWorksheet.Activate()

$qbeSRTempExcelWorksheet.Range("A1").PasteSpecial(-4163)
$qbeSRTempExcelWorksheetRange = $qbeSRTempExcelWorksheet.UsedRange

$qbeSRCount = $qbeSRTemplateExcelFileObject.WorksheetFunction.CountIf($qbeSRTempExcelWorksheetRange.Range("A1:" + "A" + $qbeSRTempExcelWorksheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of QBE Daily SR Count is : [" + $qbeSRCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $qbeSRtemplateFilePath) -Level Info
#$qbeTemplateExcelWorkbook.Save()
$qbeSRTemplateExcelWorkbook.close($true)

$qbeSRRAWdataExcelFileObject.Quit()
$qbeSRTemplateExcelFileObject.Quit()
Write-Log -Message ("**************** PROCESSED QBE DAILY SR RAW DATA *************************************") -Level Info


Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING BELRON DAILY INCIDENT OUTSTANDING RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read BELRON Daily Incident OUTSTANDING  Report Data") -Level Info

#Declare File Name, Sheet Name for BELRON Daily Incident OUTSTANDING Reports

$belronTemplateFilePath =  $parentFolderPath + $belronTemplateFileName
$belronDailyIncidentOutStandingFilePath =  $RawDataParentFolderPath + $belronDailyIncidentOutStandingFileName
$belronDailyIncidentResolvedFilePath =  $RawDataParentFolderPath + $belronDailyIncidentResolvedFileName
$belronTemplateRawDataWorkSheetName = "Belron_Raw"
$belronTemplateHuddleSheetName = "Belron_Data 4 Huddle"
$TempworksheetName = "Temp"

Write-Log -Message ("BELRON Ticket OUTSTANDING File Path --> " + $belronDailyIncidentOutStandingFilePath) -Level Info
Write-Log -Message ("BELRON Template File --> " + $belronTemplateFilePath) -Level Info

#Declare Excel Object for BELRON OUTSTANDING Incident Data
Write-Log -Message ("Going to Create Excel Object for BELRON Daily Incident OUTSTANDING  Workbook") -Level Info
$belronDailyIncidentOutStandingFileObject = New-Object -ComObject excel.application
$belronDailyIncidentOutStandingFileObject.Visible = $true
$belronDailyIncidentOutStandingFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening BELRON Daily Incident OUTSTANDING File : " + $belronDailyIncidentOutStandingFilePath) -Level Info
$belronDailyIncidentOutStandingExcelWorkbook = $belronDailyIncidentOutStandingFileObject.Workbooks.Open($belronDailyIncidentOutStandingFilePath)
$belronDailyIncidentOutStandingExcelWorkbook.activate()
$belronDailyIncidentOutStandingExcelWorkSheet = $belronDailyIncidentOutStandingExcelWorkbook.Worksheets.Item(1)
$belronDailyIncidentOutStandingExcelWorkSheetName = $belronDailyIncidentOutStandingExcelWorkSheet.Name
Write-Log -Message ("Selecting BELRON Daily Incident Worksheet : [" + $belronDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$belronDailyIncidentOutStandingExcelWorkSheet.Activate()
$belronDailyIncidentOutStandingExcelWorkSheetRange = $belronDailyIncidentOutStandingExcelWorkSheet.Range("A:V").CurrentRegion
Write-Log -Message ("Copying Cells from A to V in Worksheet : [" + $belronDailyIncidentOutStandingExcelWorkSheetName + "] from " + $belronDailyIncidentOutStandingFilePath  + " File." ) -Level Info
$belronDailyIncidentOutStandingExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to V in Worksheet : [" + $belronDailyIncidentOutStandingExcelWorkSheetName + "] from " + $belronDailyIncidentOutStandingFilePath  + " File.") -Level Info

# Copy the BELRON Daily Remedy Raw Data to Template File

Write-Log -Message ("Going to Open BELRON Template File") -Level Info
$belronTemplateFileObject = New-Object -ComObject excel.application
$belronTemplateFileObject.Visible = $true
$belronTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening BELRON Template File --> " + $belronTemplateFilePath) -Level Info
$belronTemplateExcelWorkBook = $belronTemplateFileObject.Workbooks.Open($belronTemplateFilePath)
$belronTemplateExcelWorkBook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $belronTemplateExcelWorkBook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $belronTemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



Write-Log -Message ("Selecting BELRON Template Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $belronTemplateFilePath) -Level Info
$belronTempExcelWorkSheet = $belronTemplateExcelWorkBook.Worksheets.Item($TempworksheetName)
$belronTempExcelWorkSheet.Activate()
$belronTempExcelWorkSheetRange = $belronTempExcelWorkSheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $belronTemplateFilePath) -Level Info

$belronTempExcelWorkSheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $belronTemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $belronTemplateFilePath) -Level Info
$belronTempExcelWorkSheet = $belronTemplateExcelWorkBook.Worksheets.Add()
$belronTempExcelWorkSheet.Name = $TempworksheetName
$belronTempExcelWorkSheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $belronTemplateFilePath) -Level Info

}

Write-Log -Message ("Selecting '" + $belronTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $belronTemplateFilePath) -Level Info

$belronTemplateRawDataWorkSheet = $belronTemplateExcelWorkBook.Worksheets.Item($belronTemplateRawDataWorkSheetName)
$belronTemplateRawDataWorkSheet.Activate()
#Delete Old Records from Sheet "BELRON RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To V in '" + $belronTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $belronTemplateFilePath) -Level Info
$belronTemplateRawDataWorkSheetRange = $belronTemplateRawDataWorkSheet.Range("A:V")
$belronTemplateRawDataWorkSheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To V in '" + $belronTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $belronTemplateFilePath) -Level Info
$belronTemplateRawDataWorkSheetRange = $belronTemplateRawDataWorkSheetRange.Range("A1")
Write-Log -Message ("Copying BELRON Incident OUTSTANDING Raw Data from Worksheet [" + $belronDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $belronDailyIncidentOutStandingFilePath + " to Worksheet [" + $belronTemplateRawDataWorkSheetName + " of File --> " +  $belronTemplateFilePath) -Level Info
$belronTemplateRawDataWorkSheet.Paste($belronTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied BELRON  Incident OUTSTANDING Raw Data from Worksheet [" + $belronDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $belronDailyIncidentOutStandingFilePath + " to Worksheet [" + $belronTemplateRawDataWorkSheetName + "] of File --> " +  $belronTemplateFilePath) -Level Info

$belronDailyIncidentOutStandingFileObject.Quit()
$belronTemplateRawDataWorkSheetUsedRange = $belronTemplateRawDataWorkSheet.UsedRange

$belronOutstandingIncidentCount = $belronTemplateFileObject.WorksheetFunction.CountIf($belronTemplateRawDataWorkSheetUsedRange.Range("A1:" + "A" + $belronTemplateRawDataWorkSheetUsedRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of BELRON Daily Incident OUTSTANDING Count is : [" + $belronOutstandingIncidentCount  + "]") -Level Info



#Declare Excel Object for BELRON Resolved Incident Data
Write-Log -Message ("Going to Create Excel Object for BELRON Daily Incident Resolved  Workbook") -Level Info
$belronDailyIncidentResolvedFileObject = New-Object -ComObject excel.application
$belronDailyIncidentResolvedFileObject.Visible = $true
$belronDailyIncidentResolvedFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening BELRON Daily Incident OUTSTANDING File : " + $belronDailyIncidentResolvedFilePath) -Level Info
$belronDailyIncidentResolvedExcelWorkbook = $belronDailyIncidentResolvedFileObject.Workbooks.Open($belronDailyIncidentResolvedFilePath)
$belronDailyIncidentResolvedExcelWorkbook.activate()
$belronDailyIncidentResolvedExcelWorkSheet = $belronDailyIncidentResolvedExcelWorkbook.Worksheets.Item(1)
$belronDailyIncidentResolvedExcelWorkSheetName = $belronDailyIncidentResolvedExcelWorkSheet.Name
Write-Log -Message ("Selecting BELRON Daily Incident Resolved Worksheet : [" + $belronDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$belronDailyIncidentResolvedExcelWorkSheet.Activate()
$belronDailyIncidentResolvedExcelWorkSheetRange = $belronDailyIncidentResolvedExcelWorkSheet.Range("A:V").CurrentRegion
Write-Log -Message ("Copying Cells from A to V in Worksheet : [" + $belronDailyIncidentResolvedExcelWorkSheetName + "] from " + $belronDailyIncidentResolvedFilePath  + " File." ) -Level Info
$belronDailyIncidentResolvedExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to V in Worksheet : [" + $belronDailyIncidentResolvedExcelWorkSheetName + "] from " + $belronDailyIncidentResolvedFilePath  + " File.") -Level Info

# Copy the BELRON Daily Resolved Raw Data to Template File

Write-Log -Message ("Selecting '" + $belronTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $belronTemplateFilePath) -Level Info

$belronResolvedIncidentCount = $belronOutstandingIncidentCount+2

$belronTemplateRawDataWorkSheetRange = $belronTemplateRawDataWorkSheetRange.Range("A$belronResolvedIncidentCount")
Write-Log -Message ("Copying BELRON Incident Resolved Raw Data from Worksheet [" + $belronDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $belronDailyIncidentResolvedFilePath + " to Worksheet [" + $belronTemplateRawDataWorkSheetName + " of File --> " +  $belronTemplateFilePath) -Level Info
$belronTemplateRawDataWorkSheet.Paste($belronTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied BELRON  Incident Resolved Raw Data from Worksheet [" + $belronDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $belronDailyIncidentResolvedFilePath + " to Worksheet [" + $belronTemplateRawDataWorkSheetName + "] of File --> " +  $belronTemplateFilePath) -Level Info

$beleronResolvedIncidentHeaderdataRange = $belronTemplateRawDataWorkSheet.Cells.Item($belronResolvedIncidentCount,1).EntireRow
$beleronResolvedIncidentHeaderdataRange.Delete()


#Apply Filter the CPA Daily Incident Data in "CPA Data 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of Belero Daily  Data in Worksheet [" + $belronTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$beleronTemplateHurdleExcelWorksheet = $belronTemplateExcelWorkBook.Worksheets.Item($belronTemplateHuddleSheetName)
$beleronTemplateHurdleExcelWorksheet.Activate()
$beleronTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$beleronTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$beleronTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $belronTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$belronTempExcelWorkSheet.Activate()
$belronTempExcelWorkSheet.Range("A1").PasteSpecial(-4163)
$belronTempExcelWorkSheetRange = $belronTempExcelWorkSheet.UsedRange

$TotalBeleronIncidentCount = $belronTemplateFileObject.WorksheetFunction.CountIf($belronTempExcelWorkSheetRange.Range("A1:" + "A" + $belronTempExcelWorkSheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of Beleron Daily Incident Count is : [" + $TotalBeleronIncidentCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $belronTemplateFilePath) -Level Info
$belronTemplateExcelWorkBook.Close($true)
#$cpaRemedyTemplateExcelWorkbook.Save()
$belronTemplateFileObject.Quit()
$belronDailyIncidentResolvedFileObject.Quit()
Write-Log -Message ("**************** PROCESSED BELERON DAILY INCIDENT RAW DATA *************************************") -Level Info


Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING EXOVA DAILY INCIDENT OUTSTANDING RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read Exova Daily Incident OUTSTANDING  Report Data") -Level Info

#Declare File Name, Sheet Name for Exova Daily Incident OUTSTANDING Reports

$ExovaTemplateFilePath =  $parentFolderPath + $ExovaTemplateFileName
$ExovaDailyIncidentOutStandingFilePath =  $RawDataParentFolderPath + $ExovaDailyIncidentOutStandingFileName
$ExovaDailyIncidentResolvedFilePath =  $RawDataParentFolderPath + $ExovaDailyIncidentResolvedFileName
$ExovaDailyIncidentResolvedFilePath
$ExovaTemplateRawDataWorkSheetName = "Exova_Raw"
$ExovaTemplateHuddleSheetName = "Exova_Data 4 Huddle"
$TempworksheetName = "Temp"

Write-Log -Message ("Exova Ticket OUTSTANDING File Path --> " + $ExovaDailyIncidentOutStandingFilePath) -Level Info
Write-Log -Message ("Exova Template File --> " + $ExovaTemplateFilePath) -Level Info

#Declare Excel Object for Exova OUTSTANDING Incident Data
Write-Log -Message ("Going to Create Excel Object for Exova Daily Incident OUTSTANDING  Workbook") -Level Info
$ExovaDailyIncidentOutStandingFileObject = New-Object -ComObject excel.application
$ExovaDailyIncidentOutStandingFileObject.Visible = $true
$ExovaDailyIncidentOutStandingFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening Exova Daily Incident OUTSTANDING File : " + $ExovaDailyIncidentOutStandingFilePath) -Level Info
$ExovaDailyIncidentOutStandingExcelWorkbook = $ExovaDailyIncidentOutStandingFileObject.Workbooks.Open($ExovaDailyIncidentOutStandingFilePath)
$ExovaDailyIncidentOutStandingExcelWorkbook.activate()
$ExovaDailyIncidentOutStandingExcelWorkSheet = $ExovaDailyIncidentOutStandingExcelWorkbook.Worksheets.Item(1)
$ExovaDailyIncidentOutStandingExcelWorkSheetName = $ExovaDailyIncidentOutStandingExcelWorkSheet.Name
Write-Log -Message ("Selecting Exova Daily Incident Worksheet : [" + $ExovaDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$ExovaDailyIncidentOutStandingExcelWorkSheet.Activate()
$ExovaDailyIncidentOutStandingExcelWorkSheetRange = $ExovaDailyIncidentOutStandingExcelWorkSheet.Range("A:V").CurrentRegion
Write-Log -Message ("Copying Cells from A to V in Worksheet : [" + $ExovaDailyIncidentOutStandingExcelWorkSheetName + "] from " + $ExovaDailyIncidentOutStandingFilePath  + " File." ) -Level Info
$ExovaDailyIncidentOutStandingExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to V in Worksheet : [" + $ExovaDailyIncidentOutStandingExcelWorkSheetName + "] from " + $ExovaDailyIncidentOutStandingFilePath  + " File.") -Level Info

# Copy the Exova Daily Remedy Raw Data to Template File

Write-Log -Message ("Going to Open Exova Template File") -Level Info
$ExovaTemplateFileObject = New-Object -ComObject excel.application
$ExovaTemplateFileObject.Visible = $true
$ExovaTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening Exova Template File --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTemplateExcelWorkBook = $ExovaTemplateFileObject.Workbooks.Open($ExovaTemplateFilePath)
$ExovaTemplateExcelWorkBook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $ExovaTemplateExcelWorkBook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $ExovaTemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



Write-Log -Message ("Selecting Exova Template Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTempExcelWorkSheet = $ExovaTemplateExcelWorkBook.Worksheets.Item($TempworksheetName)
$ExovaTempExcelWorkSheet.Activate()
$ExovaTempExcelWorkSheetRange = $ExovaTempExcelWorkSheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $ExovaTemplateFilePath) -Level Info

$ExovaTempExcelWorkSheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $ExovaTemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTempExcelWorkSheet = $ExovaTemplateExcelWorkBook.Worksheets.Add()
$ExovaTempExcelWorkSheet.Name = $TempworksheetName
$ExovaTempExcelWorkSheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $ExovaTemplateFilePath) -Level Info

}

Write-Log -Message ("Selecting '" + $ExovaTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $ExovaTemplateFilePath) -Level Info

$ExovaTemplateRawDataWorkSheet = $ExovaTemplateExcelWorkBook.Worksheets.Item($ExovaTemplateRawDataWorkSheetName)
$ExovaTemplateRawDataWorkSheet.Activate()
#Delete Old Records from Sheet "Exova RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To V in '" + $ExovaTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTemplateRawDataWorkSheetRange = $ExovaTemplateRawDataWorkSheet.Range("A:V")
$ExovaTemplateRawDataWorkSheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To V in '" + $ExovaTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTemplateRawDataWorkSheetRange = $ExovaTemplateRawDataWorkSheetRange.Range("A1")
Write-Log -Message ("Copying Exova Incident OUTSTANDING Raw Data from Worksheet [" + $ExovaDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $ExovaDailyIncidentOutStandingFilePath + " to Worksheet [" + $ExovaTemplateRawDataWorkSheetName + " of File --> " +  $ExovaTemplateFilePath) -Level Info
$ExovaTemplateRawDataWorkSheet.Paste($ExovaTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied Exova  Incident OUTSTANDING Raw Data from Worksheet [" + $ExovaDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $ExovaDailyIncidentOutStandingFilePath + " to Worksheet [" + $ExovaTemplateRawDataWorkSheetName + "] of File --> " +  $ExovaTemplateFilePath) -Level Info

$ExovaDailyIncidentOutStandingFileObject.Quit()
$ExovaTemplateRawDataWorkSheetUsedRange = $ExovaTemplateRawDataWorkSheet.UsedRange

$ExovaOutstandingIncidentCount = $ExovaTemplateFileObject.WorksheetFunction.CountIf($ExovaTemplateRawDataWorkSheetUsedRange.Range("A1:" + "A" + $ExovaTemplateRawDataWorkSheetUsedRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of Exova Daily Incident OUTSTANDING Count is : [" + $ExovaOutstandingIncidentCount  + "]") -Level Info



#Declare Excel Object for Exova Resolved Incident Data
Write-Log -Message ("Going to Create Excel Object for Exova Daily Incident Resolved  Workbook") -Level Info
$ExovaDailyIncidentResolvedFileObject = New-Object -ComObject excel.application
$ExovaDailyIncidentResolvedFileObject.Visible = $true
$ExovaDailyIncidentResolvedFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening Exova Daily Incident OUTSTANDING File : " + $ExovaDailyIncidentResolvedFilePath) -Level Info
$ExovaDailyIncidentResolvedExcelWorkbook = $ExovaDailyIncidentResolvedFileObject.Workbooks.Open($ExovaDailyIncidentResolvedFilePath)
$ExovaDailyIncidentResolvedExcelWorkbook.activate()
$ExovaDailyIncidentResolvedExcelWorkSheet = $ExovaDailyIncidentResolvedExcelWorkbook.Worksheets.Item(1)
$ExovaDailyIncidentResolvedExcelWorkSheetName = $ExovaDailyIncidentResolvedExcelWorkSheet.Name
Write-Log -Message ("Selecting Exova Daily Incident Resolved Worksheet : [" + $ExovaDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$ExovaDailyIncidentResolvedExcelWorkSheet.Activate()
$ExovaDailyIncidentResolvedExcelWorkSheetRange = $ExovaDailyIncidentResolvedExcelWorkSheet.Range("A:V").CurrentRegion
Write-Log -Message ("Copying Cells from A to V in Worksheet : [" + $ExovaDailyIncidentResolvedExcelWorkSheetName + "] from " + $ExovaDailyIncidentResolvedFilePath  + " File." ) -Level Info
$ExovaDailyIncidentResolvedExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to V in Worksheet : [" + $ExovaDailyIncidentResolvedExcelWorkSheetName + "] from " + $ExovaDailyIncidentResolvedFilePath  + " File.") -Level Info

# Copy the Exova Daily Resolved Raw Data to Template File

Write-Log -Message ("Selecting '" + $ExovaTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $ExovaTemplateFilePath) -Level Info

$ExovaResolvedIncidentCount = $ExovaOutstandingIncidentCount+2

$ExovaTemplateRawDataWorkSheetRange = $ExovaTemplateRawDataWorkSheetRange.Range("A$ExovaResolvedIncidentCount")
Write-Log -Message ("Copying Exova Incident Resolved Raw Data from Worksheet [" + $ExovaDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $ExovaDailyIncidentResolvedFilePath + " to Worksheet [" + $ExovaTemplateRawDataWorkSheetName + " of File --> " +  $ExovaTemplateFilePath) -Level Info
$ExovaTemplateRawDataWorkSheet.Paste($ExovaTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied Exova  Incident Resolved Raw Data from Worksheet [" + $ExovaDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $ExovaDailyIncidentResolvedFilePath + " to Worksheet [" + $ExovaTemplateRawDataWorkSheetName + "] of File --> " +  $ExovaTemplateFilePath) -Level Info

$ExovaResolvedIncidentHeaderdataRange = $ExovaTemplateRawDataWorkSheet.Cells.Item($ExovaResolvedIncidentCount,1).EntireRow
$ExovaResolvedIncidentHeaderdataRange.Delete()


#Apply Filter the CPA Daily Incident Data in "CPA Data 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of Belero Daily  Data in Worksheet [" + $ExovaTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$ExovaTemplateHurdleExcelWorksheet = $ExovaTemplateExcelWorkBook.Worksheets.Item($ExovaTemplateHuddleSheetName)
$ExovaTemplateHurdleExcelWorksheet.Activate()
$ExovaTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$ExovaTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$ExovaTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $ExovaTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$ExovaTempExcelWorkSheet.Activate()
$ExovaTempExcelWorkSheet.Range("A1").PasteSpecial(-4163)
$ExovaTempExcelWorkSheetRange = $ExovaTempExcelWorkSheet.UsedRange

$TotalExovaIncidentCount = $ExovaTemplateFileObject.WorksheetFunction.CountIf($ExovaTempExcelWorkSheetRange.Range("A1:" + "A" + $ExovaTempExcelWorkSheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of Exova Daily Incident Count is : [" + $TotalExovaIncidentCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTemplateExcelWorkBook.Close($true)
#$cpaRemedyTemplateExcelWorkbook.Save()
$ExovaTemplateFileObject.Quit()
$ExovaDailyIncidentResolvedFileObject.Quit()
Write-Log -Message ("**************** PROCESSED Exova DAILY INCIDENT RAW DATA *************************************") -Level Info




Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING NATIONAL GRID US DAILY INCIDENT OUTSTANDING RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read NATIONAL GRID US Daily Incident OUTSTANDING  Report Data") -Level Info

#Declare File Name, Sheet Name for ngrid Daily Incident OUTSTANDING Reports

$ngridTemplateFilePath =  $parentFolderPath + $ngridTemplateFileName
$ngridDailyIncidentOutStandingFilePath =  $RawDataParentFolderPath + $ngridDailyIncidentOutStandingFileName
$ngridDailyIncidentResolvedFilePath =  $RawDataParentFolderPath + $ngridDailyIncidentResolvedFileName
$ngridDailyIncidentResolvedFilePath
$ngridTemplateRawDataWorkSheetName = "NG US RAW"
$ngridTemplateHuddleSheetName = "NG US Data 4 Huddle"
$TempworksheetName = "Temp"

Write-Log -Message ("NATIONAL GRID Ticket OUTSTANDING File Path --> " + $ngridDailyIncidentOutStandingFilePath) -Level Info
Write-Log -Message ("NATIONAL GRID Template File --> " + $ngridTemplateFilePath) -Level Info

#Declare Excel Object for ngrid OUTSTANDING Incident Data
Write-Log -Message ("Going to Create Excel Object for NATIONAL GRID Daily Incident OUTSTANDING  Workbook") -Level Info
$ngridDailyIncidentOutStandingFileObject = New-Object -ComObject excel.application
$ngridDailyIncidentOutStandingFileObject.Visible = $true
$ngridDailyIncidentOutStandingFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NATIONAL GRID Daily Incident OUTSTANDING File : " + $ngridDailyIncidentOutStandingFilePath) -Level Info
$ngridDailyIncidentOutStandingExcelWorkbook = $ngridDailyIncidentOutStandingFileObject.Workbooks.Open($ngridDailyIncidentOutStandingFilePath)
$ngridDailyIncidentOutStandingExcelWorkbook.activate()
$ngridDailyIncidentOutStandingExcelWorkSheet = $ngridDailyIncidentOutStandingExcelWorkbook.Worksheets.Item(1)
$ngridDailyIncidentOutStandingExcelWorkSheetName = $ngridDailyIncidentOutStandingExcelWorkSheet.Name
Write-Log -Message ("Selecting NATIONAL GRID Daily Incident Worksheet : [" + $ngridDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$ngridDailyIncidentOutStandingExcelWorkSheet.Activate()
$ngridDailyIncidentOutStandingExcelWorkSheetRange = $ngridDailyIncidentOutStandingExcelWorkSheet.Range("A:AB").CurrentRegion
Write-Log -Message ("Copying Cells from A to AB in Worksheet : [" + $ngridDailyIncidentOutStandingExcelWorkSheetName + "] from " + $ngridDailyIncidentOutStandingFilePath  + " File." ) -Level Info
$ngridDailyIncidentOutStandingExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to AB in Worksheet : [" + $ngridDailyIncidentOutStandingExcelWorkSheetName + "] from " + $ngridDailyIncidentOutStandingFilePath  + " File.") -Level Info

# Copy the ngrid Daily Remedy Raw Data to Template File

Write-Log -Message ("Going to Open NATIONAL GRID Template File") -Level Info
$ngridTemplateFileObject = New-Object -ComObject excel.application
$ngridTemplateFileObject.Visible = $true
$ngridTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening ngrid Template File --> " + $ngridTemplateFilePath) -Level Info
$ngridTemplateExcelWorkBook = $ngridTemplateFileObject.Workbooks.Open($ngridTemplateFilePath)
$ngridTemplateExcelWorkBook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $ngridTemplateExcelWorkBook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $ngridTemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



Write-Log -Message ("Selecting ngrid Template Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $ngridTemplateFilePath) -Level Info
$ngridTempExcelWorkSheet = $ngridTemplateExcelWorkBook.Worksheets.Item($TempworksheetName)
$ngridTempExcelWorkSheet.Activate()
$ngridTempExcelWorkSheetRange = $ngridTempExcelWorkSheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $ngridTemplateFilePath) -Level Info

$ngridTempExcelWorkSheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $ngridTemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $ngridTemplateFilePath) -Level Info
$ngridTempExcelWorkSheet = $ngridTemplateExcelWorkBook.Worksheets.Add()
$ngridTempExcelWorkSheet.Name = $TempworksheetName
$ngridTempExcelWorkSheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $ngridTemplateFilePath) -Level Info

}

Write-Log -Message ("Selecting '" + $ngridTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $ngridTemplateFilePath) -Level Info

$ngridTemplateRawDataWorkSheet = $ngridTemplateExcelWorkBook.Worksheets.Item($ngridTemplateRawDataWorkSheetName)
$ngridTemplateRawDataWorkSheet.Activate()
#Delete Old Records from Sheet "ngrid RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To V in '" + $ngridTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $ngridTemplateFilePath) -Level Info
$ngridTemplateRawDataWorkSheetRange = $ngridTemplateRawDataWorkSheet.Range("A:AB")
$ngridTemplateRawDataWorkSheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To V in '" + $ngridTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $ngridTemplateFilePath) -Level Info
$ngridTemplateRawDataWorkSheetRange = $ngridTemplateRawDataWorkSheetRange.Range("A1")
Write-Log -Message ("Copying ngrid Incident OUTSTANDING Raw Data from Worksheet [" + $ngridDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $ngridDailyIncidentOutStandingFilePath + " to Worksheet [" + $ngridTemplateRawDataWorkSheetName + " of File --> " +  $ngridTemplateFilePath) -Level Info
$ngridTemplateRawDataWorkSheet.Paste($ngridTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied ngrid  Incident OUTSTANDING Raw Data from Worksheet [" + $ngridDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $ngridDailyIncidentOutStandingFilePath + " to Worksheet [" + $ngridTemplateRawDataWorkSheetName + "] of File --> " +  $ngridTemplateFilePath) -Level Info

$ngridDailyIncidentOutStandingFileObject.Quit()
$ngridTemplateRawDataWorkSheetUsedRange = $ngridTemplateRawDataWorkSheet.UsedRange

$ngridOutstandingIncidentCount = $ngridTemplateFileObject.WorksheetFunction.CountIf($ngridTemplateRawDataWorkSheetUsedRange.Range("A1:" + "A" + $ngridTemplateRawDataWorkSheetUsedRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of NATIONAL GRID US Daily Incident OUTSTANDING Count is : [" + $ngridOutstandingIncidentCount  + "]") -Level Info



#Declare Excel Object for ngrid Resolved Incident Data
Write-Log -Message ("Going to Create Excel Object for ngrid Daily Incident Resolved  Workbook") -Level Info
$ngridDailyIncidentResolvedFileObject = New-Object -ComObject excel.application
$ngridDailyIncidentResolvedFileObject.Visible = $true
$ngridDailyIncidentResolvedFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening ngrid Daily Incident OUTSTANDING File : " + $ngridDailyIncidentResolvedFilePath) -Level Info
$ngridDailyIncidentResolvedExcelWorkbook = $ngridDailyIncidentResolvedFileObject.Workbooks.Open($ngridDailyIncidentResolvedFilePath)
$ngridDailyIncidentResolvedExcelWorkbook.activate()
$ngridDailyIncidentResolvedExcelWorkSheet = $ngridDailyIncidentResolvedExcelWorkbook.Worksheets.Item(1)
$ngridDailyIncidentResolvedExcelWorkSheetName = $ngridDailyIncidentResolvedExcelWorkSheet.Name
Write-Log -Message ("Selecting ngrid Daily Incident Resolved Worksheet : [" + $ngridDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$ngridDailyIncidentResolvedExcelWorkSheet.Activate()
$ngridDailyIncidentResolvedExcelWorkSheetRange = $ngridDailyIncidentResolvedExcelWorkSheet.Range("A:AB").CurrentRegion
Write-Log -Message ("Copying Cells from A to AB in Worksheet : [" + $ngridDailyIncidentResolvedExcelWorkSheetName + "] from " + $ngridDailyIncidentResolvedFilePath  + " File." ) -Level Info
$ngridDailyIncidentResolvedExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to AB in Worksheet : [" + $ngridDailyIncidentResolvedExcelWorkSheetName + "] from " + $ngridDailyIncidentResolvedFilePath  + " File.") -Level Info

# Copy the ngrid Daily Resolved Raw Data to Template File

Write-Log -Message ("Selecting '" + $ngridTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $ngridTemplateFilePath) -Level Info

$ngridResolvedIncidentCount = $ngridOutstandingIncidentCount+2

$ngridTemplateRawDataWorkSheetRange = $ngridTemplateRawDataWorkSheetRange.Range("A$ngridResolvedIncidentCount")
Write-Log -Message ("Copying ngrid Incident Resolved Raw Data from Worksheet [" + $ngridDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $ngridDailyIncidentResolvedFilePath + " to Worksheet [" + $ngridTemplateRawDataWorkSheetName + " of File --> " +  $ngridTemplateFilePath) -Level Info
$ngridTemplateRawDataWorkSheet.Paste($ngridTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied ngrid  Incident Resolved Raw Data from Worksheet [" + $ngridDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $ngridDailyIncidentResolvedFilePath + " to Worksheet [" + $ngridTemplateRawDataWorkSheetName + "] of File --> " +  $ngridTemplateFilePath) -Level Info

$ngridResolvedIncidentHeaderdataRange = $ngridTemplateRawDataWorkSheet.Cells.Item($ngridResolvedIncidentCount,1).EntireRow
$ngridResolvedIncidentHeaderdataRange.Delete()


#Apply Filter the CPA Daily Incident Data in "NATIONAL GRID 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of NATIONAL GRID - US Daily  Data in Worksheet [" + $ngridTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$ngridTemplateHurdleExcelWorksheet = $ngridTemplateExcelWorkBook.Worksheets.Item($ngridTemplateHuddleSheetName)
$ngridTemplateHurdleExcelWorksheet.Activate()
$ngridTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$ngridTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$ngridTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $ngridTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$ngridTempExcelWorkSheet.Activate()
$ngridTempExcelWorkSheet.Range("A1").PasteSpecial(-4163)
$ngridTempExcelWorkSheetRange = $ngridTempExcelWorkSheet.UsedRange

$TotalngridIncidentCount = $ngridTemplateFileObject.WorksheetFunction.CountIf($ngridTempExcelWorkSheetRange.Range("A1:" + "A" + $ngridTempExcelWorkSheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of ngrid Daily Incident Count is : [" + $TotalngridIncidentCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $ngridTemplateFilePath) -Level Info
$ngridTemplateExcelWorkBook.Close($true)
#$cpaRemedyTemplateExcelWorkbook.Save()
$ngridTemplateFileObject.Quit()
$ngridDailyIncidentResolvedFileObject.Quit()

Write-Log -Message ("**************** PROCESSED NATIONAL GRID US DAILY INCIDENT RAW DATA *************************************") -Level Info



Write-Log -Message ("                                                                                                   ") -Level Info
Write-Log -Message ("**************** PROCESSING NATIONAL GRID UK DAILY INCIDENT OUTSTANDING RAW DATA *************************************") -Level Info
Write-Log -Message ("Going to Read NATIONAL GRID UK Daily Incident OUTSTANDING  Report Data") -Level Info

#Declare File Name, Sheet Name for NATIONAL GRID Daily Incident OUTSTANDING Reports

$ngridukTemplateFilePath =  $parentFolderPath + $ngridukTemplateFileName
$ngridukDailyIncidentOutStandingFilePath =  $RawDataParentFolderPath + $ngridukDailyIncidentOutStandingFileName
$ngridukDailyIncidentResolvedFilePath =  $RawDataParentFolderPath + $ngridukDailyIncidentResolvedFileName
$ngridukDailyIncidentResolvedFilePath
$ngridukTemplateRawDataWorkSheetName = "NG UK RAW"
$ngridukTemplateHuddleSheetName = "NG UK Data 4 Huddle"
$TempworksheetName = "Temp"

Write-Log -Message ("NATIONAL GRID UK Ticket OUTSTANDING File Path --> " + $ngridukDailyIncidentOutStandingFilePath) -Level Info
Write-Log -Message ("NATIONAL GRID UK Template File --> " + $ngridukTemplateFilePath) -Level Info

#Declare Excel Object for ngriduk OUTSTANDING Incident Data
Write-Log -Message ("Going to Create Excel Object for NATIONAL GRID UK Daily Incident OUTSTANDING  Workbook") -Level Info
$ngridukDailyIncidentOutStandingFileObject = New-Object -ComObject excel.application
$ngridukDailyIncidentOutStandingFileObject.Visible = $true
$ngridukDailyIncidentOutStandingFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NATIONAL GRID UK Daily Incident OUTSTANDING File : " + $ngridukDailyIncidentOutStandingFilePath) -Level Info
$ngridukDailyIncidentOutStandingExcelWorkbook = $ngridukDailyIncidentOutStandingFileObject.Workbooks.Open($ngridukDailyIncidentOutStandingFilePath)
$ngridukDailyIncidentOutStandingExcelWorkbook.activate()
$ngridukDailyIncidentOutStandingExcelWorkSheet = $ngridukDailyIncidentOutStandingExcelWorkbook.Worksheets.Item(1)
$ngridukDailyIncidentOutStandingExcelWorkSheetName = $ngridukDailyIncidentOutStandingExcelWorkSheet.Name
Write-Log -Message ("Selecting NATIONAL GRID UK Daily Incident Worksheet : [" + $ngridukDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$ngridukDailyIncidentOutStandingExcelWorkSheet.Activate()
$ngridukDailyIncidentOutStandingExcelWorkSheetRange = $ngridukDailyIncidentOutStandingExcelWorkSheet.Range("A:AB").CurrentRegion
Write-Log -Message ("Copying Cells from A to AB in Worksheet : [" + $ngridukDailyIncidentOutStandingExcelWorkSheetName + "] from " + $ngridukDailyIncidentOutStandingFilePath  + " File." ) -Level Info
$ngridukDailyIncidentOutStandingExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to AB in Worksheet : [" + $ngridukDailyIncidentOutStandingExcelWorkSheetName + "] from " + $ngridukDailyIncidentOutStandingFilePath  + " File.") -Level Info

# Copy the ngriduk Daily Remedy Raw Data to Template File

Write-Log -Message ("Going to Open NATIONAL GRID UK Template File") -Level Info
$ngridukTemplateFileObject = New-Object -ComObject excel.application
$ngridukTemplateFileObject.Visible = $true
$ngridukTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NATIONAL GRID UK  Template File --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTemplateExcelWorkBook = $ngridukTemplateFileObject.Workbooks.Open($ngridukTemplateFilePath)
$ngridukTemplateExcelWorkBook.activate()

# Check if Temp Sheet is already created or not. If not then create new Temp Sheet

$WorkSheets = $ngridukTemplateExcelWorkBook.WorkSheets
$flag=$false

Write-Log -Message ("Check if 'Temp' WorkSheet exists in file --> " + $ngridukTemplateFilePath) -Level Info
     

  foreach ($WorkSheet in $WorkSheets) {
    
     If ($WorkSheet.Name -eq $TempworksheetName){
     $flag = $true
     }

     }



Write-Log -Message ("Selecting NATIONAL GRID UK Template Worksheet : [" + $TempworksheetName + "]") -Level Info

if ($flag -eq $true)
{
Write-Log -Message ("'Temp' WorkSheet exists in file --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTempExcelWorkSheet = $ngridukTemplateExcelWorkBook.Worksheets.Item($TempworksheetName)
$ngridukTempExcelWorkSheet.Activate()
$ngridukTempExcelWorkSheetRange = $ngridukTempExcelWorkSheet.Range("A:I")
Write-Log -Message ("Remove All data from 'Temp' WorkSheet in File --> " + $ngridukTemplateFilePath) -Level Info

$ngridukTempExcelWorkSheetRange.clear()

}
else
{

Write-Log -Message ("'Temp' WorkSheet do not exists in file --> " + $ngridukTemplateFilePath) -Level Info
Write-Log -Message ("Going to create 'Temp' WorkSheet in file --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTempExcelWorkSheet = $ngridukTemplateExcelWorkBook.Worksheets.Add()
$ngridukTempExcelWorkSheet.Name = $TempworksheetName
$ngridukTempExcelWorkSheet.Activate()
Write-Log -Message ("'Temp' WorkSheet Successfully Created in file --> " + $ngridukTemplateFilePath) -Level Info

}

Write-Log -Message ("Selecting '" + $ngridukTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $ngridukTemplateFilePath) -Level Info

$ngridukTemplateRawDataWorkSheet = $ngridukTemplateExcelWorkBook.Worksheets.Item($ngridukTemplateRawDataWorkSheetName)
$ngridukTemplateRawDataWorkSheet.Activate()
#Delete Old Records from Sheet "ngriduk RAW DAta"
Write-Log -Message ("Deleting Old Raw Data from Column A To V in '" + $ngridukTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTemplateRawDataWorkSheetRange = $ngridukTemplateRawDataWorkSheet.Range("A:AB")
$ngridukTemplateRawDataWorkSheetRange.clear()
Write-Log -Message ("Deleted Old Raw Data from Column A To V in '" + $ngridukTemplateRawDataWorkSheetName + "' WorkSheet from file --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTemplateRawDataWorkSheetRange = $ngridukTemplateRawDataWorkSheetRange.Range("A1")
Write-Log -Message ("Copying ngriduk Incident OUTSTANDING Raw Data from Worksheet [" + $ngridukDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $ngridukDailyIncidentOutStandingFilePath + " to Worksheet [" + $ngridukTemplateRawDataWorkSheetName + " of File --> " +  $ngridukTemplateFilePath) -Level Info
$ngridukTemplateRawDataWorkSheet.Paste($ngridukTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied ngriduk  Incident OUTSTANDING Raw Data from Worksheet [" + $ngridukDailyIncidentOutStandingExcelWorkSheetName + "] of File --> " + $ngridukDailyIncidentOutStandingFilePath + " to Worksheet [" + $ngridukTemplateRawDataWorkSheetName + "] of File --> " +  $ngridukTemplateFilePath) -Level Info

$ngridukDailyIncidentOutStandingFileObject.Quit()
$ngridukTemplateRawDataWorkSheetUsedRange = $ngridukTemplateRawDataWorkSheet.UsedRange

$ngridukOutstandingIncidentCount = $ngridukTemplateFileObject.WorksheetFunction.CountIf($ngridukTemplateRawDataWorkSheetUsedRange.Range("A1:" + "A" + $ngridukTemplateRawDataWorkSheetUsedRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of NATIONAL GRID UK Daily Incident OUTSTANDING Count is : [" + $ngridukOutstandingIncidentCount  + "]") -Level Info



#Declare Excel Object for ngriduk Resolved Incident Data
Write-Log -Message ("Going to Create Excel Object for NATIONAL GRID UK Daily Incident Resolved  Workbook") -Level Info
$ngridukDailyIncidentResolvedFileObject = New-Object -ComObject excel.application
$ngridukDailyIncidentResolvedFileObject.Visible = $true
$ngridukDailyIncidentResolvedFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening ngriduk Daily Incident OUTSTANDING File : " + $ngridukDailyIncidentResolvedFilePath) -Level Info
$ngridukDailyIncidentResolvedExcelWorkbook = $ngridukDailyIncidentResolvedFileObject.Workbooks.Open($ngridukDailyIncidentResolvedFilePath)
$ngridukDailyIncidentResolvedExcelWorkbook.activate()
$ngridukDailyIncidentResolvedExcelWorkSheet = $ngridukDailyIncidentResolvedExcelWorkbook.Worksheets.Item(1)
$ngridukDailyIncidentResolvedExcelWorkSheetName = $ngridukDailyIncidentResolvedExcelWorkSheet.Name
Write-Log -Message ("Selecting ngriduk Daily Incident Resolved Worksheet : [" + $ngridukDailyIncidentOutStandingExcelWorkSheetName + "]") -Level Info

$ngridukDailyIncidentResolvedExcelWorkSheet.Activate()
$ngridukDailyIncidentResolvedExcelWorkSheetRange = $ngridukDailyIncidentResolvedExcelWorkSheet.Range("A:AB").CurrentRegion
Write-Log -Message ("Copying Cells from A to AB in Worksheet : [" + $ngridukDailyIncidentResolvedExcelWorkSheetName + "] from " + $ngridukDailyIncidentResolvedFilePath  + " File." ) -Level Info
$ngridukDailyIncidentResolvedExcelWorkSheetRange.copy()
Write-Log -Message ("Copied Cells from A to AB in Worksheet : [" + $ngridukDailyIncidentResolvedExcelWorkSheetName + "] from " + $ngridukDailyIncidentResolvedFilePath  + " File.") -Level Info

# Copy the ngriduk Daily Resolved Raw Data to Template File

Write-Log -Message ("Selecting '" + $ngridukTemplateRawDataWorkSheetName + "' WorkSheet in file --> " + $ngridukTemplateFilePath) -Level Info

$ngridukResolvedIncidentCount = $ngridukOutstandingIncidentCount+2

$ngridukTemplateRawDataWorkSheetRange = $ngridukTemplateRawDataWorkSheetRange.Range("A$ngridukResolvedIncidentCount")
Write-Log -Message ("Copying ngriduk Incident Resolved Raw Data from Worksheet [" + $ngridukDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $ngridukDailyIncidentResolvedFilePath + " to Worksheet [" + $ngridukTemplateRawDataWorkSheetName + " of File --> " +  $ngridukTemplateFilePath) -Level Info
$ngridukTemplateRawDataWorkSheet.Paste($ngridukTemplateRawDataWorkSheetRange)
Write-Log -Message ("Copied ngriduk  Incident Resolved Raw Data from Worksheet [" + $ngridukDailyIncidentResolvedExcelWorkSheetName + "] of File --> " + $ngridukDailyIncidentResolvedFilePath + " to Worksheet [" + $ngridukTemplateRawDataWorkSheetName + "] of File --> " +  $ngridukTemplateFilePath) -Level Info

$ngridukResolvedIncidentHeaderdataRange = $ngridukTemplateRawDataWorkSheet.Cells.Item($ngridukResolvedIncidentCount,1).EntireRow
$ngridukResolvedIncidentHeaderdataRange.Delete()


#Apply Filter the CPA Daily Incident Data in "NATIONAL GRID 4 Hurdle" Sheet based on Column J - "Queue Check" with condition as "1"
Write-Log -Message ("Apply Filter of NATIONAL GRID UK Daily  Data in Worksheet [" + $ngridukTemplateHuddleSheetName  + "] based on Column J - Queue Check with condition as 1") -Level Info

$ngridukTemplateHurdleExcelWorksheet = $ngridukTemplateExcelWorkBook.Worksheets.Item($ngridukTemplateHuddleSheetName)
$ngridukTemplateHurdleExcelWorksheet.Activate()
$ngridukTemplateHurdleExcelWorksheet.Range("A:J").AutoFilter(10, "1")
$ngridukTemplateHurdleExcelWorksheet.Range("A:I").Select

Write-Log -Message ("Copy Filtered Data from Range A:I in Worksheet [" + $TempworksheetName  + "]") -Level Info
$ngridukTemplateHurdleExcelWorksheet.Range("A:I").copy() | out-null
Write-Log -Message ("Copied Filtered Data from Range A:I from Worksheet [" + $ngridukTemplateHuddleSheetName  + "] to Worksheet [" + $TempworksheetName + "]") -Level Info
$ngridukTempExcelWorkSheet.Activate()
$ngridukTempExcelWorkSheet.Range("A1").PasteSpecial(-4163)
$ngridukTempExcelWorkSheetRange = $ngridukTempExcelWorkSheet.UsedRange

$TotalngridukIncidentCount = $ngridukTemplateFileObject.WorksheetFunction.CountIf($ngridukTempExcelWorkSheetRange.Range("A1:" + "A" + $ngridukTempExcelWorkSheetRange.Rows.Count), "<>") - 1


Write-Log -Message ("Total Number of NATIONAL GRID UK  Daily Incident Count is : [" + $TotalngridukIncidentCount  + "]") -Level Info

Write-Log -Message (" Saving and Closing File --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTemplateExcelWorkBook.Close($true)
$cpaRemedyTemplateExcelWorkbook.Save()
$ngridukTemplateFileObject.Quit()
$ngridukDailyIncidentResolvedFileObject.Quit()
Write-Log -Message ("**************** PROCESSED NATIONAL GRID UK DAILY INCIDENT RAW DATA *************************************") -Level Info





Write-Log -Message ("**************** PROCESSING UKNCR HUB HURDLE EXCEL FILE *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info


Write-Log -Message ("Going to Open UKI NCR HUB Dashboard Excel File") -Level Info
Write-Log -Message ("Going to Create Excel Object for UKI NCR HUB Dashboard Excel Workbook") -Level Info

$ukNCRHUBExcelRawDataWorkSheetName ="Raw Data"
$TempworksheetName = "Temp"
$ukNCRHUBExcelFilePath =  $parentFolderPath + $ukNCRHUBExcelFileName

Write-Log -Message ("UK NCR HUB Excel File Path --> " + $ukNCRHUBExcelFilePath) -Level Info


$ukNCRHUBExcelFileObject = New-Object -ComObject excel.application
$ukNCRHUBExcelFileObject.Visible = $true
$ukNCRHUBExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening UKI NCR HUB Dashboard File --> " + $ukNCRHUBExcelFilePath) -Level Info
$ukNCRHUBExcelWorkBook = $ukNCRHUBExcelFileObject.Workbooks.Open($ukNCRHUBExcelFilePath)
$ukNCRHUBExcelWorkBook.activate()


Write-Log -Message ("Selecting the worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName+"] from UK NCR Hub Dashboard File." ) -Level Info
$ukNCRHUBExcelWorkSheet = $ukNCRHUBExcelWorkBook.Worksheets.Item($ukNCRHUBExcelRawDataWorkSheetName)
$ukNCRHUBExcelWorkSheet.Activate()
$ukNCRHUBExcelWorkSheet.AutoFilterMode=$false
Write-Log -Message ("Selecting the Raw data in worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName+"] from UKI NCR HUB Template File." ) -Level Info
$ukNCRHUBExcelWorkSheetDeleteRange = $ukNCRHUBExcelWorkSheet.Range("A:I")
Write-Log -Message ("Deleting Old Raw data in worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName+"] from UKI NCR HUB Template File." ) -Level Info
$ukNCRHUBExcelWorkSheetDeleteRange.clear()

Write-Log -Message ("**************** WRITING NETWORK RAIL CONSOLIDATED INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info


Write-Log -Message ("Going to Open NWR Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for NWR Template Hurdle Workbook") -Level Info
$nwrTemplateExcelFileObject = New-Object -ComObject excel.application
$nwrTemplateExcelFileObject.Visible = $true
$nwrTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening nwr Template File --> " + $nwrtemplateFilePath) -Level Info
$nwrTemplateExcelWorkbook = $nwrTemplateExcelFileObject.Workbooks.Open($nwrtemplateFilePath)
$nwrTemplateExcelWorkbook.activate()
Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from nwr Template File." ) -Level Info
$nwrTempExcelWorkSheet = $nwrTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$nwrTempExcelWorkSheet.Activate()
#Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from nwr JIRA Template File." ) -Level Info
#$nwrTempExcelWorkSheetHeaderRowRange = $nwrTempExcelWorkSheet.Cells.Item(1,1).EntireRow
#$nwrTempExcelWorkSheetHeaderRowRange.Delete()
$nwrTempExcelWorkSheetRange = $nwrTempExcelWorkSheet.Range("A:I")
#$cpaJIRADailyIncidentRowCount = $cpaJIRATempdataExcelWorksheetRange.Rows.Count
Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from nwr JIRA Template File." ) -Level Info
$nwrTempExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from Network Rail Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info
$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A1"))
$nwrTotalIncidentCount = $nwrTemplateExcelFileObject.WorksheetFunction.CountIf($nwrTempExcelWorkSheet.Range("A1:" + "A" + $nwrTempExcelWorkSheet.Rows.Count), "<>") 
Write-Log -Message ("Total Number of Incidents Of 'Network Rail Incidents in Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File are [" + $nwrTotalIncidentCount  + "]") -Level Info
Write-Log -Message ("Closing File --> " + $nwrtemplateFilePath) -Level Info
$nwrTemplateExcelFileObject.Quit()
Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA NETWORK RAIL DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info




Write-Log -Message ("**************** WRITING XChanging CONSOLIDATED INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

Write-Log -Message ("Going to Open xchanging Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for XChanging Template Hurdle Workbook") -Level Info
$xchangingTemplateExcelFileObject = New-Object -ComObject excel.application
$xchangingTemplateExcelFileObject.Visible = $true
$xchangingTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening xchanging Template File --> " + $xchangingtemplateFilePath) -Level Info
$xchangingTemplateExcelWorkbook = $xchangingTemplateExcelFileObject.Workbooks.Open($xchangingtemplateFilePath)
$xchangingTemplateExcelWorkbook.activate()

Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from nwr Template File." ) -Level Info
$xchangingTemplateExcelWorkSheet = $xchangingTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$xchangingTemplateExcelWorkSheet.Activate()
Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from nwr JIRA Template File." ) -Level Info
$xchangingTemplateExcelWorkSheetHeaderRange = $xchangingTemplateExcelWorkSheet.Cells.Item(1,1).EntireRow
$xchangingTemplateExcelWorkSheetHeaderRange.Delete()
$xchangingTemplateExcelWorkSheetRange = $xchangingTemplateExcelWorkSheet.Range("A:I")
#$cpaJIRADailyIncidentRowCount = $cpaJIRATempdataExcelWorksheetRange.Rows.Count
Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from nwr JIRA Template File." ) -Level Info
$xchangingTemplateExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from XChanging Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info
$nwrTotalIncidentCount
$TotalIncidentCount = $nwrTotalIncidentCount + 1

$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$xChangingTotalIncidentCount = $xchangingTemplateExcelFileObject.WorksheetFunction.CountIf($xchangingTemplateExcelWorkSheet.Range("A1:" + "A" + $xchangingTemplateExcelWorkSheet.Rows.Count), "<>") 

$TotalIncidentCount = $TotalIncidentCount + $xChangingTotalIncidentCount
Write-Log -Message ("Total Number of Incidents Of 'Xchanging Incidents in Worksheet [" + $TempworksheetName + "] of UKI NCR HUB Template File are [" + $xChangingTotalIncidentCount  + "]") -Level Info
Write-Log -Message ("Closing File --> " + $xchangingtemplateFilePath) -Level Info
$xchangingTemplateExcelFileObject.Quit()


Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA XCHANGING DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info



Write-Log -Message ("**************** WRITING CPA REMEDY INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

Write-Log -Message ("Going to Open CPA Remedy Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for CPA Remedy Template Hurdle Workbook") -Level Info
Write-Log -Message ("Going to Open CPA Remedy Template File") -Level Info
$cpaRemedyTemplateExcelFileObject = New-Object -ComObject excel.application
$cpaRemedyTemplateExcelFileObject.Visible = $true
$cpaRemedyTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening CPA Remedy Template File --> " + $cpaRemedytemplateFilePath) -Level Info
$cpaRemedyTemplateExcelWorkbook = $cpaRemedyTemplateExcelFileObject.Workbooks.Open($cpaRemedytemplateFilePath)
$cpaRemedyTemplateExcelWorkbook.activate()


Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from CPA Remedy Template File." ) -Level Info
$cpaRemedyTempExcelWorkSheet = $cpaRemedyTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$cpaRemedyTempExcelWorkSheet.Activate()
Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from CPA Remedy Template File." ) -Level Info
$cpaRemedyTempExcelHeaderRange = $cpaRemedyTempExcelWorkSheet.Cells.Item(1,1).EntireRow
$cpaRemedyTempExcelHeaderRange.Delete()
$cpaRemedyTempExcelWorkSheetRange = $cpaRemedyTempExcelWorkSheet.Range("A:I")
#$cpaJIRADailyIncidentRowCount = $cpaJIRATempdataExcelWorksheetRange.Rows.Count
Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from CPA Remedy Template File." ) -Level Info
$cpaRemedyTempExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from CPA Remedy Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info

#$TotalIncidentCount = $TotalIncidentCount +1

$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$cpaRemedyIncidentCount = $cpaRemedyTemplateExcelFileObject.WorksheetFunction.CountIf($cpaRemedyTempExcelWorkSheet.Range("A1:" + "A" + $cpaRemedyTempExcelWorkSheet.Rows.Count), "<>") 

$TotalIncidentCount = $TotalIncidentCount + $cpaRemedyIncidentCount
Write-Log -Message ("Total Number of Incidents Of 'CPA Remedy Incidents in Worksheet [" + $TempworksheetName + "] of UKI NCR HUB Template File are [" + $cpaRemedyIncidentCount  + "]") -Level Info
Write-Log -Message ("Closing File --> " + $xchangingtemplateFilePath) -Level Info
$cpaRemedyTemplateExcelFileObject.Quit()


Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA CPA REMEDY DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info





Write-Log -Message ("**************** WRITING CPA JIRA INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

Write-Log -Message ("Going to Open CPA JIRA Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for CPA JIRA Template Hurdle Workbook") -Level Info
Write-Log -Message ("Going to Open CPA Remedy Template File") -Level Info
Write-Log -Message ("Going to Open CPA JIRA Template File") -Level Info
$cpaJIRATemplateExcelFileObject = New-Object -ComObject excel.application
$cpaJIRATemplateExcelFileObject.Visible = $true
$cpaJIRATemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening CPA JIRA Template File --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRATemplateExcelWorkbook = $cpaJIRATemplateExcelFileObject.Workbooks.Open($cpaJIRAtemplateFilePath)
$cpaJIRATemplateExcelWorkbook.activate()

Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from CPA JIRA Template File." ) -Level Info
$cpaJIRATemplateExcelWorkSheet = $cpaJIRATemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$cpaJIRATemplateExcelWorkSheet.Activate()
Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from CPA JIRA Template File." ) -Level Info
$cpaJIRATemplateExcelHeaderRange = $cpaJIRATemplateExcelWorkSheet.Cells.Item(1,1).EntireRow
$cpaJIRATemplateExcelHeaderRange.Delete()
$cpaJIRATemplateExcelWorkSheetRange = $cpaJIRATemplateExcelWorkSheet.Range("A:I")
#$cpaJIRADailyIncidentRowCount = $cpaJIRATempdataExcelWorksheetRange.Rows.Count
Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from CPA JIRA Template File." ) -Level Info
$cpaJIRATemplateExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from CPA JIRA Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info


$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$cpaJIRAIncidentCount = $cpaJIRATemplateExcelFileObject.WorksheetFunction.CountIf($cpaJIRATemplateExcelWorkSheet.Range("A1:" + "A" + $cpaJIRATemplateExcelWorkSheet.Rows.Count), "<>") 

$TotalIncidentCount = $TotalIncidentCount + $cpaJIRAIncidentCount
Write-Log -Message ("Total Number of Incidents Of 'CPA Remedy Incidents in Worksheet [" + $TempworksheetName + "] of UKI NCR HUB Template File are [" + $cpaRemedyIncidentCount  + "]") -Level Info
Write-Log -Message ("Closing File --> " + $cpaJIRAtemplateFilePath) -Level Info
$cpaJIRATemplateExcelFileObject.Quit()


Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA CPA JIRA DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info


Write-Log -Message ("**************** WRITING QBE INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

Write-Log -Message ("Going to Open QBE Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for QBE Template Hurdle Workbook") -Level Info
Write-Log -Message ("Going to Open QBE Template File") -Level Info
Write-Log -Message ("Going to Open QBE  Template File") -Level Info
$qbeTemplateExcelFileObject = New-Object -ComObject excel.application
$qbeTemplateExcelFileObject.Visible = $true
$qbeTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening QBE Template File --> " + $qbetemplateFilePath) -Level Info
$qbeTemplateExcelWorkbook = $qbeTemplateExcelFileObject.Workbooks.Open($qbetemplateFilePath)
$qbeTemplateExcelWorkbook.activate()

Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from QBE Template File." ) -Level Info
$qbeTemplateExcelWorksheet = $qbeTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$qbeTemplateExcelWorksheet.Activate()
Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from QBE Template File." ) -Level Info
$qbeTemplateExcelDeleteHeaderRange = $qbeTemplateExcelWorksheet.Cells.Item(1,1).EntireRow
$qbeTemplateExcelDeleteHeaderRange.Delete()
$qbeTemplateExcelWorksheetRange = $qbeTemplateExcelWorksheet.Range("A:I")
#$cpaJIRADailyIncidentRowCount = $cpaJIRATempdataExcelWorksheetRange.Rows.Count
Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from QBE JIRA Template File." ) -Level Info
$qbeTemplateExcelWorksheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from QBE JIRA Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info


$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$qbeIncidentTotalCount = $qbeTemplateExcelFileObject.WorksheetFunction.CountIf($qbeTemplateExcelWorksheetRange.Range("A1:" + "A" + $qbeTemplateExcelWorksheetRange.Rows.Count), "<>") 

$TotalIncidentCount = $TotalIncidentCount + $qbeIncidentTotalCount
Write-Log -Message ("Total Number of Incidents Of 'QBE Incidents in Worksheet [" + $TempworksheetName + "] of UKI NCR HUB Template File are [" + $qbeIncidentTotalCount  + "]") -Level Info
Write-Log -Message ("Closing File --> " + $qbetemplateFilePath) -Level Info
$qbeTemplateExcelFileObject.Quit()


Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA QBE DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info

Write-Log -Message ("**************** WRITING QBE SR FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info

Write-Log -Message ("Going to Open QBE Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for QBE Template Hurdle Workbook") -Level Info
Write-Log -Message ("Going to Open QBE Template File") -Level Info

$qbeSRTemplateExcelFileObject = New-Object -ComObject excel.application
$qbeSRTemplateExcelFileObject.Visible = $true
$qbeSRTemplateExcelFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening QBE Template File --> " + $qbeSRtemplateFilePath) -Level Info
$qbeSRTemplateExcelWorkbook = $qbeSRTemplateExcelFileObject.Workbooks.Open($qbeSRtemplateFilePath)
$qbeSRTemplateExcelWorkbook.activate()

Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from QBE Template File." ) -Level Info
$qbeSRTemplateExcelWorksheet = $qbeSRTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$qbeSRTemplateExcelWorksheet.Activate()
Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from QBE Template File." ) -Level Info
$qbeSRTemplateExcelDeleteHeaderRange = $qbeSRTemplateExcelWorksheet.Cells.Item(1,1).EntireRow
$qbeSRTemplateExcelDeleteHeaderRange.Delete()
$qbeSRTemplateExcelWorksheetRange = $qbeSRTemplateExcelWorksheet.Range("A:I")
#$cpaJIRADailyIncidentRowCount = $cpaJIRATempdataExcelWorksheetRange.Rows.Count
Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from QBE SR Template File." ) -Level Info
$qbeSRTemplateExcelWorksheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from QBE SR Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info


$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$qbeSRTotalCount = $qbeSRTemplateExcelFileObject.WorksheetFunction.CountIf($qbeSRTemplateExcelWorksheetRange.Range("A1:" + "A" + $qbeSRTemplateExcelWorksheetRange.Rows.Count), "<>") 
$TotalIncidentCount = $TotalIncidentCount + $qbeSRCount

Write-Log -Message ("Total Number of Incidents Of 'QBE SR in Worksheet [" + $TempworksheetName + "] of UKI NCR HUB Template File are [" + $qbeSRTotalCount  + "]") -Level Info
Write-Log -Message ("Closing File --> " + $qbeSRtemplateFilePath) -Level Info
$qbeSRTemplateExcelFileObject.Quit()


Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA QBE DAILY SR RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info


Write-Log -Message ("**************** WRITING BELERON CONSOLIDATED INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info


Write-Log -Message ("Going to Open BELERON Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for BELERON Hurdle Workbook") -Level Info
$belronTemplateFileObject = New-Object -ComObject excel.application
$belronTemplateFileObject.Visible = $true
$belronTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening BELERON Template File --> " + $belronTemplateFilePath) -Level Info
$belronTemplateExcelWorkbook = $belronTemplateFileObject.Workbooks.Open($belronTemplateFilePath)
$belronTemplateExcelWorkbook.activate()
Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from Beleron Template File." ) -Level Info
$belronTempExcelWorKSHEET = $belronTemplateExcelWorkbook.Worksheets.Item($TempworksheetName)
$belronTempExcelWorKSHEET.Activate()

Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from Beleron Template File." ) -Level Info
$beleronempExcelDeleteHeaderRange = $belronTempExcelWorKSHEET.Cells.Item(1,1).EntireRow
$beleronempExcelDeleteHeaderRange.Delete()



$belronTempExcelWorKSHEETRange = $belronTempExcelWorKSHEET.Range("A:I")

Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from Beleron Template File." ) -Level Info
$belronTempExcelWorKSHEETRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from Beleron Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info
Write-Log -Message ($TotalIncidentCount)
$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$BeleronTotalIncidentCount = $belronTemplateFileObject.WorksheetFunction.CountIf($belronTempExcelWorKSHEET.Range("A1:" + "A" + $belronTempExcelWorKSHEET.Rows.Count), "<>") 
$TotalIncidentCount = $TotalIncidentCount + $BeleronTotalIncidentCount


Write-Log -Message ("Closing File --> " + $belronTemplateFilePath) -Level Info
$belronTemplateFileObject.Quit()
Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA BELERON DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info



Write-Log -Message ("**************** WRITING EXEVA CONSOLIDATED INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info


Write-Log -Message ("Going to Open EXEVA Template File") -Level Info
Write-Log -Message ("Going to Create Excel Object for EXEVA Hurdle Workbook") -Level Info
$ExovaTemplateFileObject = New-Object -ComObject excel.application
$ExovaTemplateFileObject.Visible = $true
$ExovaTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening EXEVA Template File --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTemplateExcelWorkBook = $ExovaTemplateFileObject.Workbooks.Open($ExovaTemplateFilePath)
$ExovaTemplateExcelWorkBook.activate()
Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from Exova Template File." ) -Level Info
$ExovaTempExcelWorkSheet = $ExovaTemplateExcelWorkBook.Worksheets.Item($TempworksheetName)
$ExovaTempExcelWorkSheet.Activate()

Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from Exova Template File." ) -Level Info
$ExovaempExcelDeleteHeaderRange = $ExovaTempExcelWorkSheet.Cells.Item(1,1).EntireRow
$ExovaempExcelDeleteHeaderRange.Delete()



$ExovaTempExcelWorkSheetRange = $ExovaTempExcelWorkSheet.Range("A:I")

Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from Exova Template File." ) -Level Info
$ExovaTempExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from Exova Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info
Write-Log -Message ($TotalIncidentCount)
$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$ExovaTotalIncidentCount = $ExovaTemplateFileObject.WorksheetFunction.CountIf($ExovaTempExcelWorkSheet.Range("A1:" + "A" + $ExovaTempExcelWorkSheet.Rows.Count), "<>") 
$TotalIncidentCount = $TotalIncidentCount + $ExovaTotalIncidentCount


Write-Log -Message ("Closing File --> " + $ExovaTemplateFilePath) -Level Info
$ExovaTemplateFileObject.Quit()
Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA EXOVA DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info




Write-Log -Message ("**************** WRITING NATIONAL GRID US CONSOLIDATED INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info



Write-Log -Message ("Going to Open NATIONAL GRID US Template File") -Level Info
$ngridTemplateFileObject = New-Object -ComObject excel.application
$ngridTemplateFileObject.Visible = $true
$ngridTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NATIONAL GRID US Template File --> " + $ngridTemplateFilePath) -Level Info
$ngridTemplateExcelWorkBook = $ngridTemplateFileObject.Workbooks.Open($ngridTemplateFilePath)
$ngridTemplateExcelWorkBook.activate()

Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from NATIONAL GRID US Template File." ) -Level Info
$ngridTempExcelWorkSheet = $ngridTemplateExcelWorkBook.Worksheets.Item($TempworksheetName)
$ngridTempExcelWorkSheet.Activate()

Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from NATIONAL GRID US Template File." ) -Level Info
$ngridTempExcelWorkSheetRange = $ngridTempExcelWorkSheet.Cells.Item(1,1).EntireRow
$ngridTempExcelWorkSheetRange.Delete()



$ngridTempExcelWorkSheetRange = $ngridTempExcelWorkSheet.Range("A:I")

Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from NATIONAL GRID US Template File." ) -Level Info
$ngridTempExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from NATIONAL GRID US Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info
Write-Log -Message ($TotalIncidentCount)
$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$ngridTotalIncidentCount = $ngridTemplateFileObject.WorksheetFunction.CountIf($ngridTempExcelWorkSheet.Range("A1:" + "A" + $ngridTempExcelWorkSheet.Rows.Count), "<>") 
$TotalIncidentCount = $TotalIncidentCount + $ngridTotalIncidentCount


Write-Log -Message ("Closing File --> " + $ngridTemplateFileObject) -Level Info
$ngridTemplateFileObject.Quit()
Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA NATIONAL US GRID DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info


Write-Log -Message ("**************** WRITING NATIONAL GRID UK CONSOLIDATED INCIDENT FILTERED DATA TO UKI NCR HUB HURDLE RAW DATA SHEET *************************************") -Level Info
Write-Log -Message ("**************** ########################################## *************************************") -Level Info



Write-Log -Message ("Going to Open NATIONAL GRID UK Template File") -Level Info
$ngridukTemplateFileObject = New-Object -ComObject excel.application
$ngridukTemplateFileObject.Visible = $true
$ngridukTemplateFileObject.DisplayAlerts=$false
Write-Log -Message ("Opening NATIONAL GRID UK  Template File --> " + $ngridukTemplateFilePath) -Level Info
$ngridukTemplateExcelWorkBook = $ngridukTemplateFileObject.Workbooks.Open($ngridukTemplateFilePath)
$ngridukTemplateExcelWorkBook.activate()

Write-Log -Message ("Selecting the Temp data in worksheet [" + $TempworksheetName+"] from NATIONAL GRID Template File." ) -Level Info
$ngridukTempExcelWorkSheet = $ngridukTemplateFileObject.Worksheets.Item($TempworksheetName)
$ngridukTempExcelWorkSheet.Activate()

Write-Log -Message ("Deleting First Header Row from worksheet [" + $TempworksheetName+"] from NATIONAL GRID Template File." ) -Level Info
$ngridukTempExcelWorkSheetRange = $ngridukTempExcelWorkSheet.Cells.Item(1,1).EntireRow
$ngridukTempExcelWorkSheetRange.Delete()



$ngridukTempExcelWorkSheetRange = $ngridukTempExcelWorkSheet.Range("A:I")

Write-Log -Message ("Copying the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from NATIONAL GRID UK Template File." ) -Level Info
$ngridukTempExcelWorkSheetRange.copy()
Write-Log -Message ("Pasting the Temp data from Range[A:I] in worksheet [" + $TempworksheetName+"] from NATIONAL GRID UK Template File to Worksheet [" + $ukNCRHUBExcelRawDataWorkSheetName + "] of UKI NCR HUB Template File") -Level Info
Write-Log -Message ($TotalIncidentCount)
$ukNCRHUBExcelWorkSheet.Paste($ukNCRHUBExcelWorkSheet.Range("A$TotalIncidentCount"))

$ngridUKTotalIncidentCount = $ngridukTemplateFileObject.WorksheetFunction.CountIf($ngridukTempExcelWorkSheet.Range("A1:" + "A" + $ngridukTempExcelWorkSheet.Rows.Count), "<>") 
$TotalIncidentCount = $TotalIncidentCount + $ngridUKTotalIncidentCount


Write-Log -Message ("Closing File --> " + $ngridukTemplateFileObject) -Level Info
$ngridukTemplateFileObject.Quit()
Write-Log -Message ("**************** COPIED FILTERED INCIDENT DATA NATIONAL GRID UK DAILY INCIDENT RAW DATA TO UKI NCR HUB SHEET *************************************") -Level Info
Write-Log -Message ("**************** ##################################### *************************************") -Level Info



Write-Log -Message (" Saving and Closing File --> " + $ukNCRHUBExcelFilePath) -Level Info

$ukNCRHUBExcelWorkSheet.Columns.Item(2).NumberFormat = "DD-MM-YYYY HH:MM:SS"
$ukNCRHUBExcelWorkSheet.Columns.Item(9).NumberFormat = "DD-MM-YYYY HH:MM:SS"


$ukNCRHUBExcelRawDataIncidentCount = $ukNCRHUBExcelFileObject.WorksheetFunction.CountIf($ukNCRHUBExcelWorkSheet.Range("A1:" + "A" + $ukNCRHUBExcelWorkSheet.Rows.Count), "<>") 

$ukNCRHUBExcelRawDataIncidentCount

$ukNCRHUBExcelWorkSheet.Range("Q2:BH$ukNCRHUBExcelRawDataIncidentCount").Formula = $ukNCRHUBExcelWorkSheet.Range("Q2:BH2").Formula


#$ukNCRHUBExcelWorkBook.Save()
$ukNCRHUBExcelWorkBook.close($true)

$ukNCRHUBExcelFileObject.Quit()
Write-Log -Message ("**************** PROCESSED UKI NCR HUB HURDLE FILE *************************************") -Level Info

Write-Log -Message ("**************** PROCESSED UKI NCR HUB HURDLE FILE *************************************") -Level Info

Write-Log -Message ("**************** FLUSHING EXCEL PROCESS *************************************") -Level Info
Get-Process -Name "*EXCEL*" | Stop-Process
Write-Log -Message ("**************** SUCCESSFULLY FLUSHING EXCEL PROCESS *************************************") -Level Info
}




