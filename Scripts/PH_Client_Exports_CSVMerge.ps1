    Set-executionpolicy remotesigned
    Set-ExecutionPolicy unrestricted 
    <##################################################################
    #                                                                 #
    #       Name:        Client Export CSV Merge                      #
    #       Create Date: 04-17-2018                                   #
    #       Author:      Emilio Ramos                                 #
    #                                                                 #
    #                                                                 #
    #       Description: This scripts converts all of the .xlsx       #
    #                    its no CSVs (taking each tab as its own      #
    #                    file) and then combines all of those files   #
    #                    into one large CSV file for import into      # 
    #                    other downstream applications.               #
    #                                                                 #
    #                                                                 #
    #       Last Update: 04-17-2018                                   #
    #                                                                 #
    #                                                                 #
    #                                                                 #
    ##################################################################>
   
   
    
    <# User Imputs #>
    Write-Output "Initializing Base Variables..."

    $baseFolder = "C:\Users\eramos\Desktop\EPIQ\ZZ - Active Projects\AACER Throughput Productivity Reporting\Recon Test\ClientFiles - Final Output\"
    $destFolder = "C:\Users\eramos\Desktop\EPIQ\ZZ - Active Projects\AACER Throughput Productivity Reporting\Main\EXPORT"

    <# Initialize #>
    ## $excelFile = $baseFolder + $excelFileName + ".xlsx"
    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false
    
    $n = 0


    <#Function that takes an .xlsx filename, and spits out
      each tab in the workbook as a CSV or Other file#>
    Function ExportWSToCSV ($excelFileName, $csvLoc,$fileNameShort)
{
   
    $parentFolder = Split-Path (Split-Path $excelFileName -Parent) -Leaf  #Stores the Folder name (IE, the datefolder)
    $fileName = [io.path]::GetFileNameWithoutExtension($excelFileName)
    
    $wbSaveCode = 46 #6 for CSV 46 for XML

    $E = New-Object -ComObject Excel.Application
    $E.Visible = $false
    $E.DisplayAlerts = $false

    $wb = $E.Workbooks.Open($excelFileName)

    foreach ($ws in $wb.Worksheets)
    {
        
        $name = $parentFolder+ "_" + $fileName + "_" + $ws.Name + "_" <# $ws.Name is the tab name #>
        





        $ws.SaveAs($csvLoc + $name + ".xml", $wbSaveCode)


    }
    
    
    $E.Quit()
    stop-process -processname EXCEL
}

    
    Write-Output "Identifying Search Files..."
    $container = Get-ChildItem $baseFolder -filter *.xlsx -Recurse ##-Depth 4 <#Might want to change this who knows#>

    Write-Output ("Commencing File Extractions On " + $container.Count +" Excel Files...")

    foreach($c in $container)
    {
        ExportWSToCSV -excelFileName $c.FullName -csvLoc $destFolder -fileNameShort $c.Name
        $n = $n + 1
    }

    Write-Output $n + " CSV Files Created"





