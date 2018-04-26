    Set-executionpolicy remotesigned
    Set-ExecutionPolicy unrestricted 
    <##################################################################
    #                                                                 #
    #       Name:        XML to CSV Converter                         #
    #       Create Date: 04-19-2018                                   #
    #       Author:      Emilio Ramos                                 #
    #                                                                 #
    #                                                                 #
    #       Description: Converts XML files in a Directory into       #
    #                    one CSV, appending data to the file          #
    #                                                                 #
    #       Last Update: 04-26-2018                                   #
    #                                                                 #
    #                                                                 #
    #                                                                 #
    ##################################################################>


$baseFolder = "C:\Users\eramos\Desktop\All Manifest Data\"
$destFile = "C:\Users\eramos\Desktop\DocketsAndChild.csv"

$container = Get-ChildItem $baseFolder -filter *.xml -Recurse ##-Depth 4 <#Might want to change this who knows#>
$n = 0

    Write-Output ("Commencing File Extractions On " + $container.Count +" xml Files...")

    foreach($c in $container)
    {
        [xml]$inputFile = Get-Content $c.FullName

        Write-Output ("Reading " + $c.FullName + " Progress: "+ $n + "/" + $container.Count)

        #$obj = New-Object -Typename PSobject 
        $report = @()
        $csvRecords = foreach ($dockets in $inputFile.Manifest.Dockets.ChildNodes)
        {

            foreach ($images in $dockets.Images.ChildNodes)
            {
                    $line = New-Object -Typename PSobject -Property ([ordered]@{
                        TrancheNumber = $dockets.TrancheNumber;
                        FileNumber = $dockets.FileNumber;
                        Court = $dockets.Court;
                        CourtId = $dockets.CourtId;
                        CaseId = $dockets.CaseId;
                        Division = $dockets.Division;
                        CaseNumber = $dockets.CaseNumber;
                        ODocId = $dockets.ODocId;
                        EventType = $dockets.EventType;
                        CaseLink = $dockets.CaseLink;
                        ImageId = $images.ImageId;
                        ImageType = $images.ImageType;
                        PartNumber = $images.PartNumber;
                        ImageLink = $images.ImageLink;
                        ImageName = $images.ImageName;
                        Path = $images.Path
                    })

                $report += $line
              }
         }

        
        $report | Export-Csv $destFile -Append -NoTypeInformation -Delimiter:"|" -Encoding:UTF8
        $n = $n + 1

        }

    Write-Output $n + " xml Files processed"












