    Set-executionpolicy remotesigned
    Set-ExecutionPolicy unrestricted 
    <##################################################################
    #                                                                 #
    #       Name:        XML to XLSX Converter                        #
    #       Create Date: 04-19-2018                                   #
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
    #       Last Update: 04-19-2018                                   #
    #                                                                 #
    #                                                                 #
    #                                                                 #
    ##################################################################>


$in_root = "C:\Users\eramos\Desktop\TEST\ValoraSentManifests.xml"
$out_root = "C:\Users\eramos\Desktop\TEST\ValoraSentManifests.csv"

$files = Get-ChildItem -Path $in_root -File

ForEach($file in $files) {

    $xml = [xml](Get-Content -Path $file.FullName -Raw)
    $props = @{}
    ForEach($item in $xml.table.row.col) {
        $props[$item.name] = $item."#text"
    }

    [PSCustomObject]$props | 
    Export-Csv -NoTypeInformation -Path (Join-Path -Path $out_root -ChildPath "$($file.BaseName).csv")

}