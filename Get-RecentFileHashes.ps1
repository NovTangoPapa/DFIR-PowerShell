#requires -version 2.0

<#
    .SYNOPSIS
        Find recently opened files and associated SHA256 hash
    .DESCRIPTION
        Find recently opened files and associated SHA256 hash.  This is done by taking
        the History from Internet Explorer and parsing it for local files, then running
        it through the Get-FileHash function.  
    .PARAMETER  CSVFilePath
        Specifies path to export the csv file.  Defaults to current working directory.
    .PARAMETER Hash
        Specify the hashing algorithm you would like to use.  Defaults to MD5.
        [-Hash {SHA1 | SHA256 | SHA384 | SHA512 | MACTripleDES | MD5 | RIPEMD160}]

    .EXAMPLE
        C:\PS> C:\Script\Get-RecentFileHashes.ps1 -CSVFilePath C:\Temp\ -Hash SHA256

        This Command will output two files starting with the hostname of the computer and
        ending with: [hostname]IEHistory.csv, [hostname]FileHashes.csv
    .NOTES
        Version: 0.5
        Author: n0v
        Creation Date: 6/23/2019        
        This script makes partial use of the following script, for the IE History extraction:
        https://gallery.technet.microsoft.com/scriptcenter/How-to-export-the-history-b3245ae7


#>
[CmdletBinding()]
Param
(
    [Parameter(Mandatory=$false)]
    [String]$CSVFilePath=$(Get-Location).Path,
    [Parameter(Mandatory=$false)]
    [String]$Hash='MD5'
)

$Shell = New-Object -ComObject Shell.Application
$IEHistory = $Shell.NameSpace(34)

$Objs = @()
$Items = $IEHistory.Items()
Foreach($Item in $Items)
{
    $WebSiteItems = $Item.GetFolder.Items()
    Foreach($WebSiteItem in $WebSiteItems)
    {
        If($WebSiteItem.IsFolder)
        {
            $SiteFolder = $WebSiteItem.GetFolder
            $SiteFolder.Items() | ForEach-Object{$URL = $($SiteFolder.GetDetailsOf($_,0))
                                        $Date = $($SiteFolder.GetDetailsOf($_,2))

            $Obj = New-Object -TypeName PSObject -Property @{URL = $URL
                                                            Date = $Date}
            $Objs += $Obj}
        }
    }
$Objs | Where-Object {$_.URL -match "file:*"} | 
Export-Csv -Path "$CSVFilePath\$(hostname)IEHistory.csv" -NoTypeInformation -Append -NoClobber
}

$FileHashList = Import-Csv -Path "$CSVFilePath\$(hostname)IEHistory.csv" | 
Sort-Object URL -Unique |Select-Object *,"FileHash" |
ForEach-Object{$_.URL = $_.URL -replace 'file:///', ''
$_
} | 
ForEach-Object{
    $_.FileHash = $_.FileHash -replace '', `
    $(if($(Test-Path $_.URL) -eq $true)
    {
        $(Get-FileHash -Path $_.URL -Algorithm $Hash).Hash
    }
    else
    {
        $move = "file moved or deleted"
        $move
    }
        )
    $_
}

$FileHashList | Export-Csv -Path "$CSVFilePath\$(hostname)FileHashes.csv" -NoTypeInformation