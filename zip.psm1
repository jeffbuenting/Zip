function Expand-ZIPFile {

<#
    .Link
        http://www.howtogeek.com/tips/how-to-extract-zip-files-using-powershell/

    .Link
        Shell.Application.CopyHere vOptions
        https://msdn.microsoft.com/en-us/library/ms723207.aspx
#>

    [CmdletBinding()]
    param (

        [String]$Path, 

        [String]$destination,

        [Switch]$DoNotOverwrite
    )

    # ----- No options
    $vOption = 0

    If ( $DoNoTOverwrite ) {
            Write-verbose " Will not overwrite if the file exists"

            if ( -Not (Test-Path -Path 'c:\temp\TempUnzip') ) { New-Item -path 'c:\temp\TempUnzip' -ItemType directory | Out-Null }

            $shell = new-object -com shell.application
            $zip = $shell.NameSpace($Path)
            foreach($item in $zip.items()) {
                $shell.Namespace('c:\temp\TempUnzip').copyhere($item,$vOption)
            }

            Copy-Item -Path 'c:\temp\TempUnzip\*' -Destination $Destination -Recurse
            Remove-Item -Path 'c:\temp\TempUnzip' -Recurse -Force | Out-Null
        }
        else {
            Write-verbose "No Options Specified"
            
            Write-Verbose "Checking if $destination Exists"
            If ( -Not (Test-Path -Path $Destination) ) {
                Write-verbose "Creating path"
                New-Item -Path $Destination -ItemType Directory | Out-Null
            }

            $shell = new-object -com shell.application
            $zip = $shell.NameSpace($Path)
            foreach($item in $zip.items()) {
                $shell.Namespace($Destination).copyhere($item,$vOption)
            }
    }
}