###
# Script to update the Solicitud_MG and Liquidacion_MG fields in a SharePoint list
#
# Author: Jose Vega
# Date: 2023-01-16
#
# Usage:
#     powershell.exe -ExecutionPolicy Bypass -File update-lists.ps1 -ConfigFile config.psd1
#
# Parameters:
#     -ConfigFile: Path to the config.psd1 file
#
# Example:
#     powershell.exe -ExecutionPolicy Bypass -File update-lists.ps1 -ConfigFile config.psd1
#
###
$config = Import-PowerShellDataFile -Path ".\config.psd1"

Connect-PnPOnline -Url $config.SITE_URL -ClientId $config.CLIENT_ID -ClientSecret $config.CLIENT_SECRET

$ListName = "Lista: Liquidacion Detalle"
$Fields = "Solicitud", "Liquidacion", "Solicitud_MG", "Liquidacion_MG"
$PageSize = 2000
$Query = "<View><Query><Where>
            <Or>
                <IsNull><FieldRef Name='Solicitud'/></IsNull>
                <IsNull><FieldRef Name='Liquidacion'/></IsNull>
            </Or>
          </Where></Query></View>"

Get-PnPListItem -List $ListName -Fields $Fields -PageSize $PageSize -ScriptBlock { 
    Param($items) $items.Context.ExecuteQuery()
} | ForEach-Object {

    if ($null -eq $_.FieldValues["Solicitud_MG"]) {
        if (($null -ne $_.FieldValues["Solicitud"]) -and ($null -ne $_.FieldValues["Solicitud"].LookupValue) ) {
            $solicitudValue = $_.FieldValues["Solicitud"].LookupValue
            Write-Host "Updating item $($_.Id) with Solicitud_MG $($solicitudValue)"
            Set-PnPListItem -List $ListName -Identity $_ -Values @{"Solicitud_MG" = $solicitudValue }
        }
    }

    if ($null -eq $_.FieldValues["Liquidacion_MG"]) {
        if (($null -ne $_.FieldValues["Liquidacion"]) -and ($null -ne $_.FieldValues["Liquidacion"].LookupValue)) {
            $liquidacionValue = $_.FieldValues["Liquidacion"].LookupValue
            Write-Host "Updating item $($_.Id) with Liquidacion_MG $($liquidacionValue)"
            Set-PnPListItem -List $ListName -Identity $_ -Values @{"Liquidacion_MG" = $liquidacionValue }
        }
    }
}