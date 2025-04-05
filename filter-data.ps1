$config = Import-PowerShellDataFile -Path ".\config.psd1"

Connect-PnPOnline -Url $config.SITE_URL -ClientId $config.CLIENT_ID -ClientSecret $config.CLIENT_SECRET

# Import the CSV file
$csvData = Import-Csv -Path $config.LOG_LIQUIDACION_DETALLE -Delimiter ';'

# Filter the data
$filteredData = $csvData | Where-Object {
    ($_.Solicitud -ne '' -or $_.Liquidacion -ne '') -and
    ($_.Solicitud_MG -eq '' -and $_.Liquidacion_MG -eq '')
}

# Update the filtered data
$filteredData | ForEach-Object {
    if ($_.Solicitud -ne '') {
        $_.Solicitud_MG = $_.Solicitud
    }
    if ($_.Liquidacion -ne '') {
        $_.Liquidacion_MG = $_.Liquidacion
    }
}

# Export the updated data back to CSV
$filteredData | Export-Csv -Path "Filtered_liquidacionDetalle.csv" -NoTypeInformation -Delimiter ';'