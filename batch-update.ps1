$config = Import-PowerShellDataFile -Path ".\config.psd1"

Connect-PnPOnline -Url $config.SITE_URL -ClientId $config.CLIENT_ID -ClientSecret $config.CLIENT_SECRET

# Specify the list name
$listName = "Lista: Liquidacion Detalle"

# Function to update items in batches
function Update-ListItemsBatch {
    param (
        [array]$Items,
        [int]$BatchSize = 100
    )

    $batch = New-PnPBatch
    $totalItems = $Items.Count  # Total number of items in the batch

    foreach ($item in $Items) {
        $valuesToUpdate = @{
            "Solicitud_MG"   = $item.Solicitud_MG
            "Liquidacion_MG" = $item.Liquidacion_MG
        }
        Write-Host "Updating item ID: $($item.ID) with values: $($valuesToUpdate | Out-String) from a total of $totalItems items in this batch."  # Trace item update with values and total items
        Set-PnPListItem -List $listName -Identity $item.ID -Values $valuesToUpdate -Batch $batch

        # When we reach the batch size, execute the batch and create a new one
        if ($batch.RequestCount -eq $BatchSize) {
            Write-Host "Executing batch of size $BatchSize..."  # Trace batch execution
            Invoke-PnPBatch -Batch $batch
            $batch = New-PnPBatch
        }
    }

    # Execute any remaining items in the last batch
    if ($batch.RequestCount -gt 0) {
        Write-Host "Executing final batch of size $($batch.RequestCount)..."  # Trace final batch execution
        Invoke-PnPBatch -Batch $batch
    }
}

# Update function to process items in chunks
function Process-LargeList {
    param (
        [array]$AllItems,
        [int]$ChunkSize = 10000
    )

    for ($i = 0; $i -lt $AllItems.Count; $i += $ChunkSize) {
        $chunk = $AllItems | Select-Object -Skip $i -First $ChunkSize
        Write-Host "Processing chunk $($i / $ChunkSize + 1) of $([math]::Ceiling($AllItems.Count / $ChunkSize)) with $($chunk.Count) items"  # Trace chunk processing
        Update-ListItemsBatch -Items $chunk
    }
}

$allItems = Import-Csv "Filtered_liquidacionDetalle.csv" -Delimiter ";"

# Process the large list
Write-Host "Starting to process all items..."  # Trace start of processing
Process-LargeList -AllItems $allItems

# Disconnect from SharePoint
Write-Host "Disconnecting from SharePoint..."  # Trace disconnection
Disconnect-PnPOnline