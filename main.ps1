function Clear-Log {
  Param(
    [string]$LogFile
  )
  Clear-Content $LogFile -ErrorAction Ignore
}
function Write-Log {
  Param (
    [string]$LogString,
    [string]$LogFile
  )
  Add-Content -Path $LogFile -Value $LogString
}
function Export-List {
  Param(
    [string]$LogFile,
    [string]$ListName,
    [string[]]$Fields,
    [string]$Delimiter = ",",
    [Int]$PageSize = 2000,
    [string]$LogError = "errors.log",
    [string]$Query = "<View><Query></Query></View>"
  )
  Clear-Log $LogFile
  Write-Log ($Fields -join $Delimiter) -LogFile $LogFile
  Get-PnPListItem -List $ListName -Fields $Fields -PageSize $PageSize -ScriptBlock { 
    Param($items) $items.Context.ExecuteQuery()
  } | ForEach-Object {
    $a = $_
    Write-Log (($Fields | ForEach-Object {
          if ($null -ne $a.FieldValues[$_]) {
            if ($a.FieldValues[$_].GetType().Name -eq 'FieldLookupValue') {
              $a.FieldValues[$_].LookupValue
            }
            else {              
              $a.FieldValues[$_] 
            }   
          }
          else {
            Write-Log "Error in item Id: $($a.Id) from List: $($ListName) -> Field $($_) is null." -LogFile $LogError
            $null
          }
        }) -join $Delimiter) -LogFile $LogFile
  }
}
function Backup-File {
  Param(
    [object]$Item, 
    [string]$BackupPath = $config.BACKUP_PATH,
    [string]$Type = 'default',
    [Int]$Year,
    [Int]$MonthGt,
    [Int]$MonthLt
  )
  # Write-Host "Backing up item: $($item)"
  ### Validate if type='custom'
  if ($Type -eq 'custom') {
    if ($null -eq $Item.AnioSAP -or $null -eq $Item.Sociedad -or $null -eq $Item.Solicitud) {
      Write-Host 'There is some missing properties (AnioSAP, Sociedad, Solicitud)'
      return $false
    } 
  }
    
  ### Item doesn't have Detalle, so it will backed up following a default hierarchy: [year]/[month]/[filename]
  $createdDateTime = [DateTime]::ParseExact($Item.Created_x0020_Date, "yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)

  # Define date range to be backed up
  if ($createdDateTime.Year -eq $year -and $createdDateTime.Month -gt $monthGt -and $createdDateTime.Month -lt $monthLt) {
    $fileRef = $Item.FileRef
    $fileLeafRef = Split-Path $fileRef -Leaf
    if ($Type -eq 'default') {
      $dirPath = "$($BackupPath)/$($createdDateTime.Year)/$($createdDateTime.Month)"
    }
    else {
      $dirPath = "$($BackupPath)/$($Item.AnioSAP)/$($createdDateTime.Month)/$($item.Sociedad)/$($item.Solicitud)"
    }
    $fullPath = "$($dirPath)/$($fileLeafRef)"
    if (Test-Path $fullPath -PathType Leaf) {
      Write-Host "The file $($fullPath) has already been backed up."
    }
    else {
      New-Item -ItemType Directory -Force -Path $dirPath
      Write-Host  "Downloading file $($fileRef)"
      Get-PnPFile -Url $fileRef -Path $dirPath -Filename $fileLeafRef -AsFile -ErrorAction Ignore
      Write-Host  "Removing file $($fileRef)"
      Remove-PnPFile -ServerRelativeUrl $fileRef -Force
    }
  }
}

$maxRetries = 3
$retryCount = 0
$success = $false

while (-not $success -and $retryCount -lt $maxRetries) {
  try {
    Write-Host "Starting script execution (Attempt $($retryCount + 1) of $maxRetries)"
    Write-Host "Loading configuration..."
    
    $config = Import-PowerShellDataFile -Path ".\config.psd1"
    $exportDataToCSVFiles = $false
    $importCsvDataToCol = $true
    $convertColToDictionaries = $true
    $generateResultsFile = $true
    $downloadFiles = $false

    Write-Host "Connecting to SharePoint site..."
    Connect-PnPOnline -Url $config.SITE_URL -ClientId $config.CLIENT_ID -ClientSecret $config.CLIENT_SECRET

    #### Export libraries data to CSV files ####
    if ($exportDataToCSVFiles) {
      Write-Host "Exporting list data to CSV files..."
      Export-List -LogFile $config.LOG_LIQUIDACIONES -ListName $config.LISTA_LIQUIDACIONES -Fields $config.FIELDS_LIQUIDACIONES -Delimiter "|"
      Export-List -LogFile $config.LOG_LIQUIDACION_DETALLE -ListName $config.LISTA_LIQUIDACION_DETALLE -Fields $config.FIELDS_LIQUIDACION_DETALLE -Delimiter "|"
      Export-List -LogFile $config.LOG_LIQUIDACION_CABECERA -ListName $config.LISTA_LIQUIDACION_CABECERA -Fields $config.FIELDS_LIQUIDACION_CABECERA -Delimiter "|"
      Export-List -LogFile $config.LOG_SOLICITUD -ListName $config.LISTA_SOLICITUD -Fields $config.FIELDS_SOLICITUD -Delimiter "|"
    }

    #### Import CSV files to collections ####
    if ($importCsvDataToCol) { 
      Write-Host "Importing CSV files..."
      $itemsCsvLiquidaciones = Import-Csv $config.LOG_LIQUIDACIONES -Delimiter "|"
      $itemsCsvLiquidacionDetalle = Import-Csv $config.LOG_LIQUIDACION_DETALLE -Delimiter "|"
      $itemsCsvLiquidacionCabecera = Import-Csv $config.LOG_LIQUIDACION_CABECERA -Delimiter "|"
      $itemsCsvSolicitud = Import-Csv $config.LOG_SOLICITUD -Delimiter "|"
    }

    #### Convert collections to dictionaries for better performance ####
    if ($convertColToDictionaries) {
      Write-Host "Creating dictionaries..."
      $dictLiquidacionDetalle = @{}
      $dictLiquidacionCabecera = @{}
      $dictSolicitud = @{}

      foreach ($item in $itemsCsvLiquidacionDetalle) {
        if ($null -ne $item.ID -and $item.ID -ne '') {
          $dictLiquidacionDetalle[$item.ID] = $item
        }
      }

      foreach ($item in $itemsCsvLiquidacionCabecera) {
        if ($null -ne $item.Title -and $item.Title -ne '') {
          $dictLiquidacionCabecera[$item.Title] = $item
        }
      }

      foreach ($item in $itemsCsvSolicitud) {
        if ($null -ne $item.Title -and $item.Title -ne '') {
          $dictSolicitud[$item.Title] = $item
        }
      }
    }

    #### Generate results file ####
    if ($generateResultsFile) {
      # Generate Log 
      Clear-Log -LogFile $config.LOG_RESULTADOS
      Write-Log $config.FIELDS_LOG_RESULTADOS -LogFile $config.LOG_RESULTADOS

      ### Iterate liquidaciones items
      foreach ($item in $itemsCsvLiquidaciones) {
        ### Get item details
        $itemDictLiquidacionDetalle = $dictLiquidacionDetalle[$item.Detalle]
        ### Validate if item has details
        if ($null -ne $itemDictLiquidacionDetalle) {
          ### Get item header using Liquidacion_MG from details
          $itemDictLiquidacionCabecera = $dictLiquidacionCabecera[$itemDictLiquidacionDetalle.Liquidacion_MG]
          ### Validate if item has header
          if ($null -ne $itemDictLiquidacionCabecera) {
            ## Get item solicitud using Solicitud from header
            $itemDictSolicitud = $dictSolicitud[$itemDictLiquidacionCabecera.Solicitud]

            if ($itemDictSolicitud.Estado -eq "Liquidado") {
              Write-Host "Backing up item Liquidado: $($itemDictSolicitud)"
              Write-Log "$($item.ID);$($item.FileRef);$($item.Created_x0020_Date);$($item.File_x0020_Size);$($itemDictLiquidacionCabecera.PSObject.Properties.Value -join ";");$($itemDictSolicitud.Estado)" -LogFile $config.LOG_RESULTADOS
            }
          }
          else {
            ### Get item solicitud using Solicitud_MG from details
            $itemDictSolicitud = $dictSolicitud[$itemDictLiquidacionDetalle.Solicitud_MG]
            if ($null -ne $itemDictSolicitud) {
              if ($itemDictSolicitud.Estado -eq "Liquidado") {
                Write-Host "Backing up item Liquidado: $($itemDictSolicitud)"
                Write-Log "$($item.ID);$($item.FileRef);$($item.Created_x0020_Date);$($item.File_x0020_Size);$($itemDictLiquidacionCabecera.PSObject.Properties.Value -join ";");$($itemDictSolicitud.Estado)" -LogFile $config.LOG_RESULTADOS
              }
            }
            else {
              ### The item doesn't have liquidacionCabecera nor Solicitud_MG
              Backup-File -Item $item -BackupPath $config.BACKUP_PATH -Year 2024 -MonthGt 10 -MonthLt 13
            }
          }
        }
        else {
          ### The item doesn't have liquidacionDetalle
          Backup-File -Item $item -BackupPath $config.BACKUP_PATH -Year 2024 -MonthGt 10 -MonthLt 13
        }
      }
    }

    ### Downloading files from results ####
    if ($downloadFiles) {
      $itemsResultados = Import-Csv $config.LOG_RESULTADOS -Delimiter ";"

      foreach ($item in $itemsResultados) {
        Backup-File -Item $item -BackupPath $config.BACKUP_PATH -Type 'custom' -Year 2024 -MonthGt 10 -MonthLt 13
      }
    }
    $success = $true  # If we reach this point without exceptions, set success to true
  }
  catch {
    $retryCount++
    Write-Warning "An error occurred (Attempt $retryCount of $maxRetries): $($_)"
    if ($retryCount -lt $maxRetries) {
      Write-Host "Retrying in 5 seconds..."
      Start-Sleep -Seconds 5
    }
    else {
      Write-Error "Max retries reached. Script failed."
    }
  }
}

if ($success) {
  Write-Host "Script completed successfully."
}
