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

    #### Exporting libraries data to CSV files ####
    if ($exportDataToCSVFiles) {
      Write-Host "Exporting list data to CSV files..."
      Export-List -LogFile $config.LOG_LIQUIDACIONES -ListName $config.LISTA_LIQUIDACIONES -Fields $config.FIELDS_LIQUIDACIONES -Delimiter "|"
      Export-List -LogFile $config.LOG_LIQUIDACION_DETALLE -ListName $config.LISTA_LIQUIDACION_DETALLE -Fields $config.FIELDS_LIQUIDACION_DETALLE -Delimiter "|"
      Export-List -LogFile $config.LOG_LIQUIDACION_CABECERA -ListName $config.LISTA_LIQUIDACION_CABECERA -Fields $config.FIELDS_LIQUIDACION_CABECERA -Delimiter "|"
      Export-List -LogFile $config.LOG_SOLICITUD -ListName $config.LISTA_SOLICITUD -Fields $config.FIELDS_SOLICITUD -Delimiter "|"
    }

    #### Importing CSV files to collections ####
    if ($importCsvDataToCol) { 
      Write-Host "Importing CSV files..."
      $itemsLiquidaciones = Import-Csv $config.LOG_LIQUIDACIONES -Delimiter "|"
      $itemsLiquidacionDetalle = Import-Csv $config.LOG_LIQUIDACION_DETALLE -Delimiter "|"
      $itemsLiquidacionCabecera = Import-Csv $config.LOG_LIQUIDACION_CABECERA -Delimiter "|"
      $itemsSolicitud = Import-Csv $config.LOG_SOLICITUD -Delimiter "|"
    }

    #### Converting collections to dictionaries ####
    if ($convertColToDictionaries) {
      Write-Host "Creating hash tables..."
      $dictLiquidacionDetalle = @{}
      $dictLiquidacionCabecera = @{}
      $dictSolicitud = @{}

      foreach ($item in $itemsLiquidacionDetalle) {
        if ($null -ne $item.ID -and $item.ID -ne '') {
          $dictLiquidacionDetalle[$item.ID] = $item
        }
      }

      foreach ($item in $itemsLiquidacionCabecera) {
        if ($null -ne $item.Title -and $item.Title -ne '') {
          $dictLiquidacionCabecera[$item.Title] = $item
        }
      }

      foreach ($item in $itemsSolicitud) {
        if ($null -ne $item.Title -and $item.Title -ne '') {
          $dictSolicitud[$item.Title] = $item
        }
      }
    }

    #### Generating results file ####
    if ($generateResultsFile) {
      Clear-Log -LogFile $config.LOG_RESULTADOS
      Write-Log $config.FIELDS_LOG_RESULTADOS -LogFile $config.LOG_RESULTADOS

      ### Iterating liquidaciones
      foreach ($item in $itemsLiquidaciones) {
        ### Get item from dictLiquidacionDetalle
        $itemDictLiquidacionDetalle = $dictLiquidacionDetalle[$item.Detalle]
        ### Validate if itemDictLiquidacionDetalle is not null
        if ($null -ne $itemDictLiquidacionDetalle) {
          $itemDictLiquidacionCabecera = $dictLiquidacionCabecera[$itemDictLiquidacionDetalle.Liquidacion_MG]
          ### Validate if itemDictLiquidacionCabecera is not null
          if ($null -ne $itemDictLiquidacionCabecera) {
            $itemDictSolicitud = $dictSolicitud[$itemDictLiquidacionCabecera.Solicitud]
            if ($itemDictSolicitud.Estado -eq "Liquidado") {
              Write-Log "$($item.ID);$($item.FileRef);$($item.Created_x0020_Date);$($item.File_x0020_Size);$($itemDictLiquidacionCabecera.PSObject.Properties.Value -join ";");$($itemDictSolicitud.Estado)" -LogFile $config.LOG_RESULTADOS
            }
          }
          else {
            $itemDictSolicitud = $dictSolicitud[$itemDictLiquidacionDetalle.Solicitud_MG]
            if ($null -ne $itemDictSolicitud) {
              if ($itemDictSolicitud.Estado -eq "Liquidado") {
                Write-Log "$($item.ID);$($item.FileRef);$($item.Created_x0020_Date);$($item.File_x0020_Size);$($itemDictLiquidacionCabecera.PSObject.Properties.Value -join ";");$($itemDictSolicitud.Estado)" -LogFile $config.LOG_RESULTADOS
              }
            }
            else {
              $createdDateTime = [DateTime]::ParseExact($item.Created_x0020_Date, "yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
              # if ($createdDateTime.Year -eq 2022) {
                if ($createdDateTime.Year -eq 2024 -and $createdDateTime.Month -gt 3 -and $createdDateTime.Month -lt 7) {
                try {
                  $fileRef = $item.FileRef
                  $fileLeafRef = Split-Path $fileRef -Leaf
                  $pathFile = "$($config.BACKUP_PATH)/$($createdDateTime.Year )/$($createdDateTime.Month)"
                  $fullPath = "$($pathFile)/$($fileLeafRef)"
                  if (Test-Path $fullPath -PathType Leaf) {
                    Write-Host "The file $($fullPath) already exists."
                  }
                  else {
                    New-Item -ItemType Directory -Force -Path $pathFile
                    Write-Host  "Downloading file $($fileRef)"
                    Get-PnPFile -Url $fileRef -Path $pathFile -Filename $fileLeafRef -AsFile -ErrorAction Ignore
                    Write-Host  "Removing file $($fileRef)"
                    Remove-PnPFile -ServerRelativeUrl $fileRef -Force
                  }
                }
                catch {
                  <#Do this if a terminating exception happens#>
                }
              }
            }
          }
        }
        else {
          $createdDateTime = [DateTime]::ParseExact($item.Created_x0020_Date, "yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
          # if ($createdDateTime.Year -eq 2022) {
            if ($createdDateTime.Year -eq 2024 -and $createdDateTime.Month -gt 3 -and $createdDateTime.Month -lt 7) {
            $fileRef = $item.FileRef
            $fileLeafRef = Split-Path $fileRef -Leaf
            $pathFile = "$($config.BACKUP_PATH)/$($createdDateTime.Year )/$($createdDateTime.Month)"
            $fullPath = "$($pathFile)/$($fileLeafRef)"
            if (Test-Path $fullPath -PathType Leaf) {
              Write-Host "The file $($fullPath) already exists."
            }
            else {
              New-Item -ItemType Directory -Force -Path $pathFile
              Write-Host  "Downloading file $($fileRef)"
              Get-PnPFile -Url $fileRef -Path $pathFile -Filename $fileLeafRef -AsFile -ErrorAction Ignore
              Write-Host  "Removing file $($fileRef)"
              Remove-PnPFile -ServerRelativeUrl $fileRef -Force
            }
          }
        }
      }
    }

    ### Downloading files from results ####
    if ($downloadFiles) {
      $itemsResultados = Import-Csv $config.LOG_RESULTADOS -Delimiter ";"

      foreach ($item in $itemsResultados) {

        $fileRef = $item.FileRef
        $fileLeafRef = Split-Path $fileRef -Leaf
        $createdDateTime = [DateTime]::ParseExact($item.Created_x0020_Date, "yyyy-MM-ddTHH:mm:ssZ", [System.Globalization.CultureInfo]::InvariantCulture)
        #### Create path if not exists
        if ($createdDateTime.Year -eq 2024 -and $createdDateTime.Month -gt 3 -and $createdDateTime.Month -lt 7) {
          if ($null -ne $item.AnioSAP -and $null -ne $item.Sociedad -and $null -ne $item.Solicitud) {
            $pathFile = "$($config.BACKUP_PATH)/$($item.AnioSAP)/$($createdDateTime.Month)/$($item.Sociedad)/$($item.Solicitud)"
            $fullPath = "$($pathFile)/$($fileLeafRef)"
            if (Test-Path $fullPath -PathType Leaf) {
              Write-Host "The file $($fullPath) already exists."
            }
            else {
              New-Item -ItemType Directory -Force -Path $pathFile
              Get-PnPFile -Url $fileRef -Path $pathFile -Filename $fileLeafRef -AsFile
            }
            Write-Host  "Removing file $($fileRef)"
            Remove-PnPFile -ServerRelativeUrl $fileRef -Force
          }
        }
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