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
    [string]$LogError = "errors.log"
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

try {
  #### Loading config ####
  $config = Import-PowerShellDataFile -Path ".\config.psd1"

  #### Connecting to SP site ####
  Connect-PnPOnline -Url $config.SITE_URL -ClientId $config.CLIENT_ID -ClientSecret $config.CLIENT_SECRET

  #### Exporting lists to files ####
  Export-List -LogFile $config.LOG_LIQUIDACIONES -ListName $config.LISTA_LIQUIDACIONES -Fields $config.FIELDS_LIQUIDACIONES -Delimiter ";"
  Export-List -LogFile $config.LOG_LIQUIDACION_DETALLE -ListName $config.LISTA_LIQUIDACION_DETALLE -Fields $config.FIELDS_LIQUIDACION_DETALLE -Delimiter ";"
  Export-List -LogFile $config.LOG_LIQUIDACION_CABECERA -ListName $config.LISTA_LIQUIDACION_CABECERA -Fields $config.FIELDS_LIQUIDACION_CABECERA -Delimiter ";"

  #### Importing Csv exported files ####
  $itemsLiquidaciones = Import-Csv $config.LOG_LIQUIDACIONES -Delimiter ";"
  $itemsLiquidacionDetalle = Import-Csv $config.LOG_LIQUIDACION_DETALLE -Delimiter ";"
  $itemsLiquidacionCabecera = Import-Csv $config.LOG_LIQUIDACION_CABECERA -Delimiter ";"

  $hashTableLiquidacionDetalle = @{}
  $hashTableLiquidacionCabecera = @{}

  foreach ($item in $itemsLiquidacionDetalle) {
    $hashTableLiquidacionDetalle[$item.ID] = $item
  }

  foreach ($item in $itemsLiquidacionCabecera) {
    $hashTableLiquidacionCabecera[$item.Title] = $item
  }

  
  #### Generating results file ####
  Clear-Log -LogFile $config.LOG_RESULTADOS
  Write-Log "ID;FileRef;ID_CABECERA;Solicitud;AnioSAP;Sociedad;Title" -LogFile $config.LOG_RESULTADOS
  foreach ($item in $itemsLiquidaciones) {
    $itemHashTableLiquidacionDetalle = $hashTableLiquidacionDetalle[$item.Detalle]
    if ($null -ne $itemHashTableLiquidacionDetalle) {
      $itemHashTableLiquidacionCabecera = $hashTableLiquidacionCabecera[$itemHashTableLiquidacionDetalle.Liquidacion]
      if ($null -ne $itemHashTableLiquidacionCabecera) {
        Write-Log "$($item.ID);$($item.FileRef);$($itemHashTableLiquidacionCabecera.PSObject.Properties.Value -join ";")" -LogFile $config.LOG_RESULTADOS
      }
    }
  }
}
catch {
  Write-Error "An error ocurred: $($_)"
}



