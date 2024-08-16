$CLIENT_ID = "c2f855dc-928f-46ab-9c58-0973d5864b07"
$CLIENT_SECRET = "/tgWNg86Ca0V/SoibA2XAy7YEc8csapz4QYzvFNmsKk="
$URL_SITE = "https://intercorpretail.sharepoint.com/sites/AppsCorporativas/er"

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
    [Int]$PageSize = 2000
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
            Write-Log "Error in item Id: $($a.Id) from List: $($ListName) -> Field $($_) is null." -LogFile $LOG_ERRORS
            $null
          }
        }) -join $Delimiter) -LogFile $LogFile
  }
}

########################
#### EXPORTED FILES ####
########################
$LOG_ERRORS = "errors.log"
$LOG_LIQUIDACIONES = "liquidaciones.csv"
$LOG_LIQUIDACION_DETALLE = "liquidacionDetalle.csv"
$LOG_LIQUIDACION_CABECERA = "liquidacionCabecera.csv"
$LOG_RESULTADOS = "resultados.csv"

####################
#### LIST NAMES ####
####################
$LISTA_LIQUIDACIONES = "Liquidaciones"     
$LISTA_LIQUIDACION_DETALLE = "Lista: Liquidacion Detalle"
$LISTA_LIQUIDACION_CABECERA = "Lista: Liquidacion Cabecera"

################
#### FIELDS ####
################
$FIELDS_LIQUIDACIONES = "ID", "FileRef", "Detalle"
$FIELDS_LIQUIDACION_DETALLE = "ID", "Solicitud", "Liquidacion", "Title"
$FIELDS_LIQUIDACION_CABECERA = "ID", "Solicitud", "AnioSAP", "Sociedad", "Title"

Connect-PnPOnline -Url $URL_SITE -ClientId $CLIENT_ID -ClientSecret $CLIENT_SECRET

######################
#### EXPORT LISTS ####
######################

Export-List -LogFile $LOG_LIQUIDACIONES -ListName $LISTA_LIQUIDACIONES -Fields $FIELDS_LIQUIDACIONES -Delimiter ";"
# Export-List -LogFile $LOG_LIQUIDACION_DETALLE -ListName $LISTA_LIQUIDACION_DETALLE -Fields $FIELDS_LIQUIDACION_DETALLE -Delimiter ";"
# Export-List -LogFile $LOG_LIQUIDACION_CABECERA -ListName $LISTA_LIQUIDACION_CABECERA -Fields $FIELDS_LIQUIDACION_CABECERA -Delimiter ";"

$itemsLiquidaciones = Import-Csv $LOG_LIQUIDACIONES -Delimiter ";"
$itemsLiquidacionDetalle = Import-Csv $LOG_LIQUIDACION_DETALLE -Delimiter ";"
$itemsLiquidacionCabecera = Import-Csv $LOG_LIQUIDACION_CABECERA -Delimiter ";"

# $hashTableLiquidaciones = @{}
$hashTableLiquidacionDetalle = @{}
$hashTableLiquidacionCabecera = @{}


foreach ($item in $itemsLiquidacionDetalle) {
  $hashTableLiquidacionDetalle[$item.ID] = $item
}

foreach ($item in $itemsLiquidacionCabecera) {
  $hashTableLiquidacionCabecera[$item.Title] = $item
}

Write-Log "ID;FileRef;ID_CABECERA;Solicitud;AnioSAP;Sociedad;Title" -LogFile $LOG_RESULTADOS
foreach ($item in $itemsLiquidaciones) {
  $itemHashTableLiquidacionDetalle = $hashTableLiquidacionDetalle[$item.Detalle]
  if ($null -ne $itemHashTableLiquidacionDetalle) {
    $itemHashTableLiquidacionCabecera = $hashTableLiquidacionCabecera[$itemHashTableLiquidacionDetalle.Liquidacion]
    if ($null -ne $itemHashTableLiquidacionCabecera) {
      Write-Log "$($item.ID);$($item.FileRef);$($itemHashTableLiquidacionCabecera.PSObject.Properties.Value -join ";")" -LogFile $LOG_RESULTADOS
    }
  }
}

# Write-Output $itemsLiquidaciones
# Write-Output $itemsLiquidacionDetalle
# Write-Output $itemsLiquidacionCabecera

