#OBTENER ULTIMOS CORREOS CREADOS
$date1="2023-06-15T00:00:00-00:00"
$date2="2023-06-15T23:59:59-00:00"
$CorreosNuevos=Get-MgUser -all -Filter "createdDateTime ge $date1 and createdDateTime le $date2" -Property CompanyName,CreatedDateTime,DisplayName,UserPrincipalName |  Select-Object CompanyName,@{Label="FECHA"; Expression={(Get-Date($_.createdDateTime)).ToString("dd/MM/yyyy H:mm:ss")}},DisplayName,UserPrincipalName
try {$CorreosNuevos | Export-CSV -path ("C:\temp\CORREOS_NUEVO.csv") -Delimiter ";" -Encoding ASCII -NoTypeInformation}
catch { "ERROR ACCESO NO PERMITIDO, cierre el archivo de resulem.csv" }
