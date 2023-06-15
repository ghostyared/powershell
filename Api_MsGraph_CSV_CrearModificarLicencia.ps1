param ($Registro)
$NombreArchivo=$myInvocation.InvocationName
$salto=1
$global:CoteoInterno=0
$global:Nuevos=0
$global:Actualizados=0
$global:InicioReg=0
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser

$module="ImportExcel"
if (Get-Module -ListAvailable -Name $module) {
    #Write-Host "MODULO EXISTE"
    #Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose
}else{
    Install-Module ImportExcel -AllowClobber -Force
    Get-Module ImportExcel -ListAvailable | Import-Module -Force -Verbose

} 


$module="Microsoft.Graph"
if (Get-Module -ListAvailable -Name $module) {
    #Write-Host "MODULO EXISTE"
    Import-Module Microsoft.Graph.Users.Actions
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Identity.DirectoryManagement
} 
else {
    Write-Host "MODULO $module NO EXISTE Instalando, espere..."
    Install-Module Microsoft.Graph -Scope CurrentUser
    Import-Module Microsoft.Graph.Users.Actions
    Import-Module Microsoft.Graph.Users
    Import-Module Microsoft.Graph.Identity.DirectoryManagement
}

$module="MSAL.PS"
if (Get-Module -ListAvailable -Name $module) {
    #Write-Host "MODULO EXISTE"
} 
else {
    Write-Host "MODULO $module NO EXISTE Instalando, espere..."
    Install-Module MSAL.PS

}

$Directorio="TEMP"
$Dir_Result="Resultados"
Set-Location -Path C:\
MD $Directorio -Force
CD $Directorio
MD $Dir_Result -Force
CLS
# Required Powershell Module for certificate authorisation

#Connect-MgGraph -Scope Directory.Read.All
#Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All"
Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All" -UseDeviceAuthentication

Write-Host " DESARROLLADO POR YARED CORDERO GHOST!"
Write-Host ""
Write-Host ""

# Minimum Required API permission for execution 
# User.Read.Write.All
# Directory.Read.Write.All
# Directory.AccessAsUser.All
# User.ManageIdentities.All


#LICENCIAS ADQUIRIDAS
#Get-MgSubscribedSKU | Format-List


#CONEXION MICROSOFT GRAPH
$clientID = ""
$ClientSecret = ConvertTo-SecureString -String "" -AsPlainText -Force
$TenantName = "" 
$TenantID = "" 
$CertificatePath = "Cert:\CurrentUser\My\___"
#Import Certificate
#$Certificate = Get-Item $certificatePath

$path = "C:\"+$Directorio  
$File = $path+"\DATOS_USUARIOS.csv"
$Report = [System.Collections.Generic.List[Object]]::new()


#Request Token METODO 1
########################
$Params = @{
    ClientId = $clientID
    ClientSecret = $ClientSecret
    TenantId = $TenantName
    ForceRefresh = $true
    ErrorAction = 'SilentlyContinue'
    
}
$TokenAccess = (Get-MsalToken @Params).AccessToken

<#
#Request Token METODO 2
########################
$TokenResponse = Get-MsalToken -ClientId $ClientId -TenantId $TenantId -ClientCertificate $Certificate -ErrorAction SilentlyContinue
$TokenAccess = $TokenResponse.accesstoken
#>

<#
#Request Token METODO 3
########################
$TokenResponse=Get-MsalToken -clientID $clientID -clientSecret $clientSecret -tenantID $TenantID -ErrorAction SilentlyContinue
$TokenAccess = $TokenResponse.accesstoken
$TokenAccess
#>



Function LEERUsuarios
{
param (
        $IniciarRows
    )

$ReadF = Import-Csv -Path $File -Delimiter ";" -Encoding utf7 #ASCII
$Conteo = $IniciarRows
$global:InicioReg=$IniciarRows
    foreach ($RowsUser in ($ReadF | select -skip $IniciarRows))
    {
    $Conteo++
    $global:CoteoInterno++

    $InputNombre=[Text.Encoding]::utf8.GetString([Text.Encoding]::GetEncoding('Cyrillic').GetBytes($RowsUser.Nombre.ToLower()))
    $InputApellido=[Text.Encoding]::utf8.GetString([Text.Encoding]::GetEncoding('Cyrillic').GetBytes($RowsUser.Apellido.ToLower()))
    $ApellidosyNombres=NombrePropio "$InputNombre $InputApellido"
    Write-Host $Conteo " ::: INICIANDO TRABAJO CON: " $RowsUser.Correo.Trim() $ApellidosyNombres -BackgroundColor Yellow -ForegroundColor DarkBlue
    $USUARIO=$RowsUser.Correo.Trim()
    $Busqueda=Get-MgUser -ConsistencyLevel eventual -Count userCount -Search "mail:$USUARIO"
    IF($Busqueda)
        {
         #Write-Host " -- USUARIO EXISTE" $RowsUser.Correo.Trim() -ForegroundColor Yellow
         
         $DatosUsuario=Get-MgUser -UserId $RowsUser.Correo.Trim()
         MODIFICAR_USUARIO
         ASIGNAR_LICENCIAS -UserId $DatosUsuario.Id
         REPORTE_USUARIO -EstadoRegistro "ACTUALIZADO"
         $global:Actualizados++
           
        }
        Else
        {
        #Write-Host " -- USUARIO NO EXISTE" $RowsUser.Correo.Trim() -ForegroundColor Yellow
        CREAR_USUARIO
        $global++

       }

       Write-Host ""
       Write-Host ""
       
  }
}
Function REPORTE_USUARIO
{
 param (
        $EstadoRegistro
    )

if($RowsUser.Activo -eq "FALSE"){ $RptState="SI"}else{$RptState="NO" }

            #REPORTE DE SALIDA
            #-----------------------
                $obj = [PSCustomObject][ordered]@{
                    "ID" = $Conteo
                    "USUARIO" = $RowsUser.Correo.Trim()
                    "NOMBRE" = $ApellidosyNombres
                    "BLOQUEO" = $RptState
                    "PASSWORD" = $RowsUser.Password
                    "Licencias" = $RowsUser.Licencias
                    "ESTADO" = $EstadoRegistro
                }
                $report.Add($obj)
             #-----------------------

}

Function MODIFICAR_USUARIO
{
$AccountEnabled=$RowsUser.Activo

if($AccountEnabled -eq "FALSE"){
Write-Warning "cuenda desahabilidata" -WarningAction Continue
}
    $DatosUsuario=Get-MgUser -UserId $RowsUser.Correo.Trim()
    $userId = $DatosUsuario.id
    $password=$RowsUser.Password

$userPrincipalName=$RowsUser.Correo.Trim()
$displayName=$ApellidosyNombres
$Nombres=$RowsUser.Apellido
$Apellidos=$RowsUser.Nombre
$newUserPrincipalName = $userPrincipalName
$ForcePasswordNextSignIn = $RowsUser.ResetPassword

$params = @{
"userPrincipalName"=$userPrincipalName
"displayName"=$displayName
"mailNickname"=$newUserPrincipalName.Split("@")[0]
"accountEnabled"=$AccountEnabled
"UsageLocation"="PE"
"surname"= $Apellidos
"givenName"=$Nombres
"preferredLanguage"= "es-ES"
"officeLocation"="EST"
    PasswordProfile = @{ 
    ForceChangePasswordNextSignIn = $ForcePasswordNextSignIn 
    Password = $password 
    } 
}
 Write-Host "---> ACTUALIZANDO DATOS Y CONTRASEÑA " $RowsUser.Correo.Trim() " A " $password -BackgroundColor DarkBlue -ForegroundColor Yellow
Update-MgUser -UserId $userId -BodyParameter $params 


}
Function ASIGNAR_LICENCIAS
{

 param (
        $UserId
    )
    
    Write-Host "---> ASIGNANDO LICENCIA " $RowsUser.Correo.Trim()  -BackgroundColor DarkBlue -ForegroundColor Yellow

$Licenias = $RowsUser.Licencias -split "/"
if($Licenias)
    {
        #ASIGNAR LICENCIAS
        For ($i=0; $i -le ($Licenias.count-1); $i++) {
        $license = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq $Licenias[$i]}
        $UsuarioLicencia = Set-MgUserLicense -UserId $UserId -AddLicenses @{SkuId = ($license.SkuId)} -RemoveLicenses @()
            if($UsuarioLicencia)
            {
            Write-Host "-----> " $RowsUser.Correo.Trim()  " LICENCIA: " $Licenias[$i] -BackgroundColor DarkGreen -ForegroundColor white
            }
        }
    }else{
        #QUITAR LICENCIAS
    }


}

Function QUITAR_LICENCIA
{
param (
        $UserId
    )

    $params = @{
	addLicenses = @(
		@{
			disabledPlans = @(
				"11b0131d-43c8-4bbb-b2c8-e80f9a50834a"
			)
			skuId = "45715bb8-13f9-4bf6-927f-ef96c102d394"
		}
	)
	removeLicenses = @(
		"bea13e0c-3828-4daa-a392-28af7ff61a0f"
	)
}

Set-MgUserLicense -UserId $UserId -BodyParameter $params

}


Function CREAR_USUARIO
 {


 Write-Host "---> CREANDO USUARIO " $RowsUser.Correo.Trim() -BackgroundColor DarkRed -ForegroundColor white

$userPrincipalName=$RowsUser.Correo.Trim()
$displayName=$ApellidosyNombres
$password=$RowsUser.Password
$Nombres=$RowsUser.Apellido
$Apellidos=$RowsUser.Nombre
$newUserPrincipalName = $userPrincipalName
#"mobilePhone"="425-555-0101"
#$Nickname=$RowsUser.Correo.Trim()
$ForcePasswordNextSignIn = $RowsUser.ResetPassword


$CreateUserBody = @{
"userPrincipalName"=$userPrincipalName
"displayName"=$displayName
"mailNickname"=$newUserPrincipalName.Split("@")[0]
"accountEnabled"=$true
"UsageLocation"="PE"
"surname"= $Apellidos
"givenName"=$Nombres
"preferredLanguage"= "es-ES"
"officeLocation"="EST"
        
"passwordProfile"= @{
        "forceChangePasswordNextSignIn" = $ForcePasswordNextSignIn
        "forceChangePasswordNextSignInWithMfa" = $false
        "password"=$password
    }
 }

    $CreateUserUrl = "https://graph.microsoft.com/v1.0/users"

    
    $CreateUser =  Invoke-RestMethod  -Uri $CreateUserUrl -Headers @{Authorization = "Bearer $($TokenAccess)" }  -Method Post -Body $($CreateUserBody | convertto-json) -ContentType "application/json"
    #-ErrorAction SilentlyContinue
Start-Sleep -s 2

    if($CreateUser.Id)
    {
     ASIGNAR_LICENCIAS -UserId $CreateUser.Id
     REPORTE_USUARIO -EstadoRegistro "NUEVO"
    }else{
    Write-Warning "ERROR EN EL REGISTRO NUMERO: $Conteo "  -WarningAction Continue
    Write-Host "vuelva a ejecutar con el parametro Registro Ejemplo: $NombreArchivo -Registro $Conteo " -BackgroundColor DarkRed -ForegroundColor white
    break
    
    }

}

Function NombrePropio
 {
  param ($NombrePropio)

$TextInfo = (Get-Culture).TextInfo
$NombrePropio = $TextInfo.ToTitleCase($NombrePropio)
return $NombrePropio

 } 

if($Registro){
Write-Warning "*******SE ESTA INICIANDO DESDE EL REGISTRO: $Registro *******"  -WarningAction Continue
LEERUsuarios -IniciarRows ($Registro-$salto)

}else{
LEERUsuarios -IniciarRows 0
}

Write-Host "SE PROCESARON $global:CoteoInterno REGISTROS, SE INICIO EN EL REGISTRO $global:InicioReg CON $global:Nuevos NUEVOS Y $global:Actualizados ACTUALIZADOS " -BackgroundColor blue -ForegroundColor yellow

#$Report

$date = Get-Date -UFormat "%Y%m%dT%H%M%S"
#REPORTE DE RESUMEN 
#--------------------

$Filename = "Rpt_Result_"+$date+"_.xlsx" 
try {
#$report | Export-CSV -path ($path+"\"+ $Dir_Result+"\" + $filename) -Delimiter ";"  -NoTypeInformation
$report | Export-Excel -Path ($path+"\"+ $Dir_Result+"\" + $filename)

} #-Encoding utf8
catch { "ERROR ACCESO NO PERMITIDO, cierre el archivo de resulem.csv" }
#$Report
Write-host "EXPORTANDO REPORTE $path\$Dir_Result\$Filename" -ForegroundColor Green
ii "$path\$Dir_Result\$Filename"

<#
#REPORTE DE LICENIAS 
#--------------------
$Filename = "Rpt_Licenses_"+$date+"_.csv" 

$Report = [System.Collections.Generic.List[Object]]::new()
$licenses = get-MgSubscribedSku
Foreach ($license in $licenses){

    $sku = (get-MgSubscribedSku -SubscribedSkuId $license.id).SkuPartNumber
    $licensecount = (get-MgSubscribedSku -SubscribedSkuId $license.id -Property PrepaidUnits | select-object -expandproperty prepaidunits).enabled
    $usedlicenses = (get-MgSubscribedSku -SubscribedSkuId $license.id).ConsumedUnits
    
    #REPORTE DE SALIDA
    #-----------------------
    $obj = [PSCustomObject][ordered]@{
        "License SKU" = $sku
        "Total Licenses" = $licensecount
        "Used Licenses" = $usedlicenses
    }
    $report.Add($obj)
    #-----------------------
}


#$Report
try {$report | Export-CSV -path ($path + "\" + $filename) -Delimiter ";" -Encoding ASCII -NoTypeInformation}
catch { "ERROR ACCESO NO PERMITIDO, cierre el archivo de licenses.csv" }
Write-host "EXPORTANDO REPORTE $path\$Filename" -ForegroundColor Green
#>
