 #Get-ExecutionPolicy -List
 #PERMITIR LA EJECUCIÓN DE SCRIPT
 Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
 #Form
 Add-Type -AssemblyName System.Windows.Forms


Set-Location -Path C:\
$Directorio="PS-Script"
$Dir_Result="Resultados"
$Dir_XLS="Procesar"
MD $Directorio -Force
CD $Directorio
MD $Dir_Result -Force
MD $Dir_XLS -Force
$PachResult="C:\"+$Directorio+"\"+$Dir_Result+"\"
$PachProcesar="C:\"+$Directorio+"\"+$Dir_XLS+"\"
cls
$FileBrowser = New-Object System.Windows.Forms.OpenFileDialog -Property @{ 
    InitialDirectory = $PachProcesar #[Environment]::GetFolderPath('Desktop') 
    Filter = 'SpreadSheet (*.xlsx)|*.xlsx' #|Documents (*.docx)|*.docx
}
$null = $FileBrowser.ShowDialog()

if(-Not $FileBrowser.CheckFileExists){ #-ne ""
write-host "ARCHIVO NO EXISTE"
exit 5
}

if(-Not $FileBrowser.CheckFileExists){
write-host "ARCHIVO NO EXISTE"
exit 5
}

#Ver Información de archivo
#$FileBrowser.FileName

#Almacenar Ruta y Nombre de Archivo
$RutaArchivoXls=$FileBrowser.FileName



#Ver información y permisos de archivo
#(Get-Acl $RutaArchivoXls).Access

#VERIFICAR ARCHIVO ESTE CERRADO PARA ESCRIBIR
Try { 
[io.file]::OpenWrite($RutaArchivoXls).close() 
}
 Catch { 
 Write-Warning "ARCHIVO .::: $RutaArchivoXls ::.. ESTA ABIERTO O NO EXISTE, VERIFICA!"
 
 exit 5
 #Write-Warning "Unable to write to output file $FileBrowser.FileName" 
 }

 #Get-ExecutionPolicy -List
 #PERMITIR LA EJECUCIÓN DE SCRIPT
 Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
 Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser


CLS
Write-Host " DESARROLLADO POR YARED CORDERO GHOST!"

############### MODULOS ###########################
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
############### FIN DE MODULOS ###########################


############### MICROSOFT GRAPH ##########################
    Connect-MgGraph -Scopes "User.Read.All", "Group.ReadWrite.All" -UseDeviceAuthentication
    #CONEXION MICROSOFT GRAPH
    $clientID = ""
    $ClientSecret = ConvertTo-SecureString -String "" -AsPlainText -Force
    $TenantName = "" 
    $TenantID = "" 
    $CertificatePath = ""

    $Params = @{
        ClientId = $clientID
        ClientSecret = $ClientSecret
        TenantId = $TenantName
        ForceRefresh = $true
        ErrorAction = 'SilentlyContinue'
    
    }
    $TokenAccess = (Get-MsalToken @Params).AccessToken

############### FIN MICROSOFT GRAPH ######################



Function NombrePropio{ 
        param ($NombrePropio)

        $TextInfo = (Get-Culture).TextInfo
        $NombrePropio = $TextInfo.ToTitleCase($NombrePropio)
        return $NombrePropio
}


Function MODIFICAR_USUARIO{
$AccountEnabled=$true
if(-Not $AccountEnabled){
Write-Warning "..:: CUENTA BLOQUEADA ::.." -WarningAction Continue
}else{
Write-Warning "..:: CUENTA DESBLOQUEADA ::.." -WarningAction Continue
}
    $userId = $BusquedaReg.Id
    $password=$PasswordUser
    $userPrincipalName=$CorreoUser.Trim()
    $displayName=$ApellidosyNombres
    $Nombres=NombrePropio $NombresUser
    $Apellidos=NombrePropio $ApellidosUser
    $newUserPrincipalName = $userPrincipalName
    $ForcePasswordNextSignIn = $true
    $TipoUsuario = $TipoUser.ToUpper()
    $Dominio=$newUserPrincipalName.Split("@")[1]
    $EmpresaDominio=$Dominio.Split(".")[0].ToUpper()
    $date = Get-Date -UFormat "%Y-%m-%d"
    
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
    "jobTitle"=$TipoUsuario
    "CompanyName"=$EmpresaDominio
    "Department"="$TipoUsuario-$EmpresaDominio"
    "EmployeeType"=$TipoUsuario
    "EmployeeHireDate"=$date
    "PostalCode"="07001"

    #PasswordProfile = @{ 
        #ForceChangePasswordNextSignIn = $ForcePasswordNextSignIn 
        #Password = $password 
    #} 
}

Update-MgUser -UserId $userId -BodyParameter $params 
Write-Host "---> DATOS ACTUALIZADOS " $CorreoUser.Trim() " A " $password -BackgroundColor DarkBlue -ForegroundColor Yellow
return $true
}





Function CREAR_USUARIO{
  Write-Host "---> CREANDO USUARIO " $CorreoUser.Trim() -BackgroundColor DarkRed -ForegroundColor white

    $AccountEnabled=$true
        if(-Not $AccountEnabled){
            Write-Warning "..:: CUENTA BLOQUEADA ::.." -WarningAction Continue
        }else{
            Write-Warning "..:: CUENTA DESBLOQUEADA ::.." -WarningAction Continue
        }
    
    $password=$PasswordUser
    $userPrincipalName=$CorreoUser.Trim()
    $displayName=$ApellidosyNombres
    $Nombres=NombrePropio $NombresUser
    $Apellidos=NombrePropio $ApellidosUser
    $newUserPrincipalName = $userPrincipalName
    $ForcePasswordNextSignIn = $true
    $TipoUsuario = $TipoUser.ToUpper()
    $Dominio=$newUserPrincipalName.Split("@")[1]
    $EmpresaDominio=$Dominio.Split(".")[0].ToUpper()
    $date = Get-Date -UFormat "%Y-%m-%d"


$CreateUserBody = @{
"userPrincipalName"=$userPrincipalName
"displayName"=$displayName
"mailNickname"=$newUserPrincipalName.Split("@")[0]
"accountEnabled"=$AccountEnabled
"UsageLocation"="PE"
"surname"= $Apellidos
"givenName"=$Nombres
"preferredLanguage"= "es-ES"
"officeLocation"="EST"
"jobTitle"=$TipoUsuario
"CompanyName"=$EmpresaDominio
"Department"="$TipoUsuario-$EmpresaDominio"
"EmployeeType"=$TipoUsuario
"PostalCode"="07001"
        
"passwordProfile"= @{
        "forceChangePasswordNextSignIn" = $ForcePasswordNextSignIn
        "forceChangePasswordNextSignInWithMfa" = $false
        "password"=$password
    }
 }

   $CreateUserUrl = "https://graph.microsoft.com/v1.0/users"
   $CreateUser =  Invoke-RestMethod  -Uri $CreateUserUrl -Headers @{Authorization = "Bearer $($TokenAccess)" }  -Method Post -Body $($CreateUserBody | convertto-json) -ContentType "application/json"
    Start-Sleep -s 2

    if($CreateUser.Id)
    {
     Update-MgUser -UserId $CreateUser.Id -EmployeeHireDate $date
     ASIGNAR_LICENCIAS -UserId $CreateUser.Id
     return $true
     #REPORTE_USUARIO -EstadoRegistro "NUEVO"
    }else{
    return $false
    Write-Warning "ERROR EN EL REGISTRO NUMERO: $Conteo "  -WarningAction Continue
    #Write-Host "vuelva a ejecutar con el parametro Registro Ejemplo: $NombreArchivo -Registro $Conteo " -BackgroundColor DarkRed -ForegroundColor white
    #break
    }
    
 }






 Function ASIGNAR_LICENCIAS{
 param ($UserId)
    
    Write-Host "---> ASIGNANDO LICENCIA " $CorreoUser.Trim()  -BackgroundColor DarkBlue -ForegroundColor Yellow

$Licenias = $LicenciasUser -split "/"
if($Licenias)
    {
        #ASIGNAR LICENCIAS
        For ($i=0; $i -le ($Licenias.count-1); $i++) {
        $license = Get-MgSubscribedSku | Where-Object {$_.SkuPartNumber -eq $Licenias[$i]}
        $UsuarioLicencia = Set-MgUserLicense -UserId $UserId -AddLicenses @{SkuId = ($license.SkuId)} -RemoveLicenses @()
            if($UsuarioLicencia)
            {
            Write-Host "-----> " $CorreoUser.Trim()  " LICENCIA: " $Licenias[$i] -BackgroundColor DarkGreen -ForegroundColor white
            }
        }
    }else{
        #QUITAR LICENCIAS
    }


}


Function ESCRIBIR_RESULTADO{
 param ($Respuesta)

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




########################### EXCEL APLICACION ########################
#EXCELOBJ APP
    $ExcelObj = New-Object -comobject Excel.Application

#ABRIR ARCHIVO XLS
    $ExcelObj.visible=$true
    $ExcelWorkBook = $ExcelObj.Workbooks.Open($RutaArchivoXls)
    
#VER HOJAS DE TRABAJO
    #$ExcelWorkBook.Sheets| fl Name, index

#ABRIR UNA HOJA DEL LIBRO
    #$ExcelWorkSheet = $ExcelWorkBook.Sheets.Item("PROCESAR")
    $ExcelWorkSheet = $ExcelWorkBook.Worksheets.Item(1)

#Inicio en el registro 2 por los campos
    $InicioRegistro=2

#COLUMNAS
    $Col_Correo=1
    $Col_Nombres=2
    $Col_Apellidos=3
    $Col_Password=4
    #$Col_Licencias=5
    $Col_TipoUser=5
    $Col_Estado=6
    $Col_Fecha=7
    $Col_Observacion=8
    $Col_proceso=9

#Cantidad de Registros:
    $FinRegistro = ($ExcelWorkSheet.usedRange.rows).count

################### FIN EXCEL APLICACION ########################




$Conteo=1
################### RECORRER XLSX ########################
for ($ReadRows=2; $ReadRows -lt ($FinRegistro+1); $ReadRows++)
{
   
    #Extraccion de datos
        $CorreoUser=$ExcelWorkSheet.Columns.Item($Col_Correo).Rows.Item($ReadRows).Text.Trim()
        $NombresUser=$ExcelWorkSheet.Columns.Item($Col_Nombres).Rows.Item($ReadRows).Text
        $ApellidosUser=$ExcelWorkSheet.Columns.Item($Col_Apellidos).Rows.Item($ReadRows).Text
        $PasswordUser=$ExcelWorkSheet.Columns.Item($Col_Password).Rows.Item($ReadRows).Text
        #$LicenciasUser=$ExcelWorkSheet.Columns.Item($Col_Licencias).Rows.Item($ReadRows).Text
        $TipoUser=$ExcelWorkSheet.Columns.Item($Col_TipoUser).Rows.Item($ReadRows).Text.Trim()
        $Proceso=$ExcelWorkSheet.Columns.Item($Col_proceso).Rows.Item($ReadRows).Text.Trim()
    
    #TIPO DE LICENCIAS ASIGANADAS
        if($TipoUser -eq "ESTUDIANTE"){
            $LicenciasUser="STANDARDWOFFPACK_STUDENT/OFFICESUBSCRIPTION_STUDENT"
        }ELSEif($TipoUser -eq "DOCENTE" -or $TipoUser -eq "ADMINISTRATIVO"){
            $LicenciasUser="STANDARDWOFFPACK_FACULTY/VISIOCLIENT_FACULTY/OFFICESUBSCRIPTION_FACULTY"
        }ELSE{
            $LicenciasUser=""
        }
    

    #FORMATO NOMBRE PROPIO
        $InputNombre=[Text.Encoding]::utf8.GetString([Text.Encoding]::GetEncoding('Cyrillic').GetBytes($NombresUser.ToLower()))
        $InputApellido=[Text.Encoding]::utf8.GetString([Text.Encoding]::GetEncoding('Cyrillic').GetBytes($ApellidosUser.ToLower()))
        $ApellidosyNombres=NombrePropio "$InputNombre $InputApellido"

 if($Proceso -ine "PROCESADO" -and -NOT $CorreoUser -eq ""){
      Write-Host $Conteo " ::: INICIANDO TRABAJO CON: " $CorreoUser $ApellidosyNombres -BackgroundColor Yellow -ForegroundColor DarkBlue
      #Validar si Existe Correo
      $BusquedaReg=Get-MgUser -ConsistencyLevel eventual -Count userCount -Search "mail:$CorreoUser"
         IF($BusquedaReg){
            
            #Si Existe Actualizar Datos
                if(MODIFICAR_USUARIO){
                    #Usuario Actualizado
                    ASIGNAR_LICENCIAS -UserId $BusquedaReg.Id
                        $date = Get-Date -UFormat "%d/%m/%Y %H:%M:%S" #  "%Y%m%dT%H%M%S"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado) = 'ACTUALIZADO'
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Fecha) = $date
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Observacion) = "" 
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_proceso) = "PROCESADO"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Font.Bold = $true #negrita
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Interior.ColorIndex = 10 #Color Rojo
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).font.ColorIndex = 2 #Color Rojo
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).HorizontalAlignment  = -4108 #Centrado


                }else{
                    #ERROR Usuario no Actualizado
                        $date = Get-Date -UFormat "%d/%m/%Y %H:%M:%S" #  "%Y%m%dT%H%M%S"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado) = 'ERROR'
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Fecha) = $date
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_proceso) = ""
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Observacion) = "NO SE PUEDO ACTUALIZAR, REVISAR DATOS"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Font.Bold = $true #negrita
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Interior.ColorIndex = 3 #Color Rojo
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).HorizontalAlignment  = -4108 #Centrado
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).font.ColorIndex = 2 #Color Rojo

                }
         }else{
            #Si no Existe Crear Correo Nuevo
                if(CREAR_USUARIO){
                    #Usuario Nuevo Creado
                        $date = Get-Date -UFormat "%d/%m/%Y %H:%M:%S" #  "%Y%m%dT%H%M%S"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado) = 'NUEVO'
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Fecha) = $date
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Observacion) = "" 
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_proceso) = "PROCESADO"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Font.Bold = $true #negrita
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Interior.ColorIndex = 3 #Color Rojo
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).font.ColorIndex = 2 #Color Rojo
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).HorizontalAlignment  = -4108 #Centrado

                }else{
                    #ERROR Creando Usuario
                        $date = Get-Date -UFormat "%d/%m/%Y %H:%M:%S" #  "%Y%m%dT%H%M%S"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado) = 'ERROR'
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Fecha) = $date
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_proceso) = ""
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Observacion) = "NO SE PUEDE CREAR USUARIO"
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Font.Bold = $true #negrita
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).Interior.ColorIndex = 3 #Color Rojo
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).HorizontalAlignment  = -4108 #Centrado
                        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Estado).font.ColorIndex = 2 #Color Rojo
                }
         }

        
        $Conteo++
         Write-Host "-----------------------------"
         Write-Host ""
         Write-Host ""
    }else{
        
        $ExcelWorkSheet.Cells.Item($ReadRows,$Col_Observacion) = "DATOS INCOMPLETOS"
    }
}
################### FIN RECORRER XLS #####################



$date = Get-Date -UFormat "%Y-%m-%d_%H_%M_%S_"

#Reescribo Información en el Archivo Xlsx
    $ExcelObj.DisplayAlerts = $false
    $ExcelWorkBook.SaveAs($RutaArchivoXls)
    $ExcelWorkBook.SaveAs($PachResult+"Resultado"+$date)
    $ExcelWorkBook.close($false)
    $ExcelObj.Quit()
    [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelObj)
    Remove-Variable ExcelObj

#ABRIR ARCHIVO PROCESADO    
    ii -Path $RutaArchivoXls

#FIN



