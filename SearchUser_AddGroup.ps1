
$global:OU_VPN="OU=Wiener_VPN,DC=WIENERGROUP,DC=COM"

$UserVPNSearch = Read-Host -Prompt "Escribe el nombre de usuario:"
$UserVPNSearch="*$UserVPNSearch*" 
$AD_Search_Like=@(Get-ADUser -Filter {name -like $UserVPNSearch})
$global:AD_SeachCount=0
if($AD_Search_Like.count -ge 1){
    $AD_Search_Like | Format-Table @{Label="NUM"; Expression={"[ $global:AD_SeachCount ]"; $global:AD_SeachCount++} },Name,UserPrincipalName,SamAccountName,Enabled -AutoSize
    while(($AD_UserSelect = Read-Host -Prompt "Seleciona Usuario AD ::") -ige ($global:AD_SeachCount)){}
    CLS
    $ADUserSelect=$AD_Search_Like[($AD_UserSelect-1)].Name
    $ADUserAccountName=$AD_Search_Like[($AD_UserSelect-1)].SamAccountName
    Write-Host -ForegroundColor Yellow "USUARIO :: " $ADUserSelect " ID::" $ADUserAccountName
    $global:AD_SeachCount=0

    $global:VpnCount=0
    $Grupos_VPN=@(get-adobject -Filter {ObjectClass -eq "group"} -SearchBase $global:OU_VPN) 

    if($Grupos_VPN.count -ge 1){
        $Grupos_VPN | Format-Table @{Label="NUM"; Expression={"[ $global:VpnCount ]"; $global:VpnCount++} },Name -AutoSize
        while(($VPN_List = Read-Host -Prompt "SELECIONE EL GRUPO VPN::") -ige ($global:VpnCount)){}
        $VPNSelect=$Grupos_VPN[($VPN_List-1)].Name
    
    
        $global:VpnCount=0
        CLS

        $title    = " CONFIRMA LA INFORMACIÓN `n` ----------------------- `n` "
        $question = "Agregar Usuario $ADUserSelect al grupo VPN $VPNSelect" 
        Write-Host -ForegroundColor Yellow "USUARIO  :: " $ADUserSelect " ID::" $ADUserAccountName
        Write-Host -ForegroundColor Yellow "Grupo VPN:: " $VPNSelect
      
        $choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes ACEPTAR'))
        $choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No CANCELAR'))

        $decision = $Host.UI.PromptForChoice($title, $question, $choices, 0)
            if ($decision -eq 0) {
            Add-ADGroupMember -Identity $VPNSelect  -Members $ADUserAccountName
            Write-Warning  "✔ VPN $VPNSelect ASIGNADA A  $ADUserSelect " -WarningAction Continue               
            }else{
            Write-Host " ✖ PROCESO CANCELADO " -BackgroundColor DarkRed -ForegroundColor Yellow
            }


            ###########LISTA DE USUARIO EN LA VPN##########
            $ADGroupList = Get-ADGroup -Filter {name -eq $VPNSelect} | Select Name -ExpandProperty Name | Sort Name
                ForEach($Group in $ADGroupList)
                {
                  $global:CountMember=1
                  $members=Get-ADGroupMember -Identity $Group | Select Name, SAMAccountName | Sort
                      ForEach($member in $members)
                      {
                          IF($member.SAMAccountName -eq $ADUserAccountName){
                            Write-Host  ("    ✔" + $global:CountMember+","+$member.Name+","+$member.SAMAccountName+","+$Group.name) -ForegroundColor Blue              
                          }else{
                            Write-Host "     " $CountMember"," $member.Name"," $member.SAMAccountName"," $Group.name -ForegroundColor DarkGray
                          }
                      $global:CountMember++
                      }
                  $global:CountMember=1
                }
            ###########LISTA DE USUARIO EN LA VPN##########

        }

}ELSE{
}

