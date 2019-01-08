Import-module activedirectory
$OU = "OU=Computers,DC=domain,DC=local"
Get-ADOrganizationalUnit -filter * -SearchBase $ou | % {
$Count = (get-adcomputer -filter * -searchbase $_.distinguishedname -SearchScope 1 | Measure).count
     If ($Count -ge 1){
     New-Object PSobject -Property @{
          OUDN = $_.DistinguishedName
          Count = $Count
           }
     }
