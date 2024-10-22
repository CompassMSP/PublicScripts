#This script finds computers that have been left in the default "Computers" OU
if ((Get-WmiObject Win32_ComputerSystem).domainrole -eq 5) {
    #computer is a PDC
    try {
        Import-Module ActiveDirectory

        $DefaultComputersDN = "CN=Computers,$(Get-ADDomain | Select-Object -ExpandProperty DistinguishedName)"

        $ComputersInDefaultContainer = Get-ADComputer -SearchBase $DefaultComputersDN -Filter *

        if ($ComputersInDefaultContainer.count -eq 0) {
            #exit
        }
        else {
            return "The Following computers were found in the Computers container: $($ComputersInDefaultContainer.Name -join ',')"
        }
    }
    catch {
        exit
    }
}
else {
    #not a PDC
    exit
}