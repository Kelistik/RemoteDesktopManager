#Justin Morrissette - Add servers from an SNB excel spreadsheet into Remote Desktop Manager

#7/15/2018 - V 0.1 - Initial Script Creation - JMM
#9/13/2018 - V 0.2 - Added disabled param as well as the IP check section - JMM
#9/18/2018 - V 0.3 - Added username/password/domain section - JMM
#10/12/2018 - V 0.4 - Can be run in normal powershell and not in RDM (from Nate Adrian) - JMM
#12/18/2018 - V 0.5 - Added functionality for new SNB method (from Nate Adrian) - JMM

#README Section:
#If it runs but no servers are added, check if the Status (first column) says enable. It must say enable for the server to be read
#USE (from the directory of the script): & '.\Remote Desktop Manager.ps1' -client clientcode -filepath "path to the SNB"
#USE (from the directory of the script): You must have the "Cloud Services Colleague - Team Docs" folder from box synced in order to run the script
#You can also pass the parameter -disabled if you want all of the disabled servers output

param (
    [parameter(Mandatory=$true)]$client,
    [string]$filepath = "",
    [string]$username = "",
    [string]$password = "",
    [string]$domain = "",
    [switch]$disabled
)

#V 0.4 - Can be run in normal powershell and not in RDM
$module = Get-ChildItem "${env:ProgramFiles(x86)}\Devolutions" -recurse | Where-Object {$_.Name -eq "RemoteDesktopManager.PowerShellModule.psd1"}|select-object -First 1

try {
   Import-Module $module.FullName
   }

catch {
}

if ($filepath -NE ""){
    if (-not (test-path -path "$filepath")){
        Write-Output "SNB for $client does not exist at $filepath."
        exit
    }
}

#if ($password -EQ "" -AND $username -NE ""){
#    $answer = Read-Host -prompt "Did you want to enter a password? (Y or N)"
#    if ($answer -EQ "Y"){
#        $password = Read-Host -assecurestring "Enter Password"
#    }
#}

Write-Output $password

if ($filepath -EQ ""){
    if (test-path -path "$home\Documents\Documentation\SNBs\0020-Ellucian CS Server Network Build-AWS-$client.xlsx"){
        $filepath = "$home\Documents\Documentation\SNBs\0020-Ellucian CS Server Network Build-AWS-$client.xlsx"
    }
    if (test-path -path "$home\Box Sync\Cloud Services Colleague - Team Docs\SNBs\0020-Ellucian CS Server Network Build-AWS-$client.xlsx"){
        $filepath = "$home\Box Sync\Cloud Services Colleague - Team Docs\SNBs\0020-Ellucian CS Server Network Build-AWS-$client.xlsx"
    }
    if (-not(test-path -path "$home\Documents\Documentation\SNBs\0020-Ellucian CS Server Network Build-AWS-$client.xlsx") -AND -not(test-path -path "$home\Box Sync\Cloud Services Colleague - Team Docs\SNBs\0020-Ellucian CS Server Network Build-AWS-$client.xlsx")){
        Write-Output "SNB for $client does not exist in a valid location."
        exit
    }
}

New-RDMSession -type group -Name "$client" -SetSession
Update-RDMUI
New-RDMSession -type group -Name "$client\PROD" -SetSession
New-RDMSession -type group -Name "$client\NON-PROD" -SetSession
Update-RDMUI

$statusColumn = 1
$serverRoleColumn = 2
$serverNameColumn = 3
$folderColumn = 4
$internalNameColumn = 6
$serverIPColumn = 7
$operatingColumn = 40
$rowCount = 0
$row = 2

$objExcel = New-Object -ComObject Excel.Application

$workBook = $objExcel.Workbooks.Open($filepath)
$workSheet = $workBook.sheets.item("EC2-SNB")

$serverRole = $worksheet.cells.item($statusColumn, $serverRoleColumn).value2

#Begin main loop here
while ($serverRole -NE "totals" -AND $rowCount -LT 10){
    $status = $worksheet.cells.item($row, $statusColumn).value2
    $serverRole = $worksheet.cells.item($row, $serverRoleColumn).value2
    #Check if the server is marked enabled
    if ($status -EQ "enable"){
        $serverName = $worksheet.cells.item($row, $serverNameColumn).value2
        $folderName = $worksheet.cells.item($row, $folderColumn).value2
        $serverIP = $worksheet.cells.item($row, $serverIPColumn).value2
        #V 0.5 - If the Server IP is missing from the SNB, the internal hostname will be used to find it
        if ($serverIP -EQ $null) {
            $intServerName = $worksheet.cells.item($row, $internalNameColumn).value2
            $serverIP =  ([System.Net.Dns]::GetHostAddresses($intServerName)).IPAddressToString
        }
        #This line is creating the name of the tab so if you want to change it, this is where to go
        $serverRole = $serverRole + " - " + $serverName
        $serverIP = $serverIP.trim()
        #Check if the server is marked production or not
        if ($folderName -EQ "PROD"){
            $serverRole = $serverRole -replace 'PROD ',''
        }
        if ($folderName -EQ "NON-PROD"){
            $serverRole = $serverRole -replace 'Non-PROD ',''
        }
        $operatingSystem = $worksheet.cells.item($row, $operatingColumn).value2
        $error.clear()
        #Check if the RDM Session exists or not, if it does exist then it will error and skip the catch block
        try {Get-RDMSession -Name "$serverRole" | Out-Null}
        #Create the session if it does not exist by checking operating system and then organizing the sessions
        catch { 
            if ($operatingSystem -EQ "Windows"){
                if ($folderName -EQ "PROD"){
                    New-RDMSession -Name "$serverRole" -Type "RDPConfigured" -Host "$serverIP" -Group "$client\PROD" -SetSession
                }
                if ($folderName -EQ "NON-PROD"){
                    New-RDMSession -Name "$serverRole" -Type "RDPConfigured" -Host "$serverIP" -Group "$client\NON-PROD" -SetSession
                }
            }
            if ($operatingSystem -NE "Windows"){
                if ($folderName -EQ "PROD"){
                    New-RDMSession -Name "$serverRole" -Type "SSHShell" -Host "$serverIP" -Group "$client\PROD" -SetSession
                }
                if ($folderName -EQ "NON-PROD"){
                    New-RDMSession -Name "$serverRole" -Type "SSHShell" -Host "$serverIP" -Group "$client\NON-PROD" -SetSession
                }   
            }
        }
        #If the session exists, check that the IP is correct
        Update-RDMUI
        $getName = Get-RDMSession -Name $serverRole
        if (-Not ($error)){
            if ($getName.Host -NE $serverIP){
                $getName.Host = "$serverIP"; Set-RDMSession $getName
                Write-Output "Changed $serverRole IP."
            }
            if ($getName.Host -EQ $serverIP){
            Write-Output "$serverRole already exists."
            }
        }
        if ($username -NE "" -AND (-Not(Get-RDMSessionProperty -ID $getName.ID -Property "UserName"))){
            Set-RDMSessionUsername -ID $getName.ID -UserName $username
            Set-RDMSessionPassword -ID $getName.ID -Password (ConvertTo-SecureString -AsPlainText "$password" -Force)
            Write-Output "Username and Password set for $serverRole."
        }
        if ($domain -NE "" -AND $operatingSystem -EQ "Windows" -AND (-Not(Get-RDMSessionProperty -ID $getName.ID -Property "Domain"))){
            Set-RDMSessionDomain -ID $getName.ID -Domain $domain
            Write-Output "Domain set for $serverRole."
        }
    }
    #Only outputs if you run script with -disabled
    if ($status -NE "enable" -AND $disabled){
        Write-Output "$serverRole is disabled."
    }
    if([string]::IsNullOrEmpty($serverRole)){
        $rowCount = $rowCount + 1
    }
    $row = $row + 1
}

#Close excel so you don't have to open it as read only
$objexcel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($objexcel)
Remove-Variable objexcel

#Update the UI of RDM so the changes show, may still need to refresh the navigation tab
Update-RDMUI

exit