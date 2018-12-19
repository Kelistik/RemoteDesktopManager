# RemoteDesktopManager
This is to read from an SNB and update the Remote Desktop Manager application successfully

#If it runs but no servers are added, check if the Status (first column) says enable. It must say enable for the server to be read
#USE (from the directory of the script): & '.\Remote Desktop Manager.ps1' -client clientcode -filepath "path to the SNB" -username "username" -domain "domain" -password "password"
#USE (from the directory of the script): You must have the "Cloud Services Colleague - Team Docs" folder from box synced in order to run the script
#You can also pass the parameter -disabled if you want all of the disabled servers output
