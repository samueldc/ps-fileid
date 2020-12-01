New-VMSwitch -SwitchName “NATSwitch” -SwitchType Internal

Get-NetAdapter

New-NetIPAddress -IPAddress 192.168.5.1 -PrefixLength 24 -InterfaceIndex 16

New-NetNAT -Name “NATNetwork” -InternalIPInterfaceAddressPrefix 192.168.5.0/24