' Quelques fonctions de Network

   
Set net = WScript.CreateObject("WScript.Network")
Dim str

str = net.username
WScript.StdOut.WriteLine("Nom d'utilisateur actuel: " & str)

str = net.userdomain
WScript.StdOut.WriteLine("Domaine actuel: " & str)

str = net.computername
WScript.StdOut.WriteLine("Computer name: " & str)


' IP ADDRESS
' https://docs.microsoft.com/en-us/windows/win32/cimwin32prov/win32-networkadapterconfiguration
'dim NIC1, Nic, StrIP

'Set NIC1 =     GetObject("winmgmts:").InstancesOf("Win32_NetworkAdapterConfiguration")

'For Each Nic in NIC1
'    if Nic.IPEnabled then
'        StrIP = Nic.IPAddress(0)
'        WScript.StdOut.WriteLine "IP Address:  "&StrIP 
'    End if
'Next

' PING
ip0 = "192.168.178.1"
set oShell = WScript.CreateObject("WScript.Shell")
set ping0 = oShell.Exec("ping -n 3 -w 2000 " & ip0) ' exec execute une commande en arriere plan
str0 = ping0.StdOut.ReadAll  ' le retour de la commande peut etre lue ainsi

WScript.StdOut.WriteLine "Ping : " & str0
