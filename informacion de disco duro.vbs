set WshShell = WScript.CreateObject("WScript.Shell")
Dim Msg, Style, Title, Response, MyString
Msg = "info disco duro"    
Style = vbOkCancel  
Title = "disco duro"

Response = MsgBox(Msg, Style, Title)
If Response = vbOk Then
	strComputer = "."
On Error Resume Next
Set objWMIService = GetObject("winmgmts:\\" & strComputer & "\root\CIMV2") 
Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_DiskDrive")
for each objItem in colItems

	MsgBox "marca y modelo: " & objItem.model & vbNewLine & "descripcion: " & objItem.Description & vbNewLine & "numero de serie: " & objItem.SerialNumber & vbNewLine & "estado: " &objItem.Status & vbNewLine & "tipo de interfaz: " & objItem.InterfaceType & vbNewLine & "particiones: " & objItem.partitions & vbNewLine & "tamaño: " & objItem.Size/1024/1024/1024 & "Gb"
	

next
quit = "ok"
else
  MyString = "cancel"
end if