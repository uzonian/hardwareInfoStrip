strComputer = "."
Set objWMIService = GetObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")


Dim objFso, x, curDir, FullFileName


curDir = CreateObject("WScript.Shell").CurrentDirectory


FullFileName = curDir & "\System Information.csv"


Set objFso = CreateObject("Scripting.FileSystemObject")
Set x = objFso.OpenTextFile(FullFileName,8)

'x.WriteLine "User, System Name, System Manufacturer, System Model, Serial Number " 
Dim ComputerName, Manufacturer, ModelNumber, SerialNumber, UserName, AssetNum
Set colSettings = objWMIService.ExecQuery ("Select * from Win32_ComputerSystem")
For Each objComputer in colSettings 
    ComputerName = objComputer.Name 
    Manufacturer = objComputer.Manufacturer
    ModelNumber = objComputer.Model
Next

Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystemProduct") 
For Each objItem in colItems 
   SerialNumber = objItem.IdentifyingNumber 
Next

 UserName = InputBox("Name of User | Vacant | Conf Room | Public | Storage")
 AssetNum = InputBox("Asset Tag")

  MsgBox "User: " & UserName & vbNewLine _
  & "Asset Number: " & AssetNum & vbNewLine _
  & "Computer Name: " & ComputerName & vbNewLine _
  & "Manufacturer: " & Manufacturer & vbNewLine _
  & "Model Number: " & ModelNumber & vbNewLine _
  & "Serial Number: " & SerialNumber

 x.WriteLine  UserName & "," & AssetNum  & "," & ComputerName  & "," &  Manufacturer & "," & ModelNumber  & "," & SerialNumber
