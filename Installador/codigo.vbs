'Uninstall


On Error Resume Next
Dim objExcel
Dim objaddin
On Error Resume Next

Set objExcel = CreateObject("Excel.Application")
For i = 1 to objExcel.Addins.Count
    Set objAddin = objExcel.Addins.item(i)
    If objAddin.Name = "UDFHaciendaCR-AddIn.xll" Then
       objAddin.Installed = False
    End If
Next

objExcel.Quit
Set objExcel = Nothing



' Install

On Error Resume Next
Dim oXL, oAddin
Set oXL = CreateObject("Excel.Application")
oXL.Workbooks.Add
Set oAddin = oXL.AddIns.Add(Session.property("CustomActionData") & "UDFHaciendaCR-AddIn.xll", False)
oAddin.Installed = True
oXL.Quit
Set oXL = Nothing