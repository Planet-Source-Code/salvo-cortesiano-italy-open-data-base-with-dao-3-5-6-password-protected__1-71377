Attribute VB_Name = "Mod_OpenDB"
Option Explicit

Public cnAP As Database
Public rsUsers As Recordset

Public dbOpen As Boolean
Public Sub OpenDatabase(ByVal strFile As String, Optional strPassWord As String)
    
    On Error GoTo ErrorHandler
    
    If Len(strPassWord) > 0 Then Set cnAP = DBEngine.Workspaces(0).OpenDatabase(strFile, False, True, ";PWD=" & strPassWord) _
    Else Set cnAP = DBEngine.Workspaces(0).OpenDatabase(strFile, False, True)
    
    Set rsUsers = cnAP.OpenRecordset("sms")
    
    If rsUsers.RecordCount < 1 Then MsgBox "The data-Base is empty!", vbExclamation, App.Title
    
    dbOpen = True
Exit Sub
ErrorHandler:
        If Err.Number = 3031 Then 'Password Protected Database
            MsgBox "Database is 'PassWord' protected. Please enter the Correct Password!", vbExclamation, App.Title
            dbOpen = False
        Else
            MsgBox "Unexpected Error to 'Export' the Data-Base! Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
            dbOpen = False
        End If
    Exit Sub
End Sub
