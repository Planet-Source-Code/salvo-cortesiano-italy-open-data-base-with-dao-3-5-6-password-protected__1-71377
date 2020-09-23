VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Open DB"
   ClientHeight    =   3150
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5295
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3150
   ScaleWidth      =   5295
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Close Connection"
      Enabled         =   0   'False
      Height          =   465
      Left            =   1800
      TabIndex        =   4
      Top             =   2295
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open DB"
      Height          =   465
      Left            =   1800
      TabIndex        =   0
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label3 
      Caption         =   "n/a"
      Height          =   405
      Left            =   375
      TabIndex        =   3
      Top             =   1680
      Width           =   4725
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "salvocortesiano@hotmail.com"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1365
      TabIndex        =   2
      Top             =   2835
      Width           =   3285
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Open a Data-Base with DAO 3.6 password protected! Only for beginners! Data-Base PassWord is (password)..."
      Height          =   390
      Left            =   690
      TabIndex        =   1
      Top             =   210
      Width           =   4005
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFileName As String
Dim strPassWord As String
Dim ConnStatus As Boolean
Private Sub Command1_Click()
   
    On Error GoTo ErrorHandler
    
    strFileName = DialogFile(Form1.hWnd, 1, "Open Data-Base", "", "File DataBase" & Chr(0) & "*.mdb" & Chr(0) & "Tutti i files" & Chr(0) & "*.*", App.Path, "mdb")
    If Len(strFileName) = 0 Then
            Label3.Caption = "Action cancel by User!"
        Exit Sub
    End If
    
    Call OpenDatabase(strFileName)
    
getInfoDB:
    If dbOpen = False Then
        Label3.Caption = "Data-Base is PassWord protected..."
        ConnStatus = False
        Command2.Enabled = False
        GoTo sProtectdataBase
    Else
        Label3.Caption = "Data-Base opened successfully... (" & rsUsers.RecordCount * (rsUsers.PercentPosition * 0.01) + 1 & ") of (" & rsUsers.RecordCount & ") Record's!"
        ConnStatus = True
        Command1.Enabled = False
        Command2.Enabled = True
    End If
    
Exit Sub

sProtectdataBase:
    strPassWord = InputBox("Insert the appropiate 'PassWord' for this data-Base!", "PassWord?", strPassWord)
    If Len(strPassWord) = 0 Then
            Label3.Caption = "Action cancel by User!"
        Exit Sub
    End If
    
    Call OpenDatabase(strFileName, strPassWord)
    
    If dbOpen = False Then
        GoTo sProtectdataBase
    Else
        Call OpenDatabase(strFileName, strPassWord)
        GoTo getInfoDB
    End If
Exit Sub
ErrorHandler:
    If Err.Number = 3031 Then ' data-Base Protect!
        MsgBox "The data-Base that You want to open is protected from 'PassWord'! Ok to Continue", vbExclamation, App.Title
        GoTo sProtectdataBase
    Else
        MsgBox "Unexpected Error! Error #" & Err.Number & ". " & Err.Description, vbCritical, App.Title
    End If
Exit Sub
End Sub

Private Sub Command2_Click()
    If ConnStatus = True Then
        Call CloseConnection
        ConnStatus = False
        Command2.Enabled = False
        Command1.Enabled = True
        Label3.Caption = "n/a"
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If ConnStatus = True Then Call CloseConnection
    End
End Sub



Private Sub CloseConnection()
    If ConnStatus = True Then
        rsUsers.Close
        cnAP.Close
        Set rsUsers = Nothing
        Set cnAP = Nothing
    End If
End Sub
