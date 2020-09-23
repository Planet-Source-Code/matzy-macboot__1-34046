VERSION 5.00
Begin VB.Form frmRunConfig 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "frmRunConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmRunConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Hide
    If FileExists("macboot.exe") Then
        Shell "macboot.exe CONFIG", vbNormalFocus
    Else
        MsgBox "Please make sure the file MACBOOT.EXE is in the same directory as the MacBoot configuration utility", vbOKOnly & vbCritical, "An error has occured..."
    End If
    End
End Sub


Function FileExists(Filename As String) As Boolean
    Dim TempAttr As Integer
    On Error GoTo ErrorFileExist 'any errors show that the file doesnt exist, so goto this label
    TempAttr = GetAttr(Filename) 'get the attributes of the files
    FileExists = ((TempAttr And vbDirectory) = 0) 'check if its a directory and not a file
    GoTo ExitFileExist
   
ErrorFileExist:
    FileExists = False 'return that the file doesnt exist
    Resume ExitFileExist 'carry on with the code
   
ExitFileExist:
    On Error GoTo 0 'clear all errors
End Function

