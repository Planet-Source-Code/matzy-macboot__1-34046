VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmConfig 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "MacBoot Options"
   ClientHeight    =   6135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7725
   Icon            =   "frmConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   7725
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picMacOS98 
      AutoSize        =   -1  'True
      Height          =   660
      Left            =   5880
      ScaleHeight     =   600
      ScaleWidth      =   1650
      TabIndex        =   31
      Top             =   4800
      Width           =   1710
   End
   Begin VB.Frame Frame3 
      Caption         =   "Information"
      Height          =   2175
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   7455
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Click to send an EMail -->"
         Height          =   195
         Left            =   3720
         TabIndex        =   30
         Top             =   240
         Width           =   1800
      End
      Begin VB.Label lblInfo 
         Caption         =   "Program information... blah blah blah"
         Height          =   1515
         Left            =   120
         TabIndex        =   29
         Top             =   600
         UseMnemonic     =   0   'False
         Width           =   7125
      End
      Begin VB.Label lblVersion 
         Caption         =   "MacBoot vX.XX build XXXX"
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblEMail 
         AutoSize        =   -1  'True
         Caption         =   "macboot@matzy.co.uk"
         Height          =   195
         Left            =   5640
         TabIndex        =   27
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   6840
      TabIndex        =   24
      Top             =   5640
      Width           =   735
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "&Save"
      Height          =   375
      Left            =   5880
      TabIndex        =   23
      Top             =   5640
      Width           =   735
   End
   Begin VB.Frame fmSound 
      Caption         =   "Startup Sounds"
      Height          =   2295
      Left            =   5880
      TabIndex        =   18
      Top             =   2400
      Width           =   1695
      Begin VB.CommandButton cmdTestSound 
         Caption         =   "Test Sound"
         Height          =   855
         Left            =   240
         Picture         =   "frmConfig.frx":08CA
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   1200
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Sound 3"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   21
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Sound 2"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optSound 
         Caption         =   "Sound 1"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Background"
      Height          =   2535
      Left            =   120
      TabIndex        =   4
      Top             =   3480
      Width           =   5655
      Begin VB.PictureBox picTile 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         Height          =   1980
         Left            =   3480
         ScaleHeight     =   1920
         ScaleWidth      =   1920
         TabIndex        =   17
         Top             =   360
         Width           =   1980
      End
      Begin VB.OptionButton optPlain 
         Caption         =   "Plain Coloured"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   1455
      End
      Begin VB.OptionButton optTrans 
         Caption         =   "Transparent"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   360
         Width           =   1455
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Starry Night"
         Height          =   255
         Index           =   9
         Left            =   1680
         TabIndex        =   14
         Top             =   1320
         Width           =   1335
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Dark Blue Speckles"
         Height          =   255
         Index           =   8
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   1935
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Blue Jeans"
         Height          =   255
         Index           =   7
         Left            =   1680
         TabIndex        =   12
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Light Grey Fibre"
         Height          =   255
         Index           =   6
         Left            =   1680
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Blue Speckles"
         Height          =   255
         Index           =   5
         Left            =   1680
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Coarse Brown Fibre"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   9
         Top             =   1800
         Width           =   1815
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Light Blue Speckles"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   8
         Top             =   1560
         Width           =   1935
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Dark Grey Fibre"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1575
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "Blue Leather"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   1080
         Width           =   1335
      End
      Begin VB.OptionButton optBgnd 
         Caption         =   "MacOS"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.Label lblSample 
         Alignment       =   2  'Center
         Caption         =   "Click on sample pad to change colour"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2160
         Visible         =   0   'False
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "MacBoot "
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   2400
      Width           =   5655
      Begin VB.CheckBox chkAutoRun 
         Caption         =   "Run MacBoot At StartUp"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   2295
      End
      Begin VB.CheckBox chkHideMouse 
         Caption         =   "Hide Mouse Cursor"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.CheckBox chkOnTop 
         Caption         =   "Keep MacBoot On Top"
         Height          =   255
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SelectedSound As Integer
Private MacOShigh As Boolean

Sub SetupOptions()
    '
    ' Run at startup
    '
    chkAutoRun.Value = 0
    If QueryValue(RegHKey, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "111 - MacBoot") <> "" Then
        chkAutoRun.Value = 1
    End If
    '
    ' Hide mouse
    '
    chkHideMouse.Value = 0
    If QueryValue(RegHKey, RegKey, "Hide Mouse") = 1 Then
        chkHideMouse.Value = 1
    End If
    '
    ' Keep on top
    '
    chkOnTop.Value = 0
    If QueryValue(RegHKey, RegKey, "Keep On Top") = 1 Then
        chkOnTop.Value = 1
    End If
    '
    ' Set options for background to use
    '
    optTrans.Value = False
    optPlain.Value = False
    If QueryValue(RegHKey, RegKey, "BGND Transparent") = 1 Then
        optTrans.Value = True
    ElseIf QueryValue(RegHKey, RegKey, "BGND Plain") = 1 Then
        optPlain.Value = True
    Else
        optBgnd.Item(MinMax(QueryValue(RegHKey, RegKey, "BGND Tile") - 100, 0, 109)).Value = True
    End If
    '
    ' Set options for sound to use
    '
    optSound.Item(MinMax(QueryValue(RegHKey, RegKey, "SOUND Startup") - 100, 0, 102)).Value = True
End Sub

Sub UpdateOptions()
    '
    ' Run at startup
    '
    If chkAutoRun.Value = 1 Then
        SetKeyValue RegHKey, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "111 - MacBoot", App.Path & "\" & App.EXEName & ".exe", REG_SZ
    Else
        DeleteValue RegHKey, "SOFTWARE\Microsoft\Windows\CurrentVersion\RunServices", "111 - MacBoot"
    End If
    '
    ' Hide mouse
    '
    If chkHideMouse.Value = 1 Then
        SetKeyValue RegHKey, RegKey, "Hide Mouse", 1, REG_DWORD
    Else
        SetKeyValue RegHKey, RegKey, "Hide Mouse", 0, REG_DWORD
    End If
    '
    ' Keep on top
    '
    If chkOnTop.Value = 1 Then
        SetKeyValue RegHKey, RegKey, "Keep On Top", 1, REG_DWORD
    Else
        SetKeyValue RegHKey, RegKey, "Keep On Top", 0, REG_DWORD
    End If
    '
    ' Set options for background to use
    '
    If optTrans.Value = True Then
        SetKeyValue RegHKey, RegKey, "BGND Transparent", 1, REG_DWORD
        SetKeyValue RegHKey, RegKey, "BGND Plain", 0, REG_DWORD
    ElseIf optPlain.Value = True Then
        SetKeyValue RegHKey, RegKey, "BGND Transparent", 0, REG_DWORD
        SetKeyValue RegHKey, RegKey, "BGND Plain", 1, REG_DWORD
        SetKeyValue RegHKey, RegKey, "BGND Colour", picTile.BackColor, REG_DWORD
    Else
        SetKeyValue RegHKey, RegKey, "BGND Transparent", 0, REG_DWORD
        SetKeyValue RegHKey, RegKey, "BGND Plain", 0, REG_DWORD
        For i = 0 To optBgnd.Count - 1
            If optBgnd.Item(i).Value = True Then _
            SetKeyValue RegHKey, RegKey, "BGND Tile", 100 + i, REG_DWORD
        Next i
    End If
    '
    ' Set options for sound to use
    '
    For i = 0 To optSound.Count - 1
        If optSound.Item(i).Value = True Then _
        SetKeyValue RegHKey, RegKey, "SOUND Startup", 100 + i, REG_DWORD
    Next i
    End
End Sub
Private Sub cmdCancel_Click()
    End
End Sub

Private Sub cmdOk_Click()
    UpdateOptions
End Sub

Private Sub cmdTestSound_Click()
    Dim rtn As Integer
    rtn = waveOutGetNumDevs() 'check for a sound card
    If rtn = 1 Then 'if returned is greater than 1 then a sound card exists
        
        m_snd = LoadResData(100 + SelectedSound, "WAVE")
        PlaySoundResource = PlaySoundData(m_snd(0), 0, SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT)
    Else 'otherwise no sound card found
    End If
End Sub

Private Sub fmSound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.ForeColor = vbBlack
    If MacOShigh = True Then
        picMacOS98.Picture = LoadResPicture("MACOSNORM", vbResBitmap)
        picMacOS98.AutoSize = True
        MacOShigh = False
    End If
End Sub


Private Sub Form_Load()
    ProgVersion = ProgTitle & " v" & App.Major & "." & App.Minor & " build " & Right("0000" & App.Revision, 4)
    lblVersion.Caption = ProgVersion
    proginfo = ProgTitle & " was written purely for fun & performs "
    proginfo = proginfo & "no real function apart from immitating the boot up sequence of an "
    proginfo = proginfo & "Apply Mac. It may not be 100% accurate, but this could be due mainly "
    proginfo = proginfo & "to the fact that I have not got a Mac, or never even had one!"
    proginfo = proginfo & vbCrLf & vbCrLf
    proginfo = proginfo & ProgTitle & " is FREEware. Technical support is limited, but not unavailable! "
    proginfo = proginfo & "I am more than happy to recieve emails letting me know what you think, "
    proginfo = proginfo & "and even ideas you may have for improvements. "
    proginfo = proginfo & "Clicking on my EMail address above should launch you default mail editor"
    lblInfo.Caption = proginfo
    
    picTile.AutoSize = True
    picTile.Picture = LoadResPicture(100, vbResBitmap)
    picMacOS98.Picture = LoadResPicture("MACOSNORM", vbResBitmap)
    picMacOS98.AutoSize = True
        
    Dim rtn As Integer
    rtn = waveOutGetNumDevs() 'check for a sound card
    If rtn = 1 Then 'if returned is greater than 1 then a sound card exists
        fmSound.Enabled = True
        optSound.Item(0).Enabled = True
        optSound.Item(1).Enabled = True
        optSound.Item(2).Enabled = True
        cmdTestSound.Enabled = True
    Else 'otherwise no sound card found
        fmSound.Enabled = False
        optSound.Item(0).Enabled = False
        optSound.Item(1).Enabled = False
        optSound.Item(2).Enabled = False
        cmdTestSound.Enabled = False
    End If
    
    SetupOptions
End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.ForeColor = vbBlack
    If MacOShigh = True Then
        picMacOS98.Picture = LoadResPicture("MACOSNORM", vbResBitmap)
        picMacOS98.AutoSize = True
        MacOShigh = False
    End If
End Sub


Private Sub Frame1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.ForeColor = vbBlack
End Sub


Private Sub Frame2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If MacOShigh = True Then
        picMacOS98.Picture = LoadResPicture("MACOSNORM", vbResBitmap)
        picMacOS98.AutoSize = True
        MacOShigh = False
    End If

End Sub


Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.ForeColor = vbBlack
End Sub

Private Sub lblEMail_Click()
    ShellExecute Me.hwnd, "open", "mailto:macboot@matzy.co.uk?Subject=MacBoot Support", "", "C:\", SW_SHOWNORMAL
End Sub

Private Sub lblEMail_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblEMail.ForeColor = vbBlue
End Sub


Private Sub optBgnd_Click(Index As Integer)
    picTile.Visible = False
    picTile.Picture = LoadResPicture(100 + Index, vbResBitmap)
    picTile.AutoSize = True
    picTile.Height = MinMax(picTile.Height, 0, 2415)
    picTile.Width = MinMax(picTile.Width, 0, 2055)
    picTile.Visible = True
    lblSample.Visible = False
End Sub


Private Sub optPlain_Click()
    picTile.Visible = False
    picTile.Picture = LoadPicture()
    picTile.AutoSize = True
    picTile.Height = 1980
    picTile.Width = 1980
    picTile.BackColor = QueryValue(RegHKey, RegKey, "BGND Colour")
    picTile.Visible = True
    lblSample.Visible = True
End Sub

Private Sub optSound_Click(Index As Integer)
    SelectedSound = Index
End Sub


Private Sub optTrans_Click()
    picTile.Visible = False
    picTile.Picture = LoadResPicture("TRANSPARENT", vbResBitmap)
    picTile.AutoSize = True
    picTile.Height = MinMax(picTile.Height, 0, 2415)
    picTile.Width = MinMax(picTile.Width, 0, 2055)
    picTile.Visible = True
    lblSample.Visible = False
End Sub


Private Sub picMacOS98_Click()
    Dim rtn As Integer
    rtn = waveOutGetNumDevs() 'check for a sound card
    If rtn = 1 Then 'if returned is greater than 1 then a sound card exists
        m_snd = LoadResData("EEP", "WAVE")
        PlaySoundResource = PlaySoundData(m_snd(0), 0, SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT)
    Else 'otherwise no sound card found
    End If
'    ShellExecute Me.hwnd, "open", "http://members.xoom.com/macos98/", "", "C:\", SW_SHOWNORMAL
MsgBox "Since writing this program long ago this site has now closed down. ", , "Ooops..!"
    
    
End Sub

Private Sub picMacOS98_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picMacOS98.Picture = LoadResPicture("MACOSHIGH", vbResBitmap)
    picMacOS98.AutoSize = True
    MacOShigh = True
End Sub


Private Sub picTile_Click()
    If optPlain.Value = True Then
        CommonDialog1.CancelError = True
        On Error GoTo ErrHandler
        CommonDialog1.Flags = cdlCCRGBInit
        CommonDialog1.Color = picTile.BackColor
        CommonDialog1.ShowColor
        picTile.BackColor = CommonDialog1.Color
        Exit Sub
    End If

ErrHandler: 'User pressed cancel
End Sub


