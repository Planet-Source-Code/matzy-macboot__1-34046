VERSION 5.00
Begin VB.Form frmMain 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   6615
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7500
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmMain.frx":08CA
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   7500
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4860
      Left            =   120
      Picture         =   "frmMain.frx":0BD4
      ScaleHeight     =   4860
      ScaleWidth      =   6330
      TabIndex        =   0
      Top             =   120
      Width           =   6330
      Begin VB.PictureBox picpBarFill 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   720
         Picture         =   "frmMain.frx":27DC
         ScaleHeight     =   240
         ScaleWidth      =   4785
         TabIndex        =   1
         Top             =   4320
         Width           =   4785
      End
      Begin VB.PictureBox picPbarBgnd 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Left            =   720
         Picture         =   "frmMain.frx":2956
         ScaleHeight     =   240
         ScaleWidth      =   4800
         TabIndex        =   2
         Top             =   3960
         Width           =   4800
      End
   End
   Begin VB.PictureBox picDesktop 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   840
      ScaleHeight     =   480
      ScaleWidth      =   1200
      TabIndex        =   3
      Top             =   5160
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Image imgIcon 
      Appearance      =   0  'Flat
      Height          =   480
      Index           =   0
      Left            =   240
      Picture         =   "frmMain.frx":2D61
      Top             =   5160
      Width           =   480
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Sub DesktopOptions()
    If QueryValue(RegHKey, RegKey, "BGND Transparent") = 1 Then
        ' Transparent desktop - bit of a cheat. I grab a picture of the
        ' current desktop and use it as the background!
        '
        Me.Visible = False
        picDesktop.Visible = False
        picDesktop.AutoRedraw = True
        picDesktop.Width = Screen.Width
        picDesktop.Height = Screen.Height
        Me.Picture = CaptureScreen
    ElseIf QueryValue(RegHKey, RegKey, "BGND Plain") = 1 Then
        ' Blank coloured desktop
        Me.BackColor = QueryValue(RegHKey, RegKey, "BGND Colour")
    Else
        ' Tiled desktop using choosen graphic
        picDesktop.Picture = LoadResPicture(MinMax(QueryValue(RegHKey, RegKey, "BGND Tile"), 100, 109), vbResBitmap)
        TileBackground Me
'        TileForm Me, picDesktop
    End If
End Sub


Private Sub Form_Load()
    
'    HideMouse
    If QueryValue(RegHKey, RegKey, "Keep On Top") = 1 Then
        StayOnTop Me, True
    End If
    
    MaxValue = picPbarBgnd.Width - pBarFiller
    WaitPeriod = ((RunPeriod * 1000) / picPbarBgnd.Width)
    NextIndex = 0
    Do
        Randomize
        MaxIcons = MinMax(Int(Rnd * TotalIcons) + MinIconsToDisplay, 1, 33)
        If Screen.Width > (imgIcon(0).Width + IconSpacer) * MaxIcons Then
            Exit Do
        End If
    Loop
    
    picPbarBgnd.Left = (picLogo.Width / 2) - (picPbarBgnd.Width / 2)
    picPbarBgnd.Top = 4250
    
    picpBarFill.Top = picPbarBgnd.Top
    picpBarFill.Left = picPbarBgnd.Left
    picpBarFill.Width = 0
    
    imgIcon(0).Top = Screen.Height - 950
    imgIcon(0).Left = 0
    imgIcon(0).Width = 0
    imgIcon(0).Visible = False
    
    ' Display MacBoot over Mac style desktop
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    picLogo.Move (Screen.Width / 2) - (picLogo.Width / 2), (Screen.Height / 2) - (picLogo.Height / 2)
    
    DesktopOptions
    
'    ' Display MacBoot over current windows desktop
'    Call SetWindowLong(Me.hWnd, GWL_EXSTYLE, WS_EX_TRANSPARENT)
'    Call SetWindowPos(Me.hWnd, HWND_NOTOPMOST, 0&, 0&, 0&, 0&, SWP_SHOWME)
      
    EndRunPeriod = GetTickCount + (RunPeriod * 1000)
    If QueryValue(RegHKey, RegKey, "Hide Mouse") = 0 Then
        ShowMouse
    End If
    Me.Show
    Me.Refresh
    
    PlayMacSound
    
    For X = 1 To pBarFiller / 3
        UpdateProgressBar
    Next X
    For i = 1 To MaxValue Step (MaxValue / MaxIcons)
        For X = 1 To MaxValue / MaxIcons
            UpdateProgressBar
        Next X
   
        NextIndex = imgIcon.Count
        Load imgIcon(NextIndex)
        imgIcon(NextIndex).Top = imgIcon(NextIndex - 1).Top
        imgIcon(NextIndex).Left = imgIcon(NextIndex - 1).Left + imgIcon(NextIndex - 1).Width + IconSpacer
        Do
            Randomize
            IconID = Int(Rnd * TotalIcons) + 100
            If InStr(PrevIconChk, "<" & Trim(Str(IconID)) & ">") = 0 Then
                PrevIconChk = PrevIconChk & "<" & Trim(Str(IconID)) & ">"
                Exit Do
            End If
        Loop
        imgIcon(NextIndex).Picture = LoadResPicture(IconID, vbResIcon)
        imgIcon(NextIndex).Visible = True
        imgIcon(NextIndex).Refresh
        Randomize
        If EndRunPeriod >= GetTickCount Then Pause Int(Rnd * (100 * (RunPeriod))) + 100
    Next i
    For X = pBarValue To picPbarBgnd.Width
        UpdateProgressBar
    Next X
 
    Do
        If GetTickCount > EndRunPeriod Then Exit Do
        Pause 500
    Loop

    ShowMouse
    Me.Hide
    End
End Sub

Public Sub GetTwipsPerPixel()
    Me.ScaleMode = 3
    NumPix = Me.ScaleHeight
    TwipsPerPixel = Me.ScaleHeight / NumPix
    Me.ScaleMode = 1
End Sub



