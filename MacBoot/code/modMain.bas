Attribute VB_Name = "modMain"
Public ProgTitle As String
Public ProgVersion As String
Public RegKey As String
Public RegHKey As Long
   
Public pBarValue As Integer
Public MaxValue As Integer
Public NextIndex As Integer
Public MaxIcons As Integer
Public IconSpacer As Integer
Public PrevIconChk As String
Public IconID As Integer
Public TotalIcons As Integer
Public MinIconsToDisplay As Integer
Public RunPeriod As Double
Public WaitPeriod As Double
Public EndRunPeriod As Double
Public pBarFiller As Integer
Public StartTimer As Long
'***************************************************
' API Declarations for making form transparent
'***************************************************
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Const GWL_EXSTYLE = -20
Const WS_EX_TRANSPARENT = &H20
'****************************************************************
'API Declarations for keeping a form on top
'****************************************************************
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'****************************************************************
'API Declarations for playing a .WAV file
'****************************************************************
'  flag values for uFlags parameter
Public Const SND_SYNC = &H0              '  play synchronously (default)
Public Const SND_ASYNC = &H1             '  play asynchronously
Public Const SND_NODEFAULT = &H2         '  silence not default, if sound not found
Public Const SND_MEMORY = &H4            '  lpszSoundName points to a memory file
Public Const SND_ALIAS = &H10000         '  name is a WIN.INI [sounds] entry
Public Const SND_FILENAME = &H20000      '  name is a file name
Public Const SND_RESOURCE = &H40004      '  name is a resource name or atom
Public Const SND_ALIAS_ID = &H110000     '  name is a WIN.INI [sounds] entry identifier
Public Const SND_ALIAS_START = 0         '  must be > 4096 to keep strings in same section of resource file
Public Const SND_LOOP = &H8              '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10           '  don't stop any currently playing sound
Public Const SND_VALID = &H1F            '  valid flags          / ;Internal /
Public Const SND_NOWAIT = &H2000         '  don't wait if the driver is busy
Public Const SND_VALIDFLAGS = &H17201F   '  Set of valid flag bits.  Anything outside this range will raise an error
Public Const SND_RESERVED = &HFF000000   '  In particular these flags are reserved
Public Const SND_TYPE_MASK = &H170007
Public Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As String, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function PlaySoundData Lib "winmm.dll" Alias "PlaySoundA" (lpData As Any, ByVal hModule As Long, ByVal dwFlags As Long) As Long
Public Declare Function waveOutGetNumDevs Lib "winmm.dll" () As Long
'  waveform audio error return values
Public Const WAVERR_BASE = 32
Public Const WAVERR_BADFORMAT = (WAVERR_BASE + 0)       '  unsupported wave format
Public Const WAVERR_STILLPLAYING = (WAVERR_BASE + 1)    '  still something playing
Public Const WAVERR_UNPREPARED = (WAVERR_BASE + 2)      '  header not prepared
Public Const WAVERR_SYNC = (WAVERR_BASE + 3)            '  device is synchronous
Public Const WAVERR_LASTERROR = (WAVERR_BASE + 3)       '  last error in range
Public m_snd() As Byte
'****************************************************************
'API Declaration for pauseing a program for a given time
'****************************************************************
Public Declare Function GetTickCount& Lib "Kernel32" ()
'****************************************************************
'API Declaration for hiding & showing the mouse
'****************************************************************
Private lShowCursor As Long
Private rtn As Integer
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
'****************************************************************
'API Declaration for capturing the entire Screen
'****************************************************************
Type lrect
     Left As Integer
     Top As Integer
     Right As Integer
     Bottom As Integer
End Type
Declare Function GetDesktopWindow Lib "user32" () As Integer
Declare Function GetDC Lib "user32" (ByVal hwnd%) As Integer
Declare Function BitBlt Lib "gdi32" (ByVal hDestDC%, ByVal X%, ByVal Y%, ByVal nWidth%, ByVal nHeight%, ByVal hSrcDC%, ByVal XSrc%, ByVal YSrc%, ByVal dwRop&) As Integer
Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Integer, ByVal hDC As Integer) As Integer
Declare Sub GetWindowRect Lib "user32" (ByVal hwnd%, lpRect As lrect)
Public TwipsPerPixel As Single
'****************************************************************
'API Declaration for tiling a picture on a form
'****************************************************************
'Public Declare Function BitBlt Lib "gdi32" ( _
'    ByVal hDestDC As Long, _
'    ByVal X As Long, _
'    ByVal Y As Long, _
'    ByVal nWidth As Long, _
'    ByVal nHeight As Long, _
'    ByVal hSrcDC As Long, _
'    ByVal xSrc As Long, _
'    ByVal ySrc As Long, _
'    ByVal dwRop As Long _
') As Long
Private TileIt As Integer
Const SRCCOPY = &HCC0020
Private TileX As Integer, TileY As Integer
Private MaximumX As Integer, MaximumY As Integer
'****************************************************************
'API Declarations for executing programs
'****************************************************************
Public Declare Function ShellExecute _
     Lib "shell32.dll" Alias "ShellExecuteA" _
     (ByVal hwnd As Long, ByVal lpOperation As String, _
      ByVal lpFile As String, ByVal lpParameters As String, _
      ByVal lpDirectory As String, _
      ByVal nShowCmd As Long) As Long
Global Const SW_SHOWNORMAL = 1

Sub TileBackground(frm As Form)
    Dim bgdImage    As Picture
    Dim X           As Integer
    Dim Y           As Integer

    Set bgdImage = frm.picDesktop.Picture
    Y = 0
    While Y < frm.Height
        X = 0
        While X < frm.Width
            frm.PaintPicture bgdImage, X, Y
            X = X + bgdImage.Width \ 2
        Wend
        Y = Y + bgdImage.Height \ 2
    Wend
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
Public Sub HideMouse()
    Do
        lShowCursor = lShowCursor - 1
        rtn = ShowCursor(False)
    Loop Until rtn < 0
End Sub


Public Sub ShowMouse()
    Do
        lShowCursor = lShowCursor - 1
        rtn = ShowCursor(True)
    Loop Until rtn >= 0
End Sub



Sub Main()
    If App.PrevInstance Then
        origtitle = App.Title
        App.Title = "...duplicate instance"
        AppActivate origtitle
'        SendKeys "% ~", True
        End
    End If

    
    pBarValue = 0
    TotalIcons = 33
    IconSpacer = 100
    MinIconsToDisplay = 4
    RunPeriod = 6
    pBarFiller = 2000

    ProgTitle = "MacBoot"
    ProgVersion = "v" & App.Major & "." & App.Minor & " build " & Right("0000" & App.Revision, 4)
    RegKey = "Software\Danrich Software\" & ProgTitle
    RegHKey = HKEY_LOCAL_MACHINE
    CreateNewKey RegHKey, RegKey

    If InStr(UCase(Command()), "CONFIG") Then frmConfig.Show Else frmMain.Show
End Sub

Public Sub Pause(lngInterval As Double)
    Dim lngNow As Long, lngEnd As Long
    lngEnd = GetTickCount() + lngInterval
    Do
        lngNow = GetTickCount()
        DoEvents
    Loop Until lngNow >= lngEnd
End Sub

Public Sub PlayMacSound()
    Dim rtn As Integer
    rtn = waveOutGetNumDevs() 'check for a sound card
    If rtn = 1 Then 'if returned is greater than 1 then a sound card exists
        Const Flags = SND_MEMORY Or SND_ASYNC Or SND_NODEFAULT
        m_snd = LoadResData(MinMax(QueryValue(RegHKey, RegKey, "SOUND Startup"), 100, 102), "WAVE")
        PlaySoundResource = PlaySoundData(m_snd(0), 0, Flags)
    Else 'otherwise no sound card found
    End If
End Sub

Public Sub StayOnTop(frmForm As Form, fOnTop As Boolean)
    Const HWND_TOPMOST = -1
    Const HWND_NOTOPMOST = -2
        
    Dim lState As Long
    Dim iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer
    With frmForm
        iLeft = .Left / Screen.TwipsPerPixelX
        iTop = .Top / Screen.TwipsPerPixelY
        iWidth = .Width / Screen.TwipsPerPixelX
        iHeight = .Height / Screen.TwipsPerPixelY
    End With
        
    If fOnTop Then
        lState = HWND_TOPMOST
    Else
        lState = HWND_NOTOPMOST
    End If
    Call SetWindowPos(frmForm.hwnd, lState, iLeft, iTop, iWidth, iHeight, 0)
End Sub



Public Sub UpdateProgressBar()
    If QueryValue(RegHKey, RegKey, "Keep On Top") = 1 Then
        StayOnTop frmMain, True
    End If
    If QueryValue(RegHKey, RegKey, "Hide Mouse") = 1 Then
        HideMouse
    End If
    
    pBarValue = pBarValue + 1
    frmMain.picpBarFill.Width = pBarValue
    frmMain.picpBarFill.Refresh
End Sub


Public Sub TileForm(TileMe As Form, TilePic As PictureBox)
    'place the line:-
    'TileForm formname,picturebox (to use as tile)
    'in the forms RESIZE routine
    MaximumX = TileMe.Width + TilePic.Width
    MaximumY = TileMe.Height + TilePic.Height
    MaximumX = MaximumX \ Screen.TwipsPerPixelX
    MaximumY = MaximumY \ Screen.TwipsPerPixelY
    Dim TileWidth As Integer, TileHeight As Integer, src_hDC As Long, dst_hDC As Long
    TileWidth = TilePic.Width \ Screen.TwipsPerPixelX
    TileHeight = TilePic.Height \ Screen.TwipsPerPixelY
    src_hDC = TilePic.hDC
    dst_hDC = TileMe.hDC
    For TileX = 0 To MaximumX Step TileWidth 'this moves right
        TileIt = BitBlt(dst_hDC, TileX, 0, TileWidth, TileHeight, src_hDC, 0, 0, SRCCOPY)
    Next TileX
    For TileY = TileHeight To MaximumY Step TileHeight 'this moves down
        TileIt = BitBlt(dst_hDC, 0, TileY, MaximumX, TileHeight, dst_hDC, 0, 0, SRCCOPY)
    Next TileY
    
'    'place the line:-
'    'TileForm formname,picturebox (to use as tile)
'    'in the forms RESIZE routine
'    MaximumX = TileMe.Width + TilePic.Width
'    MaximumY = TileMe.Height + TilePic.Height
'    MaximumX = MaximumX \ Screen.TwipsPerPixelX
'    MaximumY = MaximumY \ Screen.TwipsPerPixelY
'    Dim TileWidth As Integer, TileHeight As Integer
'    TileWidth = TilePic.Width \ Screen.TwipsPerPixelX
'    TileHeight = TilePic.Height \ Screen.TwipsPerPixelY
'    For TileY = 0 To MaximumY Step TileHeight 'this moves down
'        For TileX = 0 To MaximumX Step TileWidth 'this moves right
'            TileIt = BitBlt(TileMe.hDC, TileX, TileY, TileWidth, TileHeight, TilePic.hDC, 0, 0, SRCCOPY)
'        Next TileX
'    Next TileY
End Sub


Public Function MinMax(Value, MinValue, MaxValue)
    MinMax = Value
    If MinValue >= 0 Then
        If Value < MinValue Then MinMax = MinValue
    End If
    If MaxValue >= 0 Then
        If Value > MaxValue Then MinMax = MaxValue
    End If
End Function



