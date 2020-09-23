VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Image Outliner"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   7935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picClient 
      Height          =   495
      Left            =   1080
      ScaleHeight     =   435
      ScaleWidth      =   1095
      TabIndex        =   9
      Top             =   480
      Width           =   1155
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load"
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   735
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   495
      Left            =   6120
      TabIndex        =   6
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.HScrollBar sclSensibility 
      Height          =   495
      Left            =   3480
      Max             =   256
      TabIndex        =   4
      Top             =   120
      Value           =   25
      Width           =   1695
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   1800
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.CommandButton cmdConvert 
      Caption         =   "Convert"
      Height          =   495
      Left            =   2640
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6120
      Top             =   -120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   120
      ScaleHeight     =   598
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   0
      Top             =   720
      Width           =   7695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   9015
      Left            =   120
      ScaleHeight     =   599
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   511
      TabIndex        =   8
      Top             =   720
      Visible         =   0   'False
      Width           =   7695
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "25"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5280
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Private Declare Function SetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long

'Type declarations
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Type BSBITMAP
    Info As BITMAP
    Bits() As Byte
End Type

'Private variable declarations
Private SBuffer As BSBITMAP
Private DBuffer As BSBITMAP

Dim sensibility As Integer
Private WithEvents m_cScroll As cScrollBars
Attribute m_cScroll.VB_VarHelpID = -1


'This is where it all takes place..
'The sub loops through all the color components of all the pixels in a picture
'determining if the difference if in colors is great enough to set a black pixel
'I have tried to optimize the code inside the loop, but I'm sure there's some work
'left over for you nit pickers out there..
Private Sub Outline24Bit(source As IPictureDisp, dest As IPictureDisp)
Dim white As Boolean
Dim x As Long, y As Long, xplus As Long, yplus As Long, xminus As Long, yminus As Long
    
    pbProgress.Value = 0

    'Get information about picture boxes and declare arrays to hold bits
    GetObject source, Len(SBuffer.Info), SBuffer.Info
    ReDim SBuffer.Bits(0 To SBuffer.Info.bmWidthBytes / SBuffer.Info.bmWidth - 1, _
                       0 To SBuffer.Info.bmWidth - 1, _
                       0 To SBuffer.Info.bmHeight - 1) As Byte
    GetObject dest, Len(DBuffer.Info), DBuffer.Info
    ReDim DBuffer.Bits(0 To DBuffer.Info.bmWidthBytes / DBuffer.Info.bmWidth - 1, _
                       0 To DBuffer.Info.bmWidth - 1, _
                       0 To DBuffer.Info.bmHeight - 1) As Byte

    'If color depth is not 24 or 32 bit color exit program
    If SBuffer.Info.bmBitsPixel < 24 Then
        MsgBox "Desktop color must be 24 or 32 bit" & vbCrLf & _
               "for program to function properly." & vbCrLf & _
               "Please exit and change desktop color" & vbCrLf & _
               "depth before running program."
    End If

    GetBitmapBits source, SBuffer.Info.bmWidthBytes * SBuffer.Info.bmHeight, SBuffer.Bits(0, 0, 0)
    
    For y = 1 To SBuffer.Info.bmHeight - 2
        For x = 1 To SBuffer.Info.bmWidth - 2
            xminus = x - 1
            yminus = y - 1
            xplus = x + 1
            yplus = y + 1
            white = False
            'wierdest thing here.. the 'not+greater than'-approach is actually a little bit faster than just a 'lesser than'-operator..
            If Not Abs(CInt(SBuffer.Bits(0, xminus, y)) - CInt(SBuffer.Bits(0, xplus, y))) > sensibility Then
            If Not Abs(CInt(SBuffer.Bits(0, x, yminus)) - CInt(SBuffer.Bits(0, x, yplus))) > sensibility Then
            If Not Abs(CInt(SBuffer.Bits(1, xminus, y)) - CInt(SBuffer.Bits(1, xplus, y))) > sensibility Then
            If Not Abs(CInt(SBuffer.Bits(1, x, yminus)) - CInt(SBuffer.Bits(1, x, yplus))) > sensibility Then
            If Not Abs(CInt(SBuffer.Bits(2, xminus, y)) - CInt(SBuffer.Bits(2, xplus, y))) > sensibility Then
            If Not Abs(CInt(SBuffer.Bits(2, x, yminus)) - CInt(SBuffer.Bits(2, x, yplus))) > sensibility Then
                DBuffer.Bits(2, x, y) = 255
                DBuffer.Bits(1, x, y) = 255
                DBuffer.Bits(0, x, y) = 255
                white = True
            End If
            End If
            End If
            End If
            End If
            End If
            If Not white Then
                DBuffer.Bits(2, x, y) = 0
                DBuffer.Bits(1, x, y) = 0
                DBuffer.Bits(0, x, y) = 0
            End If
        Next x
        pbProgress.Value = (y / SBuffer.Info.bmHeight) * 100
    Next y
    
    SetBitmapBits dest, DBuffer.Info.bmWidthBytes * DBuffer.Info.bmHeight, DBuffer.Bits(0, 0, 0)
    
    pbProgress.Value = 100
    
End Sub

Private Sub cmdConvert_Click()
Picture2.Cls
       
Outline24Bit Picture2.Picture, Picture2.Image

End Sub

'old native VB way
'Private Function pixeldiff(ByVal pixel1 As Long, ByVal pixel2 As Long) As Boolean
'Dim bytes1() As Integer, bytes2() As Integer

'splitpixel pixel1, bytes1()
'splitpixel pixel2, bytes2()

'For t = 0 To 2
'    If Abs(bytes1(t) - bytes2(t)) > sensibility Then
'        pixeldiff = True
'        Exit Function
'    End If
'Next t

'pixeldiff = False

'End Function


'Private Sub splitpixel(pixel As Long, bytes() As Integer)
'ReDim bytes(2)

'bytes(0) = pixel Mod 256
'bytes(1) = pixel \ 256 Mod 256
'bytes(2) = pixel \ 65536 Mod 256

'End Sub

Private Sub cmdOpen_Click()
Dim mywidth As Single

cd.FileName = ""
cd.ShowOpen
If cd.FileName <> "" Then
    Picture1.Visible = True
    Picture1.Picture = LoadPicture(cd.FileName)
    Picture2.Picture = LoadPicture(cd.FileName)
    Picture2.Cls
    Picture2.Move (Picture1.Left + Picture1.Width) ', 720, Picture1.Width, Picture1.Height
    mywidth = Picture2.Left + Picture2.Width
    If mywidth + 375 < 8055 Then mywidth = 8055
    Form1.Width = mywidth + 375
    Form1.Height = 1080 + Picture2.Height
    picClient.Move 0, 0, mywidth, Picture2.Top + Picture2.Height
    Form_Resize
    DoEvents
    
    Outline24Bit Picture2.Picture, Picture2.Image
    Me.Caption = "Image Outliner - " & cd.FileTitle
End If

End Sub

Private Sub cmdSave_Click()

cd.FileName = ""

cd.ShowSave

If cd.FileName <> "" Then SavePicture Picture2.Image, cd.FileName

End Sub

Private Sub cmdLoad_Click()
Dim mywidth As Single

cd.FileName = ""
cd.ShowOpen
If cd.FileName <> "" Then
    Picture1.Visible = False
    Picture2.Left = 120
    Picture2.Picture = LoadPicture(cd.FileName)
    Picture2.Cls
    mywidth = Picture2.Left + Picture2.Width
    If mywidth + 375 < 8055 Then mywidth = 8055
    Form1.Width = mywidth + 375
    Form1.Height = 1080 + Picture2.Height
    picClient.Move 0, 0, mywidth, Picture2.Top + Picture2.Height
    Form_Resize
    DoEvents

    Outline24Bit Picture2.Picture, Picture2.Image
    Me.Caption = "Image Outliner - " & cd.FileTitle
End If

End Sub

Private Sub Form_Load()
sensibility = 25
Dim ctl As Control
Dim mywidth As Single

   ' Set up scroll bars:
   Set m_cScroll = New cScrollBars
   m_cScroll.Create Me.hwnd
   'm_cScroll.SmallChange(efsVertical) = lblDemo(0).Height \ Screen.TwipsPerPixelY + 2
   
   ' To make it easier to design the form,
   ' we place all the controls on the form,
   ' then switch them into the client box
   ' at run-time.
   On Error Resume Next
   For Each ctl In Controls
      If Not ctl Is picClient Then
         If ctl.Container Is Me Then
            Set ctl.Container = picClient
         End If
      End If
   Next ctl
   picClient.BorderStyle = 0
   picClient.Move 0, 0, Me.Width, Me.Height
   
    If Command$ <> "" Then
        cd.FileName = Replace(Command$, """", "")
        Picture1.Visible = False
        Picture2.Left = 120
        Picture2.Picture = LoadPicture(cd.FileName)
        Picture2.Cls
        mywidth = Picture2.Left + Picture2.Width
        If mywidth + 375 < 8055 Then mywidth = 8055
        Form1.Width = mywidth + 375
        Form1.Height = 1440 + Picture2.Height
        picClient.Move 0, 0, mywidth, Picture2.Top + Picture2.Height
        Form_Resize
        DoEvents

        Outline24Bit Picture2.Picture, Picture2.Image
        Me.Caption = "Image Outliner - " & cd.FileTitle
    End If

End Sub

Private Sub Form_Resize()
Dim lHeight As Long
Dim lWidth As Long
Dim lProportion As Long
   
   ' Pixels are the minimum change size for a screen object.
   ' Therefore we set the scroll bars in pixels.
   Do
    DoEvents
   Loop While Me.ScaleHeight = 0
   lHeight = (picClient.Height - Me.ScaleHeight) \ Screen.TwipsPerPixelY
   If (lHeight > 0) Then
      lProportion = lHeight \ ((Me.ScaleHeight + 1) \ Screen.TwipsPerPixelY) + 1
      m_cScroll.LargeChange(efsVertical) = lHeight \ lProportion
      m_cScroll.Max(efsVertical) = lHeight
      m_cScroll.Visible(efsVertical) = True
   Else
      m_cScroll.Visible(efsVertical) = False
   End If
   
   lWidth = (picClient.Width - Me.ScaleWidth) \ Screen.TwipsPerPixelX
   If (lWidth > 0) Then
      lProportion = lWidth \ ((Me.ScaleWidth + 1) \ Screen.TwipsPerPixelX) + 1
      m_cScroll.LargeChange(efsHorizontal) = lWidth \ lProportion
      m_cScroll.Max(efsHorizontal) = lWidth
      m_cScroll.Visible(efsHorizontal) = True
   Else
      m_cScroll.Visible(efsHorizontal) = False
   End If
End Sub

Private Sub sclSensibility_Change()
sensibility = sclSensibility.Value
Label1.Caption = sensibility

End Sub

Private Sub sclSensibility_Scroll()
sensibility = sclSensibility.Value
Label1.Caption = sensibility

End Sub

Private Sub m_cScroll_Change(eBar As EFSScrollBarConstants)
   If (m_cScroll.Visible(eBar)) Then
      If (eBar = efsHorizontal) Then
         picClient.Left = -m_cScroll.Value(eBar) * Screen.TwipsPerPixelX
      Else
         picClient.Top = -m_cScroll.Value(eBar) * Screen.TwipsPerPixelY
      End If
   Else
      picClient.Move 0, 0
   End If
End Sub

Private Sub m_cScroll_Scroll(eBar As EFSScrollBarConstants)
   m_cScroll_Change eBar
End Sub
