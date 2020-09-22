VERSION 5.00
Begin VB.Form frmZoom 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "PictureZoom (right click to zoom in, left click to zoom out )"
   ClientHeight    =   6180
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   7605
   Icon            =   "frmZoom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   7605
   StartUpPosition =   1  'CenterOwner
   Begin VB.VScrollBar VSImage 
      Height          =   5460
      Left            =   7320
      TabIndex        =   4
      Top             =   0
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.HScrollBar HSImage 
      Height          =   255
      Left            =   0
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
      Width           =   7290
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "Exit"
      Height          =   315
      Left            =   3210
      TabIndex        =   2
      Top             =   5835
      Width           =   855
   End
   Begin VB.PictureBox PicScroll 
      Height          =   5505
      Left            =   0
      ScaleHeight     =   5445
      ScaleWidth      =   7275
      TabIndex        =   0
      Top             =   0
      Width           =   7335
      Begin VB.PictureBox PicZoom 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   840
         ScaleHeight     =   1215
         ScaleWidth      =   1815
         TabIndex        =   1
         Top             =   1080
         Width           =   1815
      End
      Begin VB.Image ImgOrig 
         Height          =   645
         Left            =   720
         Picture         =   "frmZoom.frx":000C
         Stretch         =   -1  'True
         Top             =   2400
         Visible         =   0   'False
         Width           =   660
      End
   End
End
Attribute VB_Name = "frmZoom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private ScrollVert As Boolean, ScrollHor As Boolean
Private ZoomFact As Single
Private IsRightButt As Boolean
Const ZFactorC As Byte = 100        ' percentage increase
Const ScrollFactorC As Byte = 20    ' used to calculate scroll max and change (can play with this value)

Private Sub cmdExit_Click()
 Unload Me
End Sub

Private Sub Form_Load()
 ScrollVert = False: ScrollHor = False
 ZoomFact = ZFactorC
' CenterX = (PicScroll.Width - PicZoom.Width) / 2
' CenterY = (PicScroll.Height - PicZoom.Height) / 2
 ' center picture in container
'  PicZoom.Move CenterX, CenterY
 ZoomPicture
End Sub

Private Sub ZoomPicture()
Dim SizeX As Single, SizeY As Single
Dim Ratio As Single
Dim Wdth As Single, Hght As Single

 Screen.MousePointer = vbHourglass
 Wdth = PicScroll.ScaleWidth
 Hght = PicScroll.ScaleHeight
 Ratio = ZoomFact / 100
 ' redimension original image
 SizeX = ImgOrig.Width * Ratio
 SizeY = ImgOrig.Height * Ratio
 
 ScrollHor = IIf(SizeX > Wdth, True, False)
 ScrollVert = IIf(SizeY > Hght, True, False)
 
 PicZoom.Cls
 PicZoom.Move 0, 0, SizeX, SizeY
 PicZoom.PaintPicture ImgOrig.Picture, 0, 0, SizeX, SizeY

 ' adjust scroll bar
 If ScrollVert Then
   VSImage.Visible = True
   VSImage.Min = 0
   VSImage.Max = (PicZoom.ScaleHeight - PicScroll.ScaleHeight) / ScrollFactorC
   VSImage.SmallChange = ScrollFactorC
   VSImage.LargeChange = PicZoom.ScaleHeight / ScrollFactorC
   VSImage.Value = VSImage.Min
 Else
   VSImage.Visible = False
 End If

 If ScrollHor Then
   HSImage.Visible = True
   HSImage.Min = 0
   HSImage.Max = (PicZoom.ScaleWidth - PicScroll.ScaleWidth) / ScrollFactorC
   HSImage.SmallChange = ScrollFactorC
   HSImage.LargeChange = PicZoom.ScaleWidth / ScrollFactorC
   HSImage.Value = HSImage.Min
 Else
   HSImage.Visible = False
 End If
 Screen.MousePointer = vbDefault
End Sub

Private Sub HSImage_Change()
 If ScrollHor Then
   PicZoom.Left = -HSImage.Value * ScrollFactorC
 End If
End Sub

Private Sub piczoom_Click()
 If IsRightButt Then
    ZoomFact = ZoomFact + ZFactorC
 Else
    ZoomFact = IIf(ZoomFact <= ZFactorC, ZFactorC, ZoomFact - ZFactorC)
 End If
 ZoomPicture
End Sub

Private Sub piczoom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
 If Button = 1 Then
    IsRightButt = True
 Else
    IsRightButt = False
 End If
End Sub

Private Sub VSImage_Change()
 If ScrollVert Then
   PicZoom.Top = -VSImage.Value * ScrollFactorC
 End If
End Sub
