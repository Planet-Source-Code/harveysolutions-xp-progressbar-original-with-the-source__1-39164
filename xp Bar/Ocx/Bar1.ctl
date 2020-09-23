VERSION 5.00
Begin VB.UserControl CarlosProgressBar 
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   ScaleHeight     =   43
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   Begin VB.PictureBox Bmid 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   960
      Picture         =   "Bar1.ctx":0000
      ScaleHeight     =   255
      ScaleWidth      =   90
      TabIndex        =   4
      Top             =   285
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.PictureBox Bright 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   975
      Picture         =   "Bar1.ctx":1F87
      ScaleHeight     =   255
      ScaleWidth      =   60
      TabIndex        =   3
      Top             =   15
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox Bleft 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   810
      Picture         =   "Bar1.ctx":3F9D
      ScaleHeight     =   255
      ScaleWidth      =   60
      TabIndex        =   2
      Top             =   270
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.PictureBox dot 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   165
      Left            =   750
      Picture         =   "Bar1.ctx":5FB6
      ScaleHeight     =   165
      ScaleWidth      =   120
      TabIndex        =   1
      Top             =   60
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.PictureBox BarPic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   75
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   39
      TabIndex        =   0
      Top             =   45
      Width           =   585
   End
End
Attribute VB_Name = "CarlosProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'*****************************************************
'*  Author : Carl Harvey
'*  Date  : 09/21/2002
'*****************************************************

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private mblnValue As Long
Private mblnMax As Long

Private Sub UserControl_Resize()
 Height = 255
 'Change scale mode to fit the new size
 ScaleMode = 1
 BarPic.Width = Width
 BarPic.Height = Height
 BarPic.Left = 0
 BarPic.Top = 0
 'put back scale mode to pixel
 ScaleMode = 3
 DrawBar
End Sub


Private Sub DrawBar()
  Dim Nb As Integer
  Dim DX
  Dim StopAtPixel As Integer
  BarPic.Cls
  Nb = (BarPic.Width - 8) / 6
  'draw middle background
  For i = 0 To Nb
    BitBlt BarPic.hDC, 4 + i * 6, 0, Bmid.Width, Bmid.Height, Bmid.hDC, 0, 0, vbSrcCopy
  Next
  
  'If the max or value is > 0 then the value must be show
  If mblnValue > 0 And mblnMax > 0 Then
    StopAtPixel = GetDotToDraw()
    DX = 6
    'draw the value dots
    Do While DX < StopAtPixel
        BitBlt BarPic.hDC, DX, 3, 8, dot.Height, dot.hDC, 0, 0, vbSrcCopy
        DX = DX + 10
    Loop
  End If
  
  'draw corners, right and left
  BitBlt BarPic.hDC, 0, 0, 4, Bmid.Height, Bleft.hDC, 0, 0, vbSrcCopy
  BitBlt BarPic.hDC, BarPic.Width - 4, 0, 4, Bmid.Height, Bright.hDC, 0, 0, vbSrcCopy
  
End Sub

Private Function GetDotToDraw() As Integer
  Dim Pc As Integer
  'This calculate the percentage and apply it to pixel width
  Pc = BarPic.Width * (mblnValue / mblnMax)
  GetDotToDraw = Pc
End Function

Public Property Get Value() As Long
   Value = mblnValue
End Property

Public Property Let Value(ByVal NewValue As Long)
   If NewValue <= mblnMax Then
     mblnValue = NewValue
     PropertyChanged "Value"
     DrawBar
   End If
End Property

Public Property Get Max() As Long
   Value = mblnMax
End Property

Public Property Let Max(ByVal NewValue As Long)
   mblnMax = NewValue
   PropertyChanged "Value"
End Property
