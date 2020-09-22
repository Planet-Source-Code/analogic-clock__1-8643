VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Analogic Clock"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3990
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   266
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   800
      Left            =   1320
      Top             =   4200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3750
      Left            =   120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   248
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   248
      TabIndex        =   0
      Top             =   120
      Width           =   3750
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   3960
      Width           =   3735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function Ellipse Lib "gdi32" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Type POINTAPI
        x As Long
        y As Long
End Type
Private pnt As POINTAPI
Private Sub Form_Load()
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Picture1.Cls
'obtaining the current time, in format HH:MM:SS
timp = Format(Time, "hh:mm:ss")
Ellipse Picture1.hdc, Picture1.Width / 2 - 5, Picture1.Height / 2 - 5, Picture1.Width / 2 + 5, Picture1.Height / 2 + 5
'setting the origin
MoveToEx Picture1.hdc, Picture1.Width / 2, Picture1.Height / 2, pnt
'obtaining hour, minutes and secundes values from "timp" string
ora = Val(Left(timp, 2))
If ora > 12 Then ora = ora - 12
minut = Val(Mid(timp, 4, 2))
secunde = Val(Right(timp, 2))
'"drawing" the seconds
xx = Picture1.Width / 2 + 120 * Cos(secunde * 6 * 3.14 / 180 - 3.14 / 2)
yy = Picture1.Height / 2 + 120 * Sin(secunde * 6 * 3.14 / 180 - 3.14 / 2)
LineTo Picture1.hdc, xx, yy
'"drawing" the minutes
MoveToEx Picture1.hdc, Picture1.Width / 2, Picture1.Height / 2, pnt
xx = Picture1.Width / 2 + 100 * Cos(minut * 6 * 3.14 / 180 - 3.14 / 2)
yy = Picture1.Height / 2 + 100 * Sin(minut * 6 * 3.14 / 180 - 3.14 / 2)
LineTo Picture1.hdc, xx, yy
MoveToEx Picture1.hdc, Picture1.Width / 2 + 1, Picture1.Height / 2 + 1, pnt
LineTo Picture1.hdc, xx + 1, yy + 1
'"drawing" the hours
MoveToEx Picture1.hdc, Picture1.Width / 2, Picture1.Height / 2, pnt
xx = Picture1.Width / 2 + 60 * Cos(ora * 30 * 3.14 / 180 - 3.14 / 2)
yy = Picture1.Height / 2 + 60 * Sin(ora * 30 * 3.14 / 180 - 3.14 / 2)
LineTo Picture1.hdc, xx, yy
MoveToEx Picture1.hdc, Picture1.Width / 2 - 1, Picture1.Height / 2 - 1, pnt
LineTo Picture1.hdc, xx - 1, yy - 1
MoveToEx Picture1.hdc, Picture1.Width / 2 + 1, Picture1.Height / 2 + 1, pnt
LineTo Picture1.hdc, xx + 1, yy + 1
Label1.Caption = timp
Beep
End Sub
