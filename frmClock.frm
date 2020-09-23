VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LED Dot Matrix Clock"
   ClientHeight    =   1710
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   114
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   548
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   240
      Top             =   60
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   780
      ScaleHeight     =   15
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   77
      TabIndex        =   0
      Top             =   60
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   330
      Index           =   0
      Left            =   720
      Shape           =   3  'Circle
      Top             =   360
      Visible         =   0   'False
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Switch As Integer

Private Sub Form_Load()
    Shape1(0).FillColor = RGB(96, 0, 0)
    Shape1(0).BorderColor = RGB(128, 0, 0)
    For x = 0 To Picture1.Width
        For y = 0 To Picture1.Height
            N = N + 1
            Load Shape1(N)
            Shape1(N).Move (x * 8) - x, (y * 8) - y, 6, 6
            Shape1(N).Visible = True
        Next y
    Next x
    DotMatrix "My Clock :-)"
End Sub

Private Sub DotMatrix(Txt As String)
    Picture1.Print Txt
    For x = 0 To Picture1.Width
        For y = 0 To Picture1.Height
            N = N + 1
            If Picture1.Point(x, y) = 0 Then
                Shape1(N).FillColor = RGB(255, 0, 0)
            Else
                Shape1(N).FillColor = RGB(96, 0, 0)
            End If
        Next y
    Next x
    Picture1.Cls
End Sub

Private Sub Timer1_Timer()
    If Switch < 4 Then
        DotMatrix Time
    Else
        DotMatrix Format(Date, "mmm d, yyyy")
    End If
    Switch = Switch + 1
    If Switch > 6 Then Switch = 0
End Sub
