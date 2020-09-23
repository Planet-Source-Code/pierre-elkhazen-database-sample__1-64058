VERSION 5.00
Begin VB.Form Form38 
   BackColor       =   &H00000040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About DataBase Sample"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4845
   Icon            =   "About.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   4845
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture7 
      Height          =   825
      Index           =   9
      Left            =   8430
      Picture         =   "About.frx":014A
      ScaleHeight     =   765
      ScaleWidth      =   1575
      TabIndex        =   6
      Top             =   3945
      Visible         =   0   'False
      Width           =   1635
   End
   Begin VB.PictureBox Picture7 
      Height          =   735
      Index           =   8
      Left            =   8445
      Picture         =   "About.frx":804C
      ScaleHeight     =   675
      ScaleWidth      =   1590
      TabIndex        =   5
      Top             =   2880
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.PictureBox Picture7 
      Height          =   750
      Index           =   7
      Left            =   8385
      Picture         =   "About.frx":FF4E
      ScaleHeight     =   690
      ScaleWidth      =   1515
      TabIndex        =   4
      Top             =   1755
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   5925
      Top             =   3150
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5880
      Left            =   105
      ScaleHeight     =   5820
      ScaleWidth      =   4560
      TabIndex        =   0
      Top             =   75
      Width           =   4620
      Begin VB.CommandButton Command1 
         BackColor       =   &H00000080&
         Height          =   195
         Left            =   4200
         MouseIcon       =   "About.frx":17E50
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   150
         Width           =   360
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00000080&
         Height          =   5685
         Left            =   4320
         ScaleHeight     =   5625
         ScaleWidth      =   105
         TabIndex        =   2
         Top             =   70
         Width           =   160
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   4110
         LargeChange     =   1000
         Left            =   6240
         SmallChange     =   1000
         TabIndex        =   1
         Top             =   15
         Width           =   225
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H00400000&
         Height          =   6480
         Left            =   150
         ScaleHeight     =   6480
         ScaleWidth      =   4245
         TabIndex        =   10
         Top             =   -15
         Width           =   4245
         Begin VB.PictureBox PicLogo 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1560
            Left            =   1185
            Picture         =   "About.frx":1815A
            ScaleHeight     =   1560
            ScaleWidth      =   1545
            TabIndex        =   16
            Top             =   330
            Width           =   1545
         End
         Begin VB.PictureBox Picture11 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1560
            Left            =   420
            ScaleHeight     =   1560
            ScaleWidth      =   3255
            TabIndex        =   11
            Top             =   2175
            Width           =   3255
            Begin VB.TextBox TextAbout 
               Alignment       =   2  'Center
               BackColor       =   &H00000080&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   300
               Index           =   2
               Left            =   180
               TabIndex        =   12
               Text            =   "DataBase 2004"
               Top             =   90
               Width           =   2745
            End
            Begin VB.Label LabelAbout 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "http://geocities.com/medjugorjesite"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Index           =   1
               Left            =   30
               MouseIcon       =   "About.frx":2005E
               MousePointer    =   99  'Custom
               TabIndex        =   19
               Top             =   855
               Width           =   3120
            End
            Begin VB.Label LabelAboutV 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   " V4.5"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Left            =   1185
               TabIndex        =   15
               Top             =   405
               Width           =   570
            End
            Begin VB.Label LabelAboutCC 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   " © 2005 Computer Club 2000+"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Left            =   270
               TabIndex        =   14
               Top             =   615
               Width           =   2400
            End
            Begin VB.Label LabelAbout 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "pierrelk@hotmail.com"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00000040&
               Height          =   210
               Index           =   2
               Left            =   390
               MouseIcon       =   "About.frx":20368
               MousePointer    =   99  'Custom
               TabIndex        =   13
               Top             =   1125
               Width           =   2235
            End
         End
         Begin VB.TextBox Text2 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            ForeColor       =   &H00800000&
            Height          =   180
            Left            =   0
            MaxLength       =   40
            TabIndex        =   17
            Top             =   495
            Width           =   3460
         End
         Begin VB.TextBox Text1 
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   1260
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   18
            Text            =   "About.frx":20672
            Top             =   1230
            Visible         =   0   'False
            Width           =   3840
         End
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Club 2000+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   705
         TabIndex        =   8
         Top             =   4365
         Width           =   3030
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Computer Club 2000+"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   735
         TabIndex        =   9
         Top             =   4395
         Width           =   2745
      End
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer Club 2000+"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   615
      Left            =   5040
      TabIndex        =   7
      Top             =   4725
      Width           =   3030
   End
End
Attribute VB_Name = "Form38"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PAnimationConstant
Dim xpos2 As Long
Dim ypos2 As Long
Dim Ptext1, Ptext2, Ptext3, Ptext4, Ptext5, Ptext6, Ptext7
Dim PImageIndex
Dim PgenConstant As Boolean
Dim Ptext(7)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'Slider Control
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
xpos2 = X
ypos2 = Y
End Sub
Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Top <= 0 Then
Command1.Top = 10
End If

If Command1.Top >= Picture1.Height - 40 Then
Command1.Top = Picture1.Height - 60
Exit Sub
End If

If Button = 1 Then
Command1.Move Command1.Left, Y + Command1.Top + ypos2
Picture2.Move Picture2.Left, -(Command1.Top) * (Picture2.Height / Picture1.Height)
End If

End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Command1.Top <= 0 Then
Command1.Top = 10
Picture2.Top = 10
End If
If Command1.Top >= Picture1.Height - 300 Then
Command1.Top = Picture1.Height - 300
End If
End Sub

Private Sub Form_Load()
On Error Resume Next
PAnimationConstant = Timer

If PgenConstant = False Then
Ptext(1) = "This Computer Program is Licensed to:"
Ptext(2) = "Pierre"
Ptext(3) = "Liscence No:"
Ptext(4) = "S191/2004"
Ptext(5) = "Address:"
Ptext(6) = "Beirut, Lebanon"
Ptext(7) = Text1.Text
End If

Picture2.ForeColor = &HFFC0C0
Picture2.CurrentY = Picture11.Top + Picture11.Height + 200 + 10
PWriteTextLoad
Picture2.Height = Picture2.CurrentY + 200
Picture2.ForeColor = &H400000
Picture2.CurrentY = Picture11.Top + Picture11.Height + 200
PWriteTextLoad


cy = Picture2.CurrentY '= Picture2.Height
Picture2.ForeColor = &H40&
Picture2.FontBold = True
Picture2.Print Text1.Text
Picture2.Height = Picture2.CurrentY + 200
Picture2.CurrentY = cy
Picture2.Print Text1.Text

End Sub

Private Sub LabelAbout_Click(Index As Integer)
If Index = 2 Then
ShellExecute Me.hwnd, "open", "mailto:pierrelk@hotmail.com", ByVal 0&, "", 3
End If
If Index = 1 Then
ShellExecute Me.hwnd, "open", "http://geocities.com/medjugorjesite", ByVal 0&, "", 3
End If

End Sub

Private Sub LabelAbout_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 2
LabelAbout(i).ForeColor = &H40&
Next
LabelAbout(Index).ForeColor = &HC0&
End Sub

Private Sub Picture11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
For i = 1 To 2
LabelAbout(i).ForeColor = &H40&
Next

End Sub

Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Vertical Slider Bar Control
If Y > Command1.Top Then
Command1.Top = Command1.Top + 500
Else
Command1.Top = Command1.Top - 500
End If
Picture2.Move Picture2.Left, -(Command1.Top) * (Picture2.Height / Picture1.Height)
End Sub

Private Sub PWriteTextLoad()
On Error Resume Next
'Print Texts and Load Image1() and Place it Over Graphic text to be able to Click on it to edit it.
Picture2.Line (100, Picture2.CurrentY)-(Picture4.Left - 300, Picture2.CurrentY)
Picture2.Print vbCrLf
Picture2.FontBold = True
Picture2.Print Ptext(1)
Picture2.FontBold = False
Picture2.Print Ptext(2)
Picture2.FontBold = True
Picture2.Print Ptext(3)
Picture2.FontBold = False
Picture2.Print Ptext(4)
Picture2.FontBold = True
Picture2.Print Ptext(5)
Picture2.FontBold = False
Picture2.Print Ptext(6)
Picture2.Print
Picture2.Line (100, Picture2.CurrentY)-(Picture4.Left - 300, Picture2.CurrentY)
Picture2.Print vbCrLf
End Sub


Private Sub Timer1_Timer()
'Animating Logo with only 3 different Pictures
If Timer <= PAnimationConstant + 0.05 Then PicLogo.Picture = Picture7(7).Picture: GoTo 100
If Timer <= PAnimationConstant + 0.1 Then PicLogo.Picture = Picture7(8).Picture: GoTo 100
If Timer <= PAnimationConstant + 0.15 Then PicLogo.Picture = Picture7(9).Picture
If Timer >= PAnimationConstant + 0.15 Then
PAnimationConstant = Timer
End If
100
End Sub

