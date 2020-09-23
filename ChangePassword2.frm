VERSION 5.00
Begin VB.Form Form15 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Password"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5505
   Icon            =   "ChangePassword2.frx":0000
   LinkTopic       =   "Form15"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2505
      Width           =   1440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2745
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2505
      Width           =   1440
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change DataBase Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   105
      TabIndex        =   0
      Top             =   165
      Width           =   5265
      Begin VB.TextBox Text2 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002B2B7A&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2370
         MaxLength       =   7
         PasswordChar    =   "*"
         TabIndex        =   4
         Top             =   900
         Width           =   2540
      End
      Begin VB.TextBox Text1 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002B2B7A&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2370
         MaxLength       =   7
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   330
         Width           =   2540
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Note: Previous Backup will be Deleted and New  Backup will be Automatically made with New Pwd."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   495
         Left            =   90
         TabIndex        =   8
         Top             =   1665
         Width           =   5040
      End
      Begin VB.Label Lb1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Make sure you remember the New Password"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   120
         TabIndex        =   5
         Top             =   1395
         Width           =   5040
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   285
         Picture         =   "ChangePassword2.frx":0442
         Top             =   525
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "New Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   195
         TabIndex        =   2
         Top             =   915
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Old Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   200
         TabIndex        =   1
         Top             =   345
         Width           =   2000
      End
   End
End
Attribute VB_Name = "Form15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If Trim(Text1.Text) = "" Or Trim(Text2.Text) = "" Then Exit Sub

'kill the backup file before changing Pwd
If Dir(App.Path & "\DoctorsBackup.mdb") <> "" Then
Kill App.Path & "\DoctorsBackup.mdb"
End If

Call pChangeDataBasePswd
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
