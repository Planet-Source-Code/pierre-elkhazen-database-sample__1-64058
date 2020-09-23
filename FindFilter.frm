VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00000080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Find Patients"
   ClientHeight    =   2340
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5340
   Icon            =   "FindFilter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form7"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "FindFilter.frx":0442
   ScaleHeight     =   2340
   ScaleWidth      =   5340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Find: Enter Part of -or- Whole Name"
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
      Height          =   2235
      Left            =   60
      TabIndex        =   1
      Top             =   45
      Width           =   5200
      Begin VB.CommandButton PFindNextRecord 
         Caption         =   "Find"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   1635
         Width           =   1065
      End
      Begin VB.CommandButton PFilterRecords 
         Caption         =   "Filter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   1590
         TabIndex        =   8
         ToolTipText     =   "Filter will show only the Records that match your Search Critirea"
         Top             =   1635
         Width           =   1200
      End
      Begin VB.ComboBox Combo2 
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
         Height          =   330
         ItemData        =   "FindFilter.frx":074C
         Left            =   915
         List            =   "FindFilter.frx":074E
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "LastName"
         Top             =   855
         Width           =   1785
      End
      Begin VB.TextBox Text2 
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
         Height          =   330
         Left            =   3090
         TabIndex        =   4
         Top             =   855
         Width           =   1965
      End
      Begin VB.TextBox Text1 
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
         Height          =   285
         Left            =   3090
         TabIndex        =   0
         Top             =   285
         Width           =   1965
      End
      Begin VB.ComboBox Combo1 
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
         Height          =   315
         ItemData        =   "FindFilter.frx":0750
         Left            =   915
         List            =   "FindFilter.frx":0752
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "FirstName"
         Top             =   285
         Width           =   1785
      End
      Begin VB.Label Label2 
         Caption         =   "AND"
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
         Height          =   240
         Left            =   345
         TabIndex        =   7
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   2775
         TabIndex        =   6
         Top             =   840
         Width           =   270
      End
      Begin VB.Label Label1 
         Caption         =   "="
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   13.5
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   2760
         TabIndex        =   3
         Top             =   255
         Width           =   315
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5250
      Top             =   5235
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*********************   DataBase Related  Code Section  ***************
'*******************************************************************************

'There are Many ways to Find Records.
'You Can Use Find Statement as shown here
'You Can Filter on that Criteria as shown in Filter Command
'You Can Build SQL Statement for that Criteria as shown throughout this Project
'You can use the Seek Statement
'......


Dim sFindStr


Private Sub PFindNextRecord_Click()
On Error GoTo 100

'Set the Find Statement.
'Adodc does not have the Ability to use "and" with Find so a loop is needed
sFindStr = Combo1.Text & " Like " & "'" & Text1.Text & "*" & "'"

'Find Loop.
Do While True

If Form1.Adodc1.Recordset.AbsolutePosition > 1 Then Form1.Adodc1.Recordset.MoveNext

'Apply Find
Form1.Adodc1.Recordset.Find sFindStr
    
    If Text2.Text = "" Then
    If Left(Form1.txtFields(2), Len(Text1.Text)) = Text1.Text Then Exit Do
    Else
    If Left(Form1.txtFields(1), Len(Text2.Text)) = Text2.Text Then Exit Do
    End If


If Form1.Adodc1.Recordset.AbsolutePosition < 1 Then Exit Do
Loop
                                        Exit Sub
100
                                        MsgBox Err.Description
End Sub


Private Sub PFilterRecords_Click()
                                On Error GoTo 300
                                If PFilterRecords.Caption = "Filter" Then
                                Form1.Adodc1.Refresh
                                'SQL Filter Query to Filter All Records that Match Criteria
                                'Here Like is used with * to Filter on Records Starting with that Criteria
                                
'Set the SQL Filter Statement
If Text2.Text = "" Then
sFilterStr = Combo1.Text & " like '" & Text1.Text & "*'"
Else
sFilterStr = Combo1.Text & " like '" & Text1.Text & "*'" & " and " & Combo2.Text & " like '" & Text2.Text & "*'"
End If

'Apply Filter
Form1.Adodc1.Recordset.Filter = sFilterStr
                                PFilterRecords.Caption = "NoFilter"
                                Exit Sub
                                End If

                                If PFilterRecords.Caption = "NoFilter" Then
'Remove Filter
Form1.Adodc1.Recordset.Filter = adFilterNone
Form1.Adodc1.Refresh
                                PFilterRecords.Caption = "Filter"
                                End If

Exit Sub
300
MsgBox Err.Description
End Sub




'*******************************************************************************
'*************************  Non Related DataBase Code Section  *********************
'*******************************************************************************


Private Sub Form_Unload(Cancel As Integer)
If PFilterRecords.Caption = "NoFilter" Then
Form1.Adodc1.Recordset.Filter = adFilterNone
Form1.Adodc1.Refresh
End If
If Form1.Adodc1.Recordset.AbsolutePosition = -3 Then Form1.Adodc1.Recordset.MoveFirst
End Sub

Private Sub Text1_KeyUp(KeyCode As Integer, Shift As Integer)
Form1.Adodc1.Recordset.MoveFirst
If Len(Text1.Text) = 1 Then
Text1 = UCase(Text1)
Text1.SelStart = Len(Text1.Text)
End If
End Sub

Private Sub Text2_KeyUp(KeyCode As Integer, Shift As Integer)
Form1.Adodc1.Recordset.MoveFirst
If Len(Text2.Text) = 1 Then
Text2 = UCase(Text2)
Text2.SelStart = Len(Text2.Text)
End If
End Sub
Private Sub Combo1_Click()
Form1.Adodc1.Recordset.MoveFirst
End Sub


Private Sub Combo2_Click()
Form1.Adodc1.Recordset.MoveFirst
End Sub

