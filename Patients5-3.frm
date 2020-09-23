VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form1 
   Caption         =   "Data Base Sample"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   6765
   Icon            =   "Patients5-3.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   6765
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   7650
      Left            =   60
      TabIndex        =   1
      Top             =   0
      Width           =   6660
      _ExtentX        =   11748
      _ExtentY        =   13494
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   617
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   " Password"
      TabPicture(0)   =   "Patients5-3.frx":014A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Lbl1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "CmdOK"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " Patients"
      TabPicture(1)   =   "Patients5-3.frx":059C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "picToolBar"
      Tab(1).Control(1)=   "FramePatient"
      Tab(1).Control(2)=   "Adodc1"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   " Grid View"
      TabPicture(2)   =   "Patients5-3.frx":126E
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "DataGrid1"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   " Reports"
      TabPicture(3)   =   "Patients5-3.frx":1632
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "LblPatients(10)"
      Tab(3).Control(1)=   "DataGrid2"
      Tab(3).Control(2)=   "Combo1"
      Tab(3).ControlCount=   3
      TabCaption(4)   =   " DB Utilities"
      TabPicture(4)   =   "Patients5-3.frx":1ADD
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Label1"
      Tab(4).Control(1)=   "LabelUtilities(0)"
      Tab(4).Control(2)=   "Shape1"
      Tab(4).ControlCount=   3
      Begin MSDataGridLib.DataGrid DataGrid1 
         Bindings        =   "Patients5-3.frx":1C37
         Height          =   6915
         Left            =   -74790
         TabIndex        =   38
         Top             =   555
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   12197
         _Version        =   393216
         AllowUpdate     =   0   'False
         Appearance      =   0
         BackColor       =   -2147483628
         BorderStyle     =   0
         Enabled         =   -1  'True
         ForeColor       =   4194304
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.ComboBox Combo1 
         ForeColor       =   &H00400000&
         Height          =   315
         ItemData        =   "Patients5-3.frx":1C4C
         Left            =   -74820
         List            =   "Patients5-3.frx":1C62
         TabIndex        =   43
         Text            =   "Select Report"
         Top             =   615
         Width           =   3765
      End
      Begin VB.CommandButton CmdOK 
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
         Height          =   390
         Left            =   2835
         TabIndex        =   37
         Top             =   4500
         Width           =   1500
      End
      Begin VB.Frame Frame1 
         Caption         =   "Please Enter Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   1260
         TabIndex        =   35
         Top             =   2745
         Width           =   4425
         Begin VB.TextBox txtPaswd 
            BackColor       =   &H80000014&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            IMEMode         =   3  'DISABLE
            Left            =   1125
            MaxLength       =   40
            PasswordChar    =   "*"
            TabIndex        =   0
            Top             =   840
            Width           =   2505
         End
         Begin VB.Label Lbl1 
            BackStyle       =   0  'Transparent
            Caption         =   "Password :"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   10
            Left            =   1185
            TabIndex        =   40
            Top             =   540
            Width           =   930
         End
         Begin VB.Label LblPWD 
            BackStyle       =   0  'Transparent
            Caption         =   "555"
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
            Left            =   2190
            TabIndex        =   39
            Top             =   540
            Width           =   1410
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   420
            Picture         =   "Patients5-3.frx":1D47
            Top             =   840
            Width           =   480
         End
      End
      Begin VB.PictureBox picToolBar 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   -72225
         ScaleHeight     =   480
         ScaleWidth      =   3690
         TabIndex        =   29
         Top             =   390
         Width           =   3690
         Begin VB.CommandButton Command1 
            BackColor       =   &H00000080&
            Height          =   375
            Left            =   2670
            Picture         =   "Patients5-3.frx":2189
            Style           =   1  'Graphical
            TabIndex        =   48
            TabStop         =   0   'False
            ToolTipText     =   "Refresh Records"
            Top             =   60
            Width           =   420
         End
         Begin VB.CommandButton CmdFind 
            BackColor       =   &H00000080&
            Height          =   375
            Left            =   1725
            Picture         =   "Patients5-3.frx":2807
            Style           =   1  'Graphical
            TabIndex        =   34
            TabStop         =   0   'False
            ToolTipText     =   "Find Patient"
            Top             =   60
            Width           =   420
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H00000080&
            Height          =   375
            Left            =   900
            Picture         =   "Patients5-3.frx":2CF9
            Style           =   1  'Graphical
            TabIndex        =   33
            TabStop         =   0   'False
            ToolTipText     =   "Delete Record"
            Top             =   60
            Width           =   420
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H00000080&
            Height          =   375
            Left            =   480
            Picture         =   "Patients5-3.frx":323B
            Style           =   1  'Graphical
            TabIndex        =   32
            TabStop         =   0   'False
            ToolTipText     =   "Add Record"
            Top             =   60
            Width           =   420
         End
         Begin VB.CommandButton CmdSort 
            BackColor       =   &H00000080&
            Height          =   375
            Left            =   1305
            Picture         =   "Patients5-3.frx":372D
            Style           =   1  'Graphical
            TabIndex        =   31
            TabStop         =   0   'False
            ToolTipText     =   "Sort Records"
            Top             =   60
            Width           =   420
         End
         Begin VB.CommandButton CmdSave 
            BackColor       =   &H00000080&
            Height          =   375
            Left            =   2145
            Picture         =   "Patients5-3.frx":3C1F
            Style           =   1  'Graphical
            TabIndex        =   30
            TabStop         =   0   'False
            ToolTipText     =   "Save Rerord"
            Top             =   60
            Width           =   420
         End
         Begin VB.Image ImageHelp 
            Height          =   330
            Index           =   1
            Left            =   3285
            MouseIcon       =   "Patients5-3.frx":4111
            MousePointer    =   99  'Custom
            Picture         =   "Patients5-3.frx":441B
            ToolTipText     =   "Help"
            Top             =   90
            Width           =   210
         End
      End
      Begin VB.Frame FramePatient 
         Caption         =   "Patients"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   6345
         Left            =   -74820
         TabIndex        =   2
         Top             =   795
         Width           =   6255
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            CausesValidation=   0   'False
            DataField       =   "ID"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   5
            Left            =   4500
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   46
            Top             =   3360
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            DataField       =   "FirstVisitDate"
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "M/d/yyyy"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   3
            EndProperty
            DataSource      =   "Adodc1"
            Height          =   315
            Left            =   2175
            TabIndex        =   44
            Top             =   3735
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   556
            _Version        =   393216
            CheckBox        =   -1  'True
            Format          =   19595265
            CurrentDate     =   38439
         End
         Begin VB.ComboBox ComboGender 
            BackColor       =   &H80000014&
            DataField       =   "Gender"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   315
            ItemData        =   "Patients5-3.frx":483A
            Left            =   2175
            List            =   "Patients5-3.frx":4844
            TabIndex        =   41
            Top             =   3360
            Width           =   1905
         End
         Begin VB.TextBox txtFields 
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
            ForeColor       =   &H00000040&
            Height          =   285
            Index           =   0
            Left            =   2175
            MaxLength       =   40
            TabIndex        =   16
            Text            =   "Pierre"
            Top             =   315
            Width           =   2325
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "State"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   4
            Left            =   2175
            MaxLength       =   20
            TabIndex        =   15
            Top             =   3015
            Width           =   2310
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Diagnosis"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   840
            Index           =   6
            Left            =   2175
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   14
            Top             =   4470
            Width           =   3550
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Notes"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   795
            Index           =   7
            Left            =   2175
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   13
            Top             =   5415
            Width           =   3550
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            CausesValidation=   0   'False
            DataField       =   "Zip"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   9
            Left            =   4515
            MaxLength       =   20
            TabIndex        =   12
            Top             =   3015
            Width           =   1215
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Address"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   645
            Index           =   10
            Left            =   2175
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   11
            Top             =   1935
            Width           =   3550
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "LastName"
            DataSource      =   "Adodc1"
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
            Index           =   1
            Left            =   2175
            MaxLength       =   30
            MouseIcon       =   "Patients5-3.frx":4856
            TabIndex        =   10
            Top             =   1170
            Width           =   3550
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "FirstName"
            DataSource      =   "Adodc1"
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
            Index           =   2
            Left            =   2175
            MaxLength       =   30
            MouseIcon       =   "Patients5-3.frx":4B60
            TabIndex        =   9
            Top             =   765
            Width           =   3550
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "City"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   3
            Left            =   2175
            MaxLength       =   40
            TabIndex        =   8
            Top             =   2655
            Width           =   3550
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Balance"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   14
            Left            =   2175
            MaxLength       =   50
            TabIndex        =   7
            Top             =   4095
            Width           =   2310
         End
         Begin VB.TextBox txtFields 
            Appearance      =   0  'Flat
            BackColor       =   &H80000014&
            DataField       =   "Telephone"
            DataSource      =   "Adodc1"
            ForeColor       =   &H00400000&
            Height          =   285
            Index           =   12
            Left            =   2175
            MaxLength       =   40
            TabIndex        =   6
            Top             =   1560
            Width           =   3550
         End
         Begin VB.CommandButton ComboGenConditionAdd 
            BackColor       =   &H00000080&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2175
            Picture         =   "Patients5-3.frx":4E6A
            Style           =   1  'Graphical
            TabIndex        =   5
            ToolTipText     =   "Add: Enter Condition then Click Add"
            Top             =   4485
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.CommandButton ComboGenConditiondelete 
            BackColor       =   &H00000080&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2460
            Picture         =   "Patients5-3.frx":4FEC
            Style           =   1  'Graphical
            TabIndex        =   4
            ToolTipText     =   "Delete: Click Delete, Select Condition, then Click Delete again"
            Top             =   4485
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.ComboBox ComboGenCondition 
            BackColor       =   &H80000014&
            ForeColor       =   &H00400000&
            Height          =   315
            ItemData        =   "Patients5-3.frx":516E
            Left            =   2175
            List            =   "Patients5-3.frx":5170
            TabIndex        =   3
            Top             =   4755
            Visible         =   0   'False
            Width           =   3570
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "ID"
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
            Index           =   0
            Left            =   4200
            TabIndex        =   45
            Top             =   3375
            Width           =   255
         End
         Begin VB.Label LblDoctor 
            BackStyle       =   0  'Transparent
            Caption         =   "Doctor"
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
            Height          =   255
            Index           =   0
            Left            =   540
            TabIndex        =   28
            Top             =   315
            Width           =   690
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "First Name"
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
            Index           =   1
            Left            =   405
            TabIndex        =   27
            Top             =   780
            Width           =   1005
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Last Name"
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
            Index           =   2
            Left            =   405
            TabIndex        =   26
            Top             =   1185
            Width           =   1095
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
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
            Index           =   3
            Left            =   405
            MouseIcon       =   "Patients5-3.frx":5172
            MousePointer    =   99  'Custom
            TabIndex        =   25
            Top             =   1965
            Width           =   1095
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Phone No."
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
            Index           =   4
            Left            =   405
            MouseIcon       =   "Patients5-3.frx":55B4
            MousePointer    =   99  'Custom
            TabIndex        =   24
            Top             =   1605
            Width           =   1095
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "First Visit"
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
            Index           =   5
            Left            =   405
            MouseIcon       =   "Patients5-3.frx":59F6
            MousePointer    =   99  'Custom
            TabIndex        =   23
            Top             =   3750
            Width           =   1350
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Diagnosis"
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
            Index           =   6
            Left            =   405
            MouseIcon       =   "Patients5-3.frx":5E38
            MousePointer    =   99  'Custom
            TabIndex        =   22
            Top             =   4485
            Width           =   1650
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Notes"
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
            Index           =   7
            Left            =   405
            TabIndex        =   21
            Top             =   5415
            Width           =   1095
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "City"
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
            Index           =   9
            Left            =   405
            TabIndex        =   20
            Top             =   2670
            Width           =   390
         End
         Begin VB.Line Line1 
            BorderColor     =   &H00000080&
            X1              =   2175
            X2              =   5745
            Y1              =   1080
            Y2              =   1080
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000080&
            X1              =   5730
            X2              =   5730
            Y1              =   765
            Y2              =   1075
         End
         Begin VB.Line Line3 
            BorderColor     =   &H00000080&
            X1              =   5730
            X2              =   5730
            Y1              =   1170
            Y2              =   1480
         End
         Begin VB.Line Line4 
            BorderColor     =   &H00000080&
            X1              =   2175
            X2              =   5745
            Y1              =   1485
            Y2              =   1485
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "State / District"
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
            Index           =   8
            Left            =   405
            TabIndex        =   19
            Top             =   3030
            Width           =   1605
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Gender"
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
            Index           =   12
            Left            =   405
            TabIndex        =   18
            Top             =   3375
            Width           =   1485
         End
         Begin VB.Label LblPatients 
            BackStyle       =   0  'Transparent
            Caption         =   "Balance"
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
            Index           =   14
            Left            =   405
            TabIndex        =   17
            Top             =   4110
            Width           =   1620
         End
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   330
         Left            =   -74820
         ToolTipText     =   "Records Navigation"
         Top             =   7200
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   582
         ConnectMode     =   3
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   8
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   128
         ForeColor       =   15329769
         Orientation     =   0
         Enabled         =   -1
         Connect         =   ""
         OLEDBString     =   ""
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   ""
         Caption         =   "Records"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   6390
         Left            =   -74850
         TabIndex        =   42
         Top             =   1080
         Width           =   6300
         _ExtentX        =   11113
         _ExtentY        =   11271
         _Version        =   393216
         AllowUpdate     =   0   'False
         BackColor       =   -2147483628
         Enabled         =   -1  'True
         ForeColor       =   4194304
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   165
         Left            =   -74280
         Shape           =   3  'Circle
         Top             =   1110
         Width           =   165
      End
      Begin VB.Label LabelUtilities 
         Caption         =   "Create a DataBase and Table  (2 Lines Code)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   225
         Index           =   0
         Left            =   -73995
         MouseIcon       =   "Patients5-3.frx":627A
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   1095
         Width           =   5000
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Watch for more Useful  DataBase Code and  Utilities in Future Updates"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   -74910
         TabIndex        =   49
         Top             =   675
         Width           =   6315
      End
      Begin VB.Label LblPatients 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Index           =   10
         Left            =   -70875
         MouseIcon       =   "Patients5-3.frx":6584
         TabIndex        =   47
         Top             =   645
         Width           =   2250
      End
      Begin VB.Label Lbl1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "DataBase Sample"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Index           =   0
         Left            =   1230
         TabIndex        =   36
         Top             =   1995
         Width           =   4395
      End
   End
   Begin VB.Menu mUtilities 
      Caption         =   "Utilities"
      Begin VB.Menu mCompactDataBase 
         Caption         =   "Compact DataBase"
      End
      Begin VB.Menu mBackupDataBase 
         Caption         =   "Backup DataBase"
      End
      Begin VB.Menu mRestoreDataBase 
         Caption         =   "Restore DataBase"
      End
      Begin VB.Menu mChangeDBPassword 
         Caption         =   "Change DB Password"
      End
      Begin VB.Menu strep3 
         Caption         =   "-"
      End
      Begin VB.Menu mVisitMyMedjugorjeWeb 
         Caption         =   "Visit My Medjugorje Web"
      End
   End
   Begin VB.Menu mAbout 
      Caption         =   "About"
      Begin VB.Menu mHelp 
         Caption         =   "Help"
      End
      Begin VB.Menu strep5 
         Caption         =   "-"
      End
      Begin VB.Menu mAbout1 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************
'*********************   DataBase Related  Code Section  ***************
'*******************************************************************************

Dim SortStr

Public Sub CmdOK_Click() 'Connecting to the DB
                                                'Log On and Connect
                                                On Error GoTo 100
                                                If txtPaswd.Text = "" Then Beep: Exit Sub
'Connecting to the DB      for Access 2000 or higher        DB name                                       DB Original Password=555
Adodc1.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Doctors.mdb" & ";Jet OLEDB:Database Password=" & txtPaswd.Text
'SQL Statement to Select * (all Fields) from  the Table Named Patients (all Patients)
Adodc1.RecordSource = "Select * from Patients"
Adodc1.Refresh
                                                PCurrtPWd = txtPaswd.Text
                                                PSetSStab
                                                mUtilities.Enabled = True
                                                Exit Sub
100
                                                MsgBox Err.Description
                                                Adodc1.ConnectionString = ""
End Sub

Private Sub cmdAdd_Click()
                                                On Error Resume Next
Adodc1.Recordset.AddNew
                                                txtFields(2).SetFocus
End Sub

Private Sub cmdDelete_Click()
                                                On Error Resume Next
                                                If MsgBox("Are you sure you want to Delete " & Adodc1.Recordset.Fields("FirstName") & " " & Adodc1.Recordset.Fields("LastName") & "?", vbCritical + vbYesNo) = vbNo Then Exit Sub
Adodc1.Recordset.Delete
End Sub
Private Sub CmdSave_Click()
                                                On Error Resume Next
Adodc1.Recordset.Update

End Sub

Private Sub CmdSort_Click()
                                                On Error GoTo 100
                                                'SortStr is the Field to Sort on - see txtFields_Click
                                                If SortStr = "" Then MsgBox "Click a Field to Sort": Exit Sub
Adodc1.Recordset.Sort = SortStr                 'or Adodc1.Recordset.Sort = "FirstName,LastName"  etc... cannot sort on Memo fields
                                                MsgBox "Records Sorted by  " & SortStr


                                                Exit Sub
100
                                                MsgBox Err.Description
End Sub
Private Sub Command1_Click()
Adodc1.Refresh
End Sub


'*******************************************************************************
'*********************   Reports Section  ***************
'*******************************************************************************

'           Use SQL and Filter Combination to get Great Results


Private Sub Combo1_Click()
                    LblPatients(10).Caption = ""
                    
                    'Example how to use ADO without the ADODC ActiveX Control
Dim cn As New ADODB.Connection
Dim rs As New ADODB.Recordset
Dim rs1 As New ADODB.Recordset
If cn.State = adStateOpen Then cn.Close
cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Doctors.mdb" & ";Persist Security Info=False;Jet OLEDB:Database Password=555"
                    'DataGrid Control will not Show Data without this but Flexgrid will
rs.CursorLocation = adUseClient
                    Select Case Combo1.ListIndex

                    Case 0
                    'SQL Example showing only FirstName,LastName,Balance Where Balance > 0
rs.Open "Select FirstName,LastName,Balance from Patients Where Balance > 0", cn, adOpenKeyset, adLockOptimistic
                    'Show Report Designer

                    'SQL Example to show the Sum of Outstanding Balance. It is better to use a seperate SQL for Calculations to eliminate confusion and complication in 1 single Statement.
rs1.Open "Select Sum(Balance) as ToTBalance from Patients", cn, adOpenKeyset, adLockOptimistic
                    LblPatients(10).Caption = "Total Balance: " & rs1.Fields("ToTBalance")

                    Case 1
                    'SQL Example showing only Patients with FirstName="Elissa"
rs.Open "Select * from Patients Where FirstName='Elissa'", cn, adOpenKeyset, adLockOptimistic
                    
                    Case 2
                    'SQL Example showing only Females Patients in the City of Beirut
rs.Open "Select FirstName,LastName,Gender,City,ID from Patients Where Gender='Female' and City='Beirut'", cn, adOpenKeyset, adLockOptimistic

                    Case 3
                    'SQL Example showing Date Report
rs.Open "Select FirstName,LastName,FirstVisitDate from Patients Where FirstVisitDate >= DateValue('1/1/2004')", cn, adOpenKeyset, adLockOptimistic

                    Case 4
                    'SQL Example All Patients Starting with FirstName "E"
rs.Open "Select * from Patients Where left(FirstName,1)='E'", cn, adOpenKeyset, adLockOptimistic

                    Case 5
                    'SQL Example using 'Like' showing All Patients with Diagnosis "Stomach Pain"
rs.Open "Select FirstName,LastName,Diagnosis from Patients", cn, adOpenKeyset, adLockOptimistic
rs.Filter = "Diagnosis Like '*Stomach Pain*'"
                    
                    End Select
'                   'Connect DataGrid2 to rs Recordset
Set DataGrid2.DataSource = rs
DataGrid2.Refresh
If Combo1.ListIndex = 5 Then DataGrid2.Columns("Diagnosis").Width = 5000

'Show Report Designer for Balance Report
If Combo1.ListIndex = 0 Then
If MsgBox("Show Print Preview?", vbYesNo) = vbNo Then Exit Sub
Set DataReport1.DataSource = rs
DataReport1.Show
End If

End Sub

'*******************************************************************************
'*********************   Additional DB Utilities  ***************
'*******************************************************************************
Private Sub LabelUtilities_Click(Index As Integer)
On Error GoTo 100
Select Case Index

Case 0 'Create DataBase
Dim DB As Database
'1) Create DataBase  name "Clients" using DAO 3.6 (see Project > References)
Set DB = CreateDatabase(App.Path & "\Clients", dbLangGeneral, dbEncrypt)

'to Create DataBase with PWD
'Set DB = CreateDatabase(App.Path & "\Clients", dbLangGeneral & ";pwd=555", dbEncrypt)

'2) Create Tables and Fields at once using DAO and the Microsoft Jet SQL
'                  Table "ClientsInfo"     FirstName (Text)     LastName (Text)     Age(No.) Address (Memo)
DB.Execute "CREATE TABLE ClientsInfo " & "(FirstName CHAR (50), LastName CHAR (50), Age INT, Address NOTE);"

DB.Close 'This Line is Optional
MsgBox "DataBase Created. DataBase name 'Clients' and Table name 'ClientsInfo'"
End Select
Exit Sub

100
MsgBox Err.Description

End Sub



'*******************************************************************************
'*************************  Non Related DataBase Code Section  *********************
'*******************************************************************************
Private Sub Form_Load()
'PCurrtPWd = ""
SSTab1.Tab = 0
For i = 1 To 4
SSTab1.TabEnabled(i) = False
Next
mUtilities.Enabled = False
DTPicker1.Value = Date

'This Retrieves the current DB Password from DB_Pwd.text and show it on LblPWD Label
If Dir(App.Path & "\DB_Pwd.text") <> "" Then
    Open App.Path & "\DB_Pwd.text" For Input As #1
    Input #1, pRecord
    LblPWD.Caption = pRecord
    Close #1
End If
End Sub
Private Sub PSetSStab()
SSTab1.TabVisible(0) = False
For i = 1 To 4
SSTab1.TabEnabled(i) = True
Next
SSTab1.Tab = 1
End Sub

Private Sub ImageHelp_Click(Index As Integer)
ShellExecute Me.hwnd, "open", App.Path & "\DataBaseTutotial.htm", ByVal 0&, "", 3

End Sub


Private Sub mAbout1_Click()
Form38.Show
End Sub

Private Sub mBackupDataBase_Click()
Call pBackupDB
End Sub

Private Sub mHelp_Click()
ShellExecute Me.hwnd, "open", App.Path & "\DataBaseTutotial.htm", ByVal 0&, "", 3

End Sub

Private Sub mRestoreDataBase_Click()
If Dir(App.Path & "\DoctorsBackup.mdb") = "" Then
MsgBox "Backup the DataBase First."
Exit Sub
End If

Call pRestoreDB
End Sub




Private Sub txtFields_Change(Index As Integer)
'I have put this Code here not on Adodc1_MoveComplete because of Microsoft bug.
'you get errors If you switch between ActiveX Data Object Reference Versions.
On Error Resume Next
If Index = 5 Then
Adodc1.Caption = "Record No. " & (Adodc1.Recordset.AbsolutePosition) & "/" & Adodc1.Recordset.RecordCount & " :  " & Adodc1.Recordset.Fields("FirstName").Value & " " & Adodc1.Recordset.Fields("LastName").Value
End If
End Sub

Private Sub txtFields_Click(Index As Integer)
SortStr = txtFields(Index).DataField
End Sub

Private Sub txtFields_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
If Len(txtFields(Index).Text) = 1 Then
txtFields(Index) = UCase(txtFields(Index))
txtFields(Index).SelStart = Len(txtFields(Index).Text)
End If

End Sub


Private Sub CmdFind_Click()
Form7.Show
End Sub
Private Sub mChangeDBPassword_Click()
Form15.Text1 = ""
Form15.Text2 = ""
Form15.Show
End Sub

Private Sub mCompactDataBase_Click()
Call pCompactDB
End Sub

Private Sub mVisitMyMedjugorjeWeb_Click()
ShellExecute Me.hwnd, "open", "http://geocities.com/medjugorjesite", ByVal 0&, "", 3
End Sub


Private Sub txtPaswd_Click()
txtPaswd.Text = ""
End Sub
'Private Sub Adodc1_Error(ByVal ErrorNumber As Long, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, fCancelDisplay As Boolean)
'                                                MsgBox Err.Description
'End Sub


'Private Sub Adodc1_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset) 'See Error Notes Below
'********************************************************************
'ActiveX Data Object Reference is set to 2.0.
'If you have VB6 Service Packs, You will get an Error.
'Change it to 2.5  (From  the Menu: Project > References)
'********************************************************************
'Compile Error Description: Procedure declaration does not match description of event or procedure having the same name.
'********************************************************************
                                               'On Error Resume Next
'Adodc1.Caption = "Record No. " & (Adodc1.Recordset.AbsolutePosition) & "/" & Adodc1.Recordset.RecordCount & " :  " & Adodc1.Recordset.Fields("FirstName").Value & " " & Adodc1.Recordset.Fields("LastName").Value
'End Sub

