VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Micro-Search - [Source] 1.01"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6000
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Extras"
      Height          =   4335
      Left            =   7920
      TabIndex        =   43
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
      Begin VB.ListBox ListIDS 
         Height          =   645
         ItemData        =   "frmMain.frx":0442
         Left            =   120
         List            =   "frmMain.frx":051B
         TabIndex        =   45
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.ListBox States 
         Height          =   840
         ItemData        =   "frmMain.frx":0639
         Left            =   120
         List            =   "frmMain.frx":070F
         TabIndex        =   44
         Top             =   3360
         Visible         =   0   'False
         Width           =   1815
      End
   End
   Begin VB.Timer T1 
      Interval        =   1
      Left            =   8040
      Top             =   4680
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Time - Date"
      ForeColor       =   &H80000011&
      Height          =   1095
      Left            =   3600
      TabIndex        =   32
      Top             =   4800
      Width           =   4095
      Begin VB.TextBox TxtDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   2880
         TabIndex        =   34
         Text            =   "Date error"
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TxtTime 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00000080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   33
         Text            =   "Time error"
         Top             =   480
         Width           =   1095
      End
      Begin VB.Line Line27 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   3960
         Y1              =   360
         Y2              =   360
      End
      Begin VB.Line Line22 
         BorderColor     =   &H00FFFFFF&
         X1              =   120
         X2              =   3960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Line Line21 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   3960
         Y1              =   960
         Y2              =   960
      End
      Begin VB.Label lblButton1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "http://ycrack.net"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1200
         MouseIcon       =   "frmMain.frx":0A53
         MousePointer    =   99  'Custom
         TabIndex        =   35
         Top             =   480
         Width           =   1695
      End
      Begin VB.Line Line28 
         BorderColor     =   &H00000080&
         BorderWidth     =   2
         X1              =   120
         X2              =   3960
         Y1              =   360
         Y2              =   360
      End
   End
   Begin MSWinsockLib.Winsock Winsock2 
      Left            =   3960
      Top             =   2760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.whitepages.com"
      RemotePort      =   80
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   2640
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "www.whitepages.com"
      RemotePort      =   80
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Reverse Phone Number"
      ForeColor       =   &H00000000&
      Height          =   1215
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.CheckBox Check4 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AutoSave (New)"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   840
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         MaxLength       =   7
         TabIndex        =   4
         Text            =   "6634004"
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "706"
         Top             =   480
         Width           =   855
      End
      Begin VB.Label Command2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         Enabled         =   0   'False
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1680
         TabIndex        =   8
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Command1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Process"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   2520
         TabIndex        =   7
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Number: [7-Digits]"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   2175
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Area Code"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   5535
      Left            =   10200
      TabIndex        =   25
      Top             =   600
      Visible         =   0   'False
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   9763
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmMain.frx":0D5D
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Find a Person"
      Height          =   4335
      Left            =   3600
      TabIndex        =   13
      Top             =   120
      Width           =   4095
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmMain.frx":0DD9
         Left            =   2160
         List            =   "frmMain.frx":0EB2
         TabIndex        =   42
         Text            =   "State"
         Top             =   1680
         Width           =   1815
      End
      Begin VB.CheckBox Check5 
         BackColor       =   &H00C0C0C0&
         Caption         =   "AutoSave (New)"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   2040
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1815
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":1204
         Left            =   120
         List            =   "frmMain.frx":1206
         MultiSelect     =   2  'Extended
         TabIndex        =   26
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox Text5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Text            =   "30240"
         Top             =   1680
         Width           =   1935
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Begins With"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2160
         TabIndex        =   19
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Text            =   "Doe"
         Top             =   1080
         Width           =   1935
      End
      Begin VB.CheckBox Check1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0C0&
         Caption         =   "Begins With"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2160
         TabIndex        =   15
         Top             =   480
         Value           =   1  'Checked
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Text            =   "John"
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Cancel"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   28
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Command3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Process"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   27
         Top             =   2040
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "State or Province"
         Height          =   255
         Left            =   2160
         TabIndex        =   22
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label9 
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "City, zip, or postal code"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "Last Name *"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Caption         =   "First Name"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Address"
      ForeColor       =   &H00000000&
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   3375
      Begin VB.CheckBox Check3 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Append to list"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1680
         Width           =   1335
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1335
         IntegralHeight  =   0   'False
         ItemData        =   "frmMain.frx":1208
         Left            =   120
         List            =   "frmMain.frx":120A
         MultiSelect     =   2  'Extended
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Clear"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   41
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Save"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   40
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   3375
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   120
      Y1              =   3840
      Y2              =   6000
   End
   Begin VB.Line Line26 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3480
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Line Line24 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   3840
   End
   Begin VB.Line Line19 
      BorderColor     =   &H00FFFFFF&
      X1              =   1200
      X2              =   2280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Line Line18 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   1200
      X2              =   2280
      Y1              =   5760
      Y2              =   5760
   End
   Begin VB.Label LblBetter1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beta"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   360
      TabIndex        =   39
      Top             =   5280
      Width           =   2895
   End
   Begin VB.Label LblBetter2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Beta"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   240
      TabIndex        =   38
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label lblVersion2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "v1.01"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   2400
      TabIndex        =   37
      Top             =   5160
      Width           =   975
   End
   Begin VB.Label lblVersion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "v1.01"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2400
      TabIndex        =   36
      Top             =   5160
      Width           =   855
   End
   Begin VB.Line Line10 
      BorderColor     =   &H00FFFFFF&
      X1              =   6840
      X2              =   6960
      Y1              =   4680
      Y2              =   4560
   End
   Begin VB.Line Line9 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   3480
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line8 
      BorderWidth     =   2
      X1              =   120
      X2              =   3480
      Y1              =   4080
      Y2              =   4080
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   3480
      X2              =   3480
      Y1              =   6000
      Y2              =   3840
   End
   Begin VB.Line Line6 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   3480
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Micro-Search"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   360
      TabIndex        =   30
      Top             =   4800
      Width           =   3015
   End
   Begin VB.Shape Shape1 
      Height          =   1215
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   4680
      Width           =   3135
   End
   Begin VB.Label SB1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Micro-Search"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   4200
      Width           =   3375
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00C00000&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3840
      Width           =   15
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Micro-Search"
      BeginProperty Font 
         Name            =   "Impact"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   495
      Left            =   480
      TabIndex        =   31
      Top             =   4800
      Width           =   2655
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   6960
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line5 
      BorderWidth     =   2
      X1              =   6960
      X2              =   120
      Y1              =   4560
      Y2              =   4560
   End
   Begin VB.Line Line7 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   3480
      X2              =   3480
      Y1              =   6000
      Y2              =   3840
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   6840
      X2              =   3480
      Y1              =   4680
      Y2              =   4680
   End
   Begin VB.Line Line11 
      BorderWidth     =   2
      X1              =   6840
      X2              =   6960
      Y1              =   4680
      Y2              =   4560
   End
   Begin VB.Line Line23 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   120
      Y1              =   4080
      Y2              =   6000
   End
   Begin VB.Line Line25 
      BorderColor     =   &H00000080&
      BorderWidth     =   2
      X1              =   120
      X2              =   3480
      Y1              =   3840
      Y2              =   3840
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   3840
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-|---[Recent Update: 09/30/04]-----------------]---|
'-|---[Micro-Search v1.01 (build 0904)----------]---|
' |---[-----------------------------------------]---|
'-[_________________________________________________|
'__[PlanetSourceCode.Com]___________________________)
'-[Micro Search]-: Created by: Brian Matthews [...]|)
'-[YCrack.Net]-: (C) Copyright 2004 YCrack : -[tm]-/|
'-[Credits:  http://YCrack.Net][http://keithware.com]-/|
'__[Microsoft Murder_[INC]_________________________/|
'__[Re-release_Date]:__*[10/16/04]*________________/|
'____________________[Comments]: (None) ___________/|
'__Begin: [Declarations]___________________________/|
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Dim AreaCode As String    '________________________/|
Dim PhoneNumber As String '________________________/|
Dim AllData As String     '________________________/|
Dim FirstName As String
Dim LastName As String
Dim ZipCode As String
Dim StateID As String
Dim State As String
Private Sub Combo1_Click()
On Error GoTo HandleError
StateID = Combo1.ListIndex
ListIDS.ListIndex = StateID
Exit Sub
'ListIDS.ListIndex = Combo1.ListIndex
HandleError: MsgBox ("Unknown error, please try again"), vbCritical, "Unknown Problem - Micro Search"
End Sub
Private Sub Command3_Click()
On Error Resume Next
Winsock2.Close
Winsock2.Connect
End Sub
Private Sub Form_Activate()
On Error Resume Next
TxtTime.Text = Time
TxtDate.Text = Date
'Me.Width = GetSetting("MicroSearch", "frmMain", "Width", frmMain.Width)
'Me.Height = GetSetting("MicroSearch", "frmMain", "Height", frmMain.Height)
Check1.Value = GetSetting("MicroSearch", "Checkbox", "Check1", Check1.Value)
Check2.Value = GetSetting("MicroSearch", "Checkbox", "Check2", Check2.Value)
Check3.Value = GetSetting("MicroSearch", "Checkbox", "Check3", Check3.Value)
Check4.Value = GetSetting("MicroSearch", "Checkbox", "Check4", Check4.Value)
Check5.Value = GetSetting("MicroSearch", "Checkbox", "Check5", Check5.Value)
Text1.Text = GetSetting("MicroSearch", "Text", "Text1", Text1.Text)
Text2.Text = GetSetting("MicroSearch", "Text", "Text2", Text2.Text)
End Sub
'__End:   [Declarations]___________________________/|
'-[________________________________________________/|
'-[________________________________________________|)
' |---|-----------------------------------------|---|

Private Sub Form_Initialize()
Dim x As Long
x = InitCommonControls
End Sub
Private Sub Command1_Click()
If Text1.Text = "" Or Text2.Text = "" Then SB1.Caption = "Invalid Phone Number...": Exit Sub
If Len(Text1.Text) < 3 Or Len(Text2.Text) < 7 Then SB1.Caption = "Invalid Phone Number Length...": Exit Sub
Winsock1.Close
Winsock1.Connect
AreaCode = Text1.Text
PhoneNumber = Text2.Text
Command1.Enabled = False
Command2.Enabled = True
SB1.Caption = "Processing Phone Number: " & "(" & Text1 & ")" & " " & Text2 & "..."
Label4.Width = Label3.Width / 4
Label5.Caption = "25%"
End Sub
Private Sub Command1_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command1.BackStyle = 1
End Sub
Private Sub Command1_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command1.BackStyle = 0
End Sub
Private Sub Command2_Click()
Winsock1.Close
SB1.Caption = "Process Canceled..."
Command2.Enabled = False
Command1.Enabled = True
Label4.Width = 0
End Sub
Private Sub Command2_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command2.BackStyle = 1
End Sub
Private Sub Command2_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)
Command2.BackStyle = 0
End Sub
Private Sub Form_Load()
On Error Resume Next
If Check5.Value = vbChecked Then
Text3.Text = GetSetting("MicroSearch", "Text", "Text3", Text3.Text)
Text4.Text = GetSetting("MicroSearch", "Text", "Text4", Text4.Text)
Text5.Text = GetSetting("MicroSearch", "Text", "Text5", Text5.Text)
Combo1.Text = GetSetting("MicroSearch", "Combo", "Combo1", Combo1.Text)
Else
End If
'Exit Sub
End Sub
Private Sub Form_Unload(Cancel As Integer)
'On Error Resume Next
If Check5.Value = vbChecked Then
SaveSetting "MicroSearch", "Text", "Text3", Text3.Text
SaveSetting "MicroSearch", "Text", "Text4", Text4.Text
SaveSetting "MicroSearch", "Text", "Text5", Text5.Text
SaveSetting "MicroSearch", "Combo", "Combo1", Combo1.Text
Else
End If
SaveSetting "MicroSearch", "frmMain", "Width", frmMain.Width
SaveSetting "MicroSearch", "frmMain", "Height", frmMain.Height
SaveSetting "MicroSearch", "Checkbox", "Check1", Check1.Value
SaveSetting "MicroSearch", "Checkbox", "Check2", Check2.Value
SaveSetting "MicroSearch", "Checkbox", "Check3", Check3.Value
SaveSetting "MicroSearch", "Checkbox", "Check4", Check4.Value
SaveSetting "MicroSearch", "Checkbox", "Check5", Check5.Value
SaveSetting "MicroSearch", "Text", "Text1", Text1.Text
SaveSetting "MicroSearch", "Text", "Text2", Text2.Text
Winsock1.Close: End
'Exit Sub
End Sub
Private Sub Label14_Click()
On Error Resume Next
List1.Clear
End Sub
Private Sub lblButton1_Click()
On Error Resume Next
OpenURL "http://www.ycrack.net"
End Sub
Private Sub RTB1_Change()
On Error Resume Next
frmData.RTB1.Text = frmMain.RTB1.Text
End Sub
Private Sub RTB1_DblClick()
On Error Resume Next
frmData.Show
End Sub
Private Sub T1_Timer()
TxtTime.Text = Time
TxtDate.Text = Date
End Sub
Private Sub Text1_Change()
On Error Resume Next
If IsNumeric(Text1.Text) = False Then SendKeys "{BackSpace}": Exit Sub
If Len(Text1.Text) = 3 Then Text2.SetFocus
End Sub
Private Sub Text2_Change()
On Error Resume Next
If IsNumeric(Text2.Text) = False Then SendKeys "{BackSpace}": Exit Sub
End Sub
Private Sub Winsock1_Close()
'n Error Resume Next
Label4.Width = Label3.Width
Label5.Caption = "100%"
Command1.Enabled = True
Command2.Enabled = False
If InStr(1, AllData, "Search Information:") Then
SB1.Caption = "Processing Complete...": ProcessAddress: GetSearchTotal
Else
SB1.Caption = "No Results..."
'List1.AddItem " ----------------- " & "No Results" & " ------------------ "
List1.AddItem " ---- " & "[No Results: " & "(" & Text1 & ")" & " " & Text2 & "]" & " ---- "
Exit Sub
End If
Exit Sub
End Sub
Private Sub Winsock1_Connect()
On Error Resume Next
AllData = ""
Winsock1.SendData "GET http://www.whitepages.com/search/Reverse_Phone?npa=" & AreaCode & "&phone=" & PhoneNumber & vbCrLf & vbCrLf
Label4.Width = Label3.Width / 2
Label5.Caption = "50%"
End Sub
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
Dim Data As String
Dim EXtra As Integer
Winsock1.GetData Data
AllData = AllData & Data
EXtra = Label3.Width / 100
Label4.Width = EXtra * 75
Label5.Caption = "75%"
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
On Error Resume Next
Winsock1.Close
Command1.Enabled = True
Command2.Enabled = False
SB1.Caption = Description
Label4.Width = 0
End Sub
Private Sub ProcessAddress()
On Error Resume Next
Dim Needed As String
Dim FullName As String
If Check3.Value = vbChecked Then
List1.AddItem " ----------------- "
Else
List1.Clear
End If
Needed = Get_Between(1, AllData, "<span class=""text"" style=""line-height:13pt;"">", "<br></span>")
FullName = Get_Between(1, AllData, "<img src=""/static/common/trans.gif"" width=""1"" height=""12"" border=""0""><a href=""javascript:moreInfo(1);""", "</td></tr>")
FullName = Replace(FullName, vbLf, "")
FullName = Replace(FullName, "</a>", "")
FullName = Replace(FullName, ">", "")
FullName = Replace(FullName, Chr(9), "")
FullName = Replace(FullName, ",", ", ")
FullName = Replace(FullName, "&amp;", "&")
List1.AddItem FullName
List1.AddItem Get_Item(1, Needed, "<br>")
List1.AddItem Get_Item(2, Needed, "<br>")
List1.AddItem Get_Item(3, Needed, "<br>")
List1.ListIndex = List1.ListCount - 1
End Sub
Private Sub ProcessInfo()
On Error Resume Next
Dim Needed As String
Dim FirstName As String
Dim LastName As String
Dim ZipCode As String
Dim StateID As String
List2.AddItem " ----------------- "
Needed = Get_Between(1, AllData, "<span class=""text"" style=""line-height:13pt;"">", "<br></span>")
FullName = Get_Between(1, AllData, "<img src=""/static/common/trans.gif"" width=""1"" height=""12"" border=""0""><a href=""javascript:moreInfo(1);""", "</td></tr>")
FullName = Replace(FullName, vbLf, "")
FullName = Replace(FullName, "</a>", "")
FullName = Replace(FullName, ">", "")
FullName = Replace(FullName, Chr(9), "")
FullName = Replace(FullName, ",", ", ")
FullName = Replace(FullName, "&amp;", "&")
List2.AddItem FullName
List2.AddItem Get_Item(1, Needed, "<br>")
List2.AddItem Get_Item(2, Needed, "<br>")
List2.AddItem Get_Item(3, Needed, "<br>")
List2.ListIndex = List2.ListCount - 1
End Sub
Private Sub GetSearchTotal()
'On Error Resume Next
Dim Needed As String
'Needed = Get_Between(1, AllData, "<td align=""left"" valign=""top"" class=""subtext=", "Total Results</td>") '"</td>")
Needed = Get_Between(1, AllData, "<td align=""left"" valign=""top""><select name=""alpha_limit"" style=""background-color:#E8E8E8;font-size:11px;"" onchange=""javascript:submit();""><option value="""""">", "</option>") '"</td>")
'Search took
'Needed = Replace(Needed, "&quot;", "")
'Needed = Replace(Needed, ">", "")
Me.Caption = Me.Caption & " " & Needed
List2.AddItem Needed
End Sub
Private Sub Winsock2_Close()
On Error Resume Next
Label4.Width = Label3.Width
Label5.Caption = "100%"
Command1.Enabled = True
Command2.Enabled = False
If InStr(1, AllData, "Search Information:") Then
SB1.Caption = "Processing Complete...": ProcessInfo
Else
SB1.Caption = "No Results...": List2.AddItem " -------------------- " & vbCrLf & "No Results": Exit Sub
End If
Exit Sub
End Sub
Private Sub Winsock2_Connect()
AllData = ""
'http://whitepages.com/search/Find_Person?firstname_begins_with=1&firstname=&name=matthews&city_zip=30240&state_id=GA
Winsock2.SendData "GET http://whitepages.com/search/Find_Person?firstname=" & Text3.Text & "&name=" & Text4.Text & "&city_zip=" & Text5.Text & "&state_id=" & ListIDS.Text & vbCrLf & vbCrLf
'http://whitepages.com/search/Find_Person?firstname_begins_with=1&name_begins_with=0&firstname=&name=matthews&city_zip=&state_id=GA
Label4.Width = Label3.Width / 2
Label5.Caption = "50%"
End Sub
Private Sub Winsock2_DataArrival(ByVal bytesTotal As Long)
'On Error Resume Next
Dim Data As String
Winsock2.GetData Data
AllData = AllData & Data
RTB1.Text = Data
End Sub
Private Sub Winsock2_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
'On Error Resume Next
Winsock2.Close
Command1.Enabled = True
Command2.Enabled = False
SB1.Caption = Description
Label4.Width = 0
End Sub
