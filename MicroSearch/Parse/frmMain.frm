VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Parse"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8295
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
   ScaleHeight     =   4575
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CD1 
      Left            =   2040
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Load"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   375
      Left            =   7440
      TabIndex        =   8
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      Height          =   375
      Left            =   6600
      TabIndex        =   7
      Top             =   4080
      Width           =   735
   End
   Begin VB.ListBox List3 
      Height          =   3375
      Left            =   6600
      TabIndex        =   6
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   375
      Left            =   4200
      TabIndex        =   5
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Save"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   4080
      Width           =   735
   End
   Begin VB.ListBox List2 
      Height          =   3375
      Left            =   3360
      TabIndex        =   3
      Top             =   600
      Width           =   3135
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   345
      Left            =   120
      TabIndex        =   2
      Top             =   150
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   375
      Left            =   5640
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "frmMain.frx":0442
      Left            =   120
      List            =   "frmMain.frx":051B
      TabIndex        =   0
      Top             =   600
      Width           =   3135
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function InitCommonControls Lib "Comctl32.dll" () As Long
Private Sub Form_Initialize()
Dim X As Long
X = InitCommonControls
End Sub
Private Sub Command2_Click()
On Error Resume Next
With CD1
        .CancelError = True
        .DialogTitle = "Save list..."
        .Filter = "Untitled (*.txt)|*.txt"
        .ShowSave
    
        MyFile = FreeFile
        Open .FileName For Output As MyFile
        n = 0
        
        While n <> List2.ListCount
            Print #MyFile, List2.List(n)
            n = n + 1
        Wend
        Close MyFile
End With

error:
Exit Sub
End Sub
Private Sub Command3_Click()
List2.Clear
List3.Clear
End Sub
Private Sub Command4_Click()
On Error Resume Next
With CD1
        .CancelError = True
        .DialogTitle = "Save list..."
        .Filter = "Untitled (*.txt)|*.txt"
        .ShowSave
    
        MyFile = FreeFile
        Open .FileName For Output As MyFile
        n = 0
        
        While n <> List3.ListCount
            Print #MyFile, List3.List(n)
            n = n + 1
        Wend
        Close MyFile
End With

error:
Exit Sub
End Sub
Private Sub Command6_Click()
On Error GoTo error
Dim MyFile
Dim FreeFile
Dim Data$
With CD1
        .DialogTitle = "Load list..."
        .Filter = "Untitled(*.txt)|*.txt"
        .ShowOpen
        MyFile = FreeFile
        Open .FileName For Input As MyFile
            While Not EOF(MyFile)
            Input #MyFile, Data$
            List1.AddItem Data$
                            Wend
        Close MyFile
        .FileName = ""
End With

error:
Exit Sub
End Sub
Private Sub Command1_Click()
On Error Resume Next
Dim Hit As String
Dim Hit2 As String
Dim X As Integer
For X = 0 To List1.ListCount
If List1.ListIndex = List1.ListCount - 1 Then List1.ListIndex = 0: Exit Sub
'Text1.Text = List1.Text
'AllData = Text1.Text
'Hit = Get_Between(1, Text1, "<option value=""AL"">", "</option>'")
'Hit = Get_Between(1, Text1, "<option value=""", ">")
Hit = Get_Between(1, List1.Text, "<option value=""", """>")
Text1.Text = Hit: List2.AddItem Hit
Hit2 = Get_Between(1, List1.Text, """>", "</option>")
Text1.Text = Hit & " : " & Hit2
List3.AddItem Hit2
If List1.ListIndex = 0 Then
List1.ListIndex = 1
Else
List1.ListIndex = List1.ListIndex + 1
End If
Next X
'<option value="AL">Alabama</option>'
End Sub
Private Sub List1_Click()
Text1.Text = List1.Text
End Sub
Private Sub List2_DblClick()
List2.RemoveItem List2.ListIndex
End Sub
Private Sub List3_DblClick()
List3.RemoveItem List3.ListIndex
End Sub
