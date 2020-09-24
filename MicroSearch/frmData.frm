VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmData 
   Caption         =   "Data View Window"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7710
   Icon            =   "frmData.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox RTB1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   10398
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmData.frx":0442
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
On Error Resume Next
Me.Width = GetSetting("MicroSearch", "frmData", "Width", frmData.Width)
Me.Height = GetSetting("MicroSearch", "frmData", "Height", frmData.Height)
End Sub
Private Sub Form_Resize()
On Error Resume Next
RTB1.Height = Me.Height - 750
RTB1.Width = Me.Width - 375
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
SaveSetting "MicroSearch", "frmData", "Width", frmData.Width
SaveSetting "MicroSearch", "frmData", "Height", frmData.Height
End Sub
