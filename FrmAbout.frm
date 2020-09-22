VERSION 5.00
Begin VB.Form FrmAbout 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2700
   ClientLeft      =   45
   ClientTop       =   45
   ClientWidth     =   3975
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   HasDC           =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "Okay!"
      Height          =   375
      Left            =   1470
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Click here for more software!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   3
      Left            =   900
      TabIndex        =   4
      ToolTipText     =   "http://members.xoom.com/devsfort/index.html"
      Top             =   2280
      Width           =   2175
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   690
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Caption         =   "Dev Fade"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Index           =   0
      Left            =   1470
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   450
      Left            =   240
      Picture         =   "FrmAbout.frx":0000
      Top             =   240
      Width           =   450
   End
End
Attribute VB_Name = "FrmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWMAXIMIZED = 3

Private Sub cmd_Click()
Unload Me
End Sub

Private Sub Form_Load()
lbl(1).Caption = "Version: " & App.Major & "." & App.Minor & App.Revision: lbl(2).Caption = App.LegalCopyright
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set FrmAbout = Nothing
End Sub

Private Sub lbl_Click(Index As Integer)
Dim go
If Index = 3 Then
go = ShellExecute(Me.hwnd, vbNullString, "http://members.xoom.com/devsfort/index.html", vbNullString, "c:\", SW_SHOWNORMAL)
End If
End Sub
