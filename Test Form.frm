VERSION 5.00
Begin VB.Form Frm1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dev Fade OCX Test"
   ClientHeight    =   2145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2145
   ScaleWidth      =   4470
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      Height          =   375
      Left            =   3480
      TabIndex        =   20
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "About"
      Height          =   375
      Left            =   3480
      TabIndex        =   19
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colour 2"
      Height          =   975
      Index           =   1
      Left            =   1800
      TabIndex        =   2
      Top             =   960
      Width           =   1575
      Begin VB.CommandButton Col2 
         BackColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   7
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H00808080&
         Height          =   255
         Index           =   6
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H00C000C0&
         Height          =   255
         Index           =   5
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H0000C0C0&
         Height          =   255
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H0000C000&
         Height          =   255
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Col2 
         BackColor       =   &H000000C0&
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Colour 1"
      Height          =   975
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   1575
      Begin VB.CommandButton Col1 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Index           =   7
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H00C0C0C0&
         Height          =   255
         Index           =   6
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H00FF00FF&
         Height          =   255
         Index           =   5
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H00000000&
         Height          =   255
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   600
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Index           =   3
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H0000FF00&
         Height          =   255
         Index           =   2
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H00FF0000&
         Height          =   255
         Index           =   1
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   255
      End
      Begin VB.CommandButton Col1 
         BackColor       =   &H000000FF&
         Height          =   255
         Index           =   0
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   255
      End
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Dev Fade"
      Top             =   600
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Add Dev Fade OCX Here....."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   21
      Top             =   120
      Width           =   4215
   End
End
Attribute VB_Name = "Frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Col1_Click(Index As Integer)
 Fade1.Color1 = Col1(Index).BackColor
End Sub

Private Sub Col2_Click(Index As Integer)
 Fade1.Color2 = Col2(Index).BackColor
End Sub

Private Sub Command1_Click()
 Fade1.aboutbox
End Sub

Private Sub Command2_Click()
 Unload Me
End Sub

Private Sub Text1_Change()
 If Len(Text1) > 1 Then Fade1.Caption = Text1
End Sub
