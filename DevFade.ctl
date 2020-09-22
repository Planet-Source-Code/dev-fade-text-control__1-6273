VERSION 5.00
Begin VB.UserControl Fade 
   AutoRedraw      =   -1  'True
   BackStyle       =   0  'Transparent
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2745
   ScaleHeight     =   1860
   ScaleWidth      =   2745
   ToolboxBitmap   =   "DevFade.ctx":0000
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HasDC           =   0   'False
      Height          =   1455
      Left            =   0
      ScaleHeight     =   1455
      ScaleWidth      =   2295
      TabIndex        =   0
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "Fade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'   Dev Fade OCX By Dev

'   http://www.brechin.clara.net/

Private m_bEnabled As Boolean
Dim txt As String, col1 As Long, col2 As Long

Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event Click()
Event Resize()


Public Sub FadeTxT(ByVal canvas As Object, ByVal start_x As Single, ByVal start_y As Single, ByVal txt As String)
Attribute FadeTxT.VB_MemberFlags = "40"
Dim r As Single, g As Single, b As Single
Dim blue1, red1, green1
Dim blue2, red2, green2
Dim txt_len As Integer
Dim i As Integer

blue1 = Int(col1 / 65536)
green1 = Int((col1 - (blue1 * 65536)) / 256)
red1 = col1 - (blue1 * 65536) - (green1 * 256)

blue2 = Int(col2 / 65536)
green2 = Int((col2 - (blue2 * 65536)) / 256)
red2 = col2 - (blue2 * 65536) - (green2 * 256)

    txt_len = Len(txt)
    dr = (red2 - red1) / (txt_len - 1)
    dg = (green2 - green1) / (txt_len - 1)
    db = (blue2 - blue1) / (txt_len - 1)
    r = red1
    g = green1
    b = blue1
    canvas.CurrentX = start_x
    canvas.CurrentY = start_y
    For i = 1 To txt_len
        canvas.ForeColor = RGB(r, g, b)
        canvas.Print Mid$(txt, i, 1);
        r = r + dr
        g = g + dg
        b = b + db
    Next i
End Sub

Public Sub Refresh()
    Pic1.Cls
    FadeTxT UserControl.Pic1, 20, 20, txt
    Pic1.Refresh
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   m_bRunTime = (UserControl.Ambient.UserMode)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)

    UserControl.Pic1.BackColor = PropBag.ReadProperty("BackColor", vbButtonFace)
    col1 = PropBag.ReadProperty("Color1", vbRed)
    col2 = PropBag.ReadProperty("Color2", vbBlack)
        
    UserControl.Pic1.FontBold = PropBag.ReadProperty("FontBold", 0)
    UserControl.Pic1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Set UserControl.Pic1.Font = PropBag.ReadProperty("Font", "Arial")

    UserControl.Pic1.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    txt = PropBag.ReadProperty("Caption", "Dev Fade")
    Pic1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
End Sub

Private Sub UserControl_Resize()
Pic1.Width = UserControl.Width
Pic1.Height = UserControl.Height
Pic1.Cls
FadeTxT UserControl.Pic1, 20, 20, txt
End Sub

Private Sub UserControl_Show()
FadeTxT UserControl.Pic1, 20, 20, txt
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)

    Call PropBag.WriteProperty("Color1", col1, vbRed)
    Call PropBag.WriteProperty("Color2", col2, vbBlack)
    Call PropBag.WriteProperty("BackColor", UserControl.Pic1.BackColor, vbButtonFace)
    
    Call PropBag.WriteProperty("FontBold", UserControl.Pic1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", UserControl.Pic1.FontItalic, 0)
    Call PropBag.WriteProperty("FontUnderline", Pic1.FontUnderline, 0)
    Call PropBag.WriteProperty("Font", UserControl.Pic1.Font, "Arial")
    
    Call PropBag.WriteProperty("BorderStyle", UserControl.Pic1.BorderStyle, 0)

    Call PropBag.WriteProperty("Caption", txt, "Dev Fade")
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whever the control can respond to user-generated events such as clicking."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    Pic1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property
'

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    txt = "Dev Fade"
    col1 = vbBlack
    col2 = vbRed
End Sub


Public Property Get Color1() As OLE_COLOR
Attribute Color1.VB_Description = "Returns/set the first text color in the fade."
  Color1 = col1
End Property

Public Property Let Color1(ByVal c As OLE_COLOR)
  col1 = c
  Call Refresh
End Property

Public Property Get Color2() As OLE_COLOR
Attribute Color2.VB_Description = "Returns/set the second text color in the fade."
  Color2 = col2
End Property

Public Property Let Color2(ByVal c As OLE_COLOR)
    col2 = c
    Call Refresh
End Property

Public Sub AboutBox()
Attribute AboutBox.VB_UserMemId = -552
    FrmAbout.Show vbModal
End Sub

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Pic1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Pic1.Font = New_Font
    PropertyChanged "Font"
    
    Pic1.Cls
    FadeTxT UserControl.Pic1, 20, 20, txt
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in the control."
    BackColor = Pic1.BackColor
End Property

Public Property Let BackColor(c As OLE_COLOR)
Pic1.BackColor = c
FadeTxT UserControl.Pic1, 20, 20, txt
End Property

Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Pic1.FontBold
    Call Refresh
End Property

Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Pic1.FontItalic
    Call Refresh
End Property

Public Property Get Caption() As Variant
Attribute Caption.VB_Description = "Returns/Sets text displayed in the control."
Attribute Caption.VB_UserMemId = 0
    Caption = txt
End Property

Public Property Let Caption(ByVal vNewValue As Variant)
  If Len(vNewValue) > 1 Then
    txt = vNewValue
    UserControl.Pic1.Cls
    FadeTxT UserControl.Pic1, 20, 20, txt
  Else
  MsgBox "Caption must contain 2 or more characters.", vbCritical, "Error"
  End If
End Property

Private Sub Pic1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Pic1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub Pic1_Click()
    RaiseEvent Click
End Sub

Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = Pic1.FontUnderline
    Call Refresh
End Property
