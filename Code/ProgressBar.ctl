VERSION 5.00
Begin VB.UserControl ProgressBar 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FF8080&
   BackStyle       =   0  'Transparent
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3810
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   254
   Begin VB.PictureBox pic 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BeginProperty DataFormat 
         Type            =   0
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1050
         SubFormatType   =   0
      EndProperty
      FillColor       =   &H00C00000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   120
      ScaleHeight     =   24
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   229
      TabIndex        =   0
      Top             =   120
      Width           =   3495
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "100 %"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   0
         TabIndex        =   1
         Top             =   120
         Width           =   3495
      End
   End
End
Attribute VB_Name = "ProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False



Private Sub UserControl_Resize()
  pic.left = 0
  pic.top = 0
 
  pic.Width = UserControl.ScaleWidth
  pic.Height = UserControl.ScaleHeight
  lbl.Width = UserControl.ScaleWidth

End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
  pic.ForeColor = PropBag.ReadProperty("C_ForeColor", &H8000000F)
  pic.BackColor = PropBag.ReadProperty("C_BackColor", &H8000000F)
  lbl.ForeColor = PropBag.ReadProperty("C_PercColor", &HFF0000)
  Set lbl.Font = PropBag.ReadProperty("Font", Ambient.Font)

End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
  Call PropBag.WriteProperty("C_ForeColor", C_ForeColor, &H8000000F)
  Call PropBag.WriteProperty("C_BackColor", C_BackColor, &H8000000F)
  Call PropBag.WriteProperty("C_PercColor", C_PercColor, &HFF0000)
  Call PropBag.WriteProperty("Font", lbl.Font, Ambient.Font)


End Sub

Public Property Get C_ForeColor() As OLE_COLOR
  C_ForeColor = pic.ForeColor
End Property
Public Property Let C_ForeColor(ByVal NewValue As OLE_COLOR)
  pic.ForeColor = NewValue
  PropertyChanged "C_ForeColor"
End Property

Public Property Get C_BackColor() As OLE_COLOR
  C_BackColor = pic.BackColor
End Property
Public Property Let C_BackColor(ByVal NewValue As OLE_COLOR)
  pic.BackColor = NewValue
  PropertyChanged "C_BackColor"
End Property

Public Property Get C_PercColor() As OLE_COLOR
  C_PercColor = lbl.ForeColor
End Property
Public Property Let C_PercColor(ByVal NewValue As OLE_COLOR)
  lbl.ForeColor = NewValue
  PropertyChanged "C_PercColor"
End Property

Public Property Get Font() As Font
  Set Font = lbl.Font
End Property
Public Property Set Font(ByVal NewValue As Font)
  Set lbl.Font = NewValue
  UserControl_Resize
  PropertyChanged "Font"
End Property

Public Sub Progress(Value As Integer)
On Error Resume Next
  pic.Cls
  pic.ScaleMode = 0
  pic.ScaleWidth = 100
  pic.ScaleHeight = 10
  pic.CurrentY = 2
  pic.CurrentX = pic.ScaleWidth / 2 - (pic.ScaleWidth / 15)
  'pic.Print Trim(Str(Value)) & "%"
  
  pic.Line (0, 0)-(Value, pic.ScaleHeight), C_ForeColor, BF
  pic.Refresh
  
  'pic.Print Trim(Str(Value)) & "%"
  lbl = Trim(Str(Value)) & "%"
  lbl.Refresh
  
End Sub
