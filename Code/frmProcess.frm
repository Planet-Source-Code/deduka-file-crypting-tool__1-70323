VERSION 5.00
Begin VB.Form frmProcess 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Crypting process..."
   ClientHeight    =   2355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   ControlBox      =   0   'False
   FillColor       =   &H0000FFFF&
   Icon            =   "frmProcess.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   157
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_cancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2565
      TabIndex        =   1
      Top             =   1680
      Width           =   1935
   End
   Begin FileCryptingTool.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   225
      TabIndex        =   0
      Top             =   480
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   873
      C_ForeColor     =   16711680
      C_BackColor     =   8421504
      C_PercColor     =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_status 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1080
      Width           =   6615
   End
End
Attribute VB_Name = "frmProcess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Activated As Boolean
Dim Aborted As Boolean

Private Sub Form_Load()
  Activated = False
  Aborted = False
  '
End Sub


Private Sub Form_Activate()
  On Error GoTo ErrorHandler
  Me.Refresh
    For nItem = 0 To frm_main.lst.ListCount - 1
      DoEvents
      If Aborted Then Exit For
      If frm_main.lst.Selected(nItem) Then
        lbl_status = "Processing file " & frm_main.lst.List(nItem)
        Crypt frm_main.lst.List(nItem)
      End If
      ProgressBar1.Progress CInt((nItem + 1) / frm_main.lst.ListCount * 100)
    Next
    
  If Not Aborted Then
    ProgressBar1.Progress 100
    MsgBox "All files processed successfuly."
  End If
  Unload Me
  Exit Sub
  
ErrorHandler:
MsgBox "Error processing."
Unload Me

End Sub


Private Sub cmd_cancel_Click()
  Aborted = True
End Sub


Private Sub Crypt(FileName As String)
  F = FreeFile
  Open FileName For Binary As F
  
  ReDim DATA1(LOF(F) - 1)
  ReDim DATA2(LOF(F) - 1)
  Get F, 1, DATA1
  DATA2 = DATA1
  
  ToByte = IIf(UBound(DATA1) < 2050, UBound(DATA1) / 3, 1024)
  
  For i = 0 To ToByte
    DATA1(i) = DATA2(UBound(DATA1) - i)
  Next
  For i = 0 To ToByte
    DATA1(UBound(DATA1) - i) = DATA2(i)
  Next
  
  Put F, 1, DATA1
  Close F
End Sub


