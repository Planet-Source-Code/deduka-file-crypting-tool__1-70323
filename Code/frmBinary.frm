VERSION 5.00
Begin VB.Form frmBinary 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Binary File Maker"
   ClientHeight    =   7290
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8505
   FillColor       =   &H0000FFFF&
   Icon            =   "frmBinary.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7290
   ScaleWidth      =   8505
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd 
      Caption         =   "Remove all"
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
      Index           =   2
      Left            =   4440
      TabIndex        =   8
      Top             =   4200
      Width           =   2175
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Start Encrypt / Decrypt"
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
      Index           =   4
      Left            =   5520
      TabIndex        =   7
      Top             =   6000
      Width           =   2775
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Help ..."
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
      Index           =   3
      Left            =   6720
      TabIndex        =   6
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Remove selected"
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
      Index           =   1
      Left            =   2160
      TabIndex        =   5
      Top             =   4200
      Width           =   2175
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Top             =   240
      Width           =   8175
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Add JPEGS ..."
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
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   4200
      Width           =   1935
   End
   Begin VB.CommandButton cmd 
      Caption         =   "Close"
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
      Index           =   5
      Left            =   5520
      TabIndex        =   1
      Top             =   6600
      Width           =   2775
   End
   Begin BinFiles.ProgressBar ProgressBar1 
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   873
      C_ForeColor     =   16711680
      C_BackColor     =   4210752
      C_PercColor     =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lbl_status 
      Caption         =   "Status: Ready."
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
      Left            =   120
      TabIndex        =   4
      Top             =   5400
      Width           =   8175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBinary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Click(Index As Integer)
Select Case Index
  Case 0
    Dim Dialog As New Dialog
    Dialog.VBGetOpenFileName "Filename", , , True, , , "All|*.*", 1, , , , , 0
  Case 1
  
  Case 2
  
  Case 3
  
  Case 4
  
  Case 5
    Unload Me
End Select
End Sub



Private Sub cmdStart_Click()
  lblStatus = "Working..."
  lblStatus.Refresh
  
  cmd1.Enabled = False
  For i = 1 To 100
    Crypt (file.List(i - 1))
    ProgressBar1.Progress CInt(i)
    DoEvents
  Next
  cmd1.Enabled = True
  
  lblStatus = "Done."
End Sub

Private Sub Crypt(file As String)
  F = FreeFile
  Open AppPath & "\TEST_IMAGES\" & file For Binary As F
  
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




