VERSION 5.00
Begin VB.Form frmHelp 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Help"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8745
   FillColor       =   &H0000FFFF&
   Icon            =   "frmHelp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   396
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   583
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmd_close 
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
      Left            =   6960
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lbl 
      Caption         =   $"frmHelp.frx":1042
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   6
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   7935
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "A file that is not previously encrypted with this tool, will be encrypted."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Index           =   5
      Left            =   240
      TabIndex        =   6
      Top             =   4680
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   $"frmHelp.frx":1195
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Index           =   4
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   8055
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "By clicking on Start button, process starts."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   240
      TabIndex        =   4
      Top             =   3720
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "If a file is previously encrypted with this tool, it will be decrypted."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   240
      TabIndex        =   3
      Top             =   4200
      Width           =   6615
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "To encrypt/decrypt files, you have to choose one or more files, or folder where files you want to process exists."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   240
      TabIndex        =   2
      Top             =   2160
      Width           =   7935
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl 
      Caption         =   "File Crypting Tool is the crypting program created for encrypt/decrypt files."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   8175
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

txt = "" & _
"BRANK File Crypting Tool is the fast low-level crypting program created for crypt files used in Brank programs." & vbCrLf & _
"To encrypt/decrypt files, you have to choose one or more files, or folder where files you want to process exists." & vbCrLf & _
"Only selected files (when check-box is selected) will be processed." & vbCrLf & _
"When you choose some files or folders, all files in list-box are selected automaticaly." & vbCrLf & _
"If you want to exclude a file from encrypt/decrypt process, uncheck it." & vbCrLf & _
"By clicking on Start button, process starts." & vbCrLf & _
"If a file is previously encrypted with this tool, it will be decrypted." & vbCrLf & _
"A file that is not previously encrypted with this tool, will be encrypted." & vbCrLf



End Sub



Private Sub cmd_close_Click()
  Unload Me
End Sub

