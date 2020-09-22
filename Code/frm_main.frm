VERSION 5.00
Begin VB.Form frm_main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Crypting Tool"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9495
   FillColor       =   &H0000FFFF&
   Icon            =   "frm_main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   419
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   633
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_clear_list 
      Caption         =   "Clear list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6240
      TabIndex        =   14
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmd_help 
      Caption         =   "Help ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5400
      Width           =   1455
   End
   Begin VB.ListBox lst 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3435
      Left            =   120
      Sorted          =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   9255
   End
   Begin VB.CommandButton cmd_add_dir 
      Caption         =   "Add Dir ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   8
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmd_uncheck_all 
      Caption         =   "Remove all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   7
      Top             =   4800
      Width           =   1335
   End
   Begin VB.CommandButton cmd_start 
      Caption         =   "Start"
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
      Left            =   7920
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmd_about 
      Caption         =   "About ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmd_check_all 
      Caption         =   "Select all"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton cmd_add_files 
      Caption         =   "Add files ..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4800
      Width           =   1455
   End
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
      Left            =   7920
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin VB.Label lbl_files 
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1440
      TabIndex        =   13
      Top             =   4350
      Width           =   1575
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_folder 
      AutoSize        =   -1  'True
      Caption         =   "D:\ActualProjects\POSTCARDS_CRYPT\TEST_IMAGES\"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   6000
      Visible         =   0   'False
      Width           =   4230
   End
   Begin VB.Label Label2 
      Caption         =   "Click Help for additional information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   360
      Width           =   9255
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      Caption         =   "Choose files to Encrypt/Decrypt, and click Start to process..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9255
      WordWrap        =   -1  'True
   End
   Begin VB.Label lbl_status 
      AutoSize        =   -1  'True
      Caption         =   "Selected files:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   120
      TabIndex        =   3
      Top             =   4350
      Width           =   1215
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub lst_ItemCheck(Item As Integer)
  CountChecked
End Sub

Private Function CountChecked() As Integer
  Dim nChecked As Integer
  nChecked = 0
  For nItem = 0 To lst.ListCount - 1
    If lst.Selected(nItem) Then nChecked = nChecked + 1
  Next
  lbl_files = nChecked
  CountChecked = nChecked
End Function



'COMMANDS
Private Sub cmd_add_files_Click()
  Dim Files As Variant
  Files = GetFiles()
  If UBound(Files) <> -1 Then
    For nItem = 1 To UBound(Files)
      lst.AddItem Files(nItem)
      lst.Selected(lst.ListCount - 1) = True
    Next
  End If
  lst.Refresh
  CountChecked
End Sub

Private Sub cmd_add_dir_Click()
  Folder = BrowseFolders(hwnd, "Select a folder that contains JPEGS.", BrowseForFolders, , lbl_folder)
  lbl_folder = Folder
  sFile = Dir(lbl_folder)
  Do While sFile <> ""
    lst.AddItem (lbl_folder & sFile)
    lst.Selected(lst.ListCount - 1) = True
    sFile = Dir()
  Loop
End Sub

Private Sub cmd_check_all_Click()
  For nItem = 0 To lst.ListCount - 1
    lst.Selected(nItem) = True
  Next
  lst.Refresh
  CountChecked
End Sub

Private Sub cmd_uncheck_all_Click()
  For nItem = 0 To lst.ListCount - 1
    lst.Selected(nItem) = False
  Next
  CountChecked
End Sub

Private Sub cmd_clear_list_Click()
  lst.Clear
  CountChecked
End Sub

Private Sub cmd_start_Click()
  If CountChecked() = 0 Then
    MsgBox "You need at least one file selected to process."
  Else
    frmProcess.Show vbModal
  End If
End Sub

Private Sub cmd_close_Click()
  Unload Me
End Sub

Private Sub cmd_help_Click()
  frmHelp.Show vbModal
End Sub

Private Sub cmd_about_Click()
  frmAbout.Show vbModal
End Sub




