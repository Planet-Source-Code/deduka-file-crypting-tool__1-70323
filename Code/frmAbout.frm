VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7065
   FillColor       =   &H0000FFFF&
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   471
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
      Left            =   5520
      TabIndex        =   0
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "File Crypting Tool"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lbl_status 
      AutoSize        =   -1  'True
      Caption         =   "This program was created by deduka, and it's fee to use. :-)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   6060
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_close_Click()
  Unload Me
End Sub


