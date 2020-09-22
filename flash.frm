VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Flash Player"
   ClientHeight    =   7140
   ClientLeft      =   3060
   ClientTop       =   2100
   ClientWidth     =   9420
   Icon            =   "flash.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "flash.frx":0442
   ScaleHeight     =   7140
   ScaleWidth      =   9420
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Exit"
      Height          =   330
      Left            =   8370
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   420
      Width           =   960
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Stop"
      Height          =   315
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   450
      Width           =   2340
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6210
      Left            =   105
      TabIndex        =   2
      Top             =   840
      Width           =   9210
      ExtentX         =   16245
      ExtentY         =   10954
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Browse &  Play"
      Height          =   315
      Left            =   135
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   105
      Width           =   2340
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " Flash Player v.1.0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   7395
      TabIndex        =   1
      Top             =   120
      Width           =   1920
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This code written by Gürkan Tümer ® 2000
'contact:albaarn@yahoo.com
Option Explicit
Private Sub Command1_Click()
    On Error GoTo err
        With CommonDialog1
            .Filter = "Flash Movies(*.swf)|*.swf|All Files (*.*)|*.*"
            .FilterIndex = 1
            .ShowOpen
        End With
            Label1.Caption = CommonDialog1.FileName
            WebBrowser1.Visible = True
            Form1.Height = 7515
    Call openfile
    Exit Sub
err:
    Exit Sub
End Sub
Private Sub openfile()
        WebBrowser1.Navigate (Label1.Caption)
End Sub
Private Sub Command2_Click()
        WebBrowser1.Stop: WebBrowser1.Navigate ("")
        WebBrowser1.Visible = False
        Form1.Height = 1255
End Sub
Private Sub Command3_Click()
       End
End Sub
Private Sub Form_Load()
        Form1.Height = 1255
        WebBrowser1.Visible = False
End Sub
