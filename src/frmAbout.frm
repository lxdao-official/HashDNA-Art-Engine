VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3840
   ClientLeft      =   7005
   ClientTop       =   3990
   ClientWidth     =   7545
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2650.437
   ScaleMode       =   0  'User
   ScaleWidth      =   7085.146
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   345
      Left            =   6120
      TabIndex        =   6
      Top             =   3120
      Width           =   1140
   End
   Begin VB.TextBox txtDonate 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   735
      Left            =   240
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   5
      Text            =   "frmAbout.frx":7E6A
      Top             =   2640
      Width           =   4815
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":7E73
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   345
      Left            =   6120
      TabIndex        =   0
      Top             =   2640
      Width           =   1140
   End
   Begin VB.Label lblWebsite 
      Caption         =   "https://lxdao.io"
      BeginProperty Font 
         Name            =   "ËÎÌו"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   1080
      TabIndex        =   7
      Top             =   2160
      Width           =   3885
   End
   Begin VB.Image Image1 
      Height          =   2175
      Left            =   5160
      Picture         =   "frmAbout.frx":FCDD
      Stretch         =   -1  'True
      Top             =   120
      Width           =   2205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   6986.545
      Y1              =   1687.583
      Y2              =   1687.583
   End
   Begin VB.Label lblDescription 
      ForeColor       =   &H00000000&
      Height          =   930
      Left            =   1080
      TabIndex        =   2
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "LXDAO Art Eengine"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Command1_Click()
Clipboard.Clear
Clipboard.SetText WalletAddress
End Sub

Private Sub Form_Load()
    Me.Caption = "About " & App.Title
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    txtDonate.Text = "Donate" & vbCrLf & vbCrLf & "ENS: " & ENS & vbCrLf & WalletAddress
    lblDescription.Caption = Description
End Sub

Private Sub lblWebsite_Click()
    Shell "explorer " & "https://lxdao.io/", vbNormalFocus
End Sub
