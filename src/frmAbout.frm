VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   3840
   ClientLeft      =   6585
   ClientTop       =   4875
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
   Begin VB.CommandButton cmdCopyWalletAddress 
      Cancel          =   -1  'True
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   345
      Left            =   6120
      TabIndex        =   5
      Tag             =   "11"
      ToolTipText     =   "Copy the wallet address"
      Top             =   3120
      Width           =   1140
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      Picture         =   "frmAbout.frx":7F6A
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
      Tag             =   "1"
      Top             =   2640
      Width           =   1140
   End
   Begin VB.Label lblDonate 
      ForeColor       =   &H00000000&
      Height          =   1050
      Left            =   255
      TabIndex        =   7
      Tag             =   "1"
      Top             =   2640
      Width           =   5085
   End
   Begin VB.Label lblWebsite 
      Caption         =   "https://hashdna.art"
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
      TabIndex        =   6
      Top             =   2160
      Width           =   3885
   End
   Begin VB.Image imgDonate 
      Height          =   2175
      Left            =   5160
      Picture         =   "frmAbout.frx":FED4
      Stretch         =   -1  'True
      Tag             =   "01"
      ToolTipText     =   "Scan the wallet QR code"
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
      Tag             =   "1"
      Top             =   1125
      Width           =   3885
   End
   Begin VB.Label lblTitle 
      Caption         =   "HashDNA Art Eengine"
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
      Tag             =   "1"
      Top             =   780
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2022 LXDAO

'This file is part of HashDNA Art Eengine.
'
'HashDNA Art Eengine is free software: you can redistribute it and/or modify it under the terms
'of the'GNU General Public License as published by the Free Software Foundation, either
'version 3 of the License, or (at your option) any later version.
'
'HashDNA Art Eengine is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY;
'without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See
'the GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License along with Foobar. If not,
'see <https://www.gnu.org/licenses/>.

Option Explicit

Private Sub cmdCopyWalletAddress_Click()
    Clipboard.Clear
    Clipboard.SetText WalletAddress
End Sub

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    TranslateForm Me
    Me.Caption = Me.Caption & " " & App.Title
    lblVersion.Caption = lblVersion.Caption & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = App.Title
    lblDonate.Caption = lblDonate.Caption & vbCrLf & vbCrLf & "ENS: " & ENS & vbCrLf & WalletAddress
End Sub

Private Sub lblDonate_Click()
    Shell "explorer " & "https://etherscan.io/address/" & WalletAddress
End Sub

Private Sub lblWebsite_Click()
    Shell "explorer " & lblWebsite.Caption, vbNormalFocus
End Sub
