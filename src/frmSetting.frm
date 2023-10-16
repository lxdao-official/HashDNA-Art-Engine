VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   9795
   ClientLeft      =   5010
   ClientTop       =   2055
   ClientWidth     =   10665
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   Begin VB.Frame FrameCharSettings 
      Caption         =   "Special Characters Setting"
      Height          =   2655
      Left            =   6960
      TabIndex        =   51
      Tag             =   "1"
      Top             =   6120
      Width           =   3255
      Begin VB.TextBox txtReplace 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   360
         TabIndex        =   61
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtSign 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Index           =   2
         Left            =   1920
         TabIndex        =   59
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtReplace 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   360
         TabIndex        =   58
         Text            =   "/"
         Top             =   1440
         Width           =   975
      End
      Begin VB.CheckBox chkReplace 
         Caption         =   "Replace Sign"
         Height          =   375
         Left            =   360
         TabIndex        =   57
         Tag             =   "1"
         Top             =   360
         Width           =   2655
      End
      Begin VB.TextBox txtSign 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Index           =   1
         Left            =   1920
         TabIndex        =   54
         Text            =   "_S_"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtSign 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   1920
         TabIndex        =   53
         Tag             =   "01"
         Text            =   "_C_"
         ToolTipText     =   "The sign in the file name."
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox txtReplace 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   52
         Tag             =   "01"
         Text            =   ":"
         ToolTipText     =   "Special Character"
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblSign 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         ForeColor       =   &H80000011&
         Height          =   330
         Index           =   2
         Left            =   1440
         TabIndex        =   60
         Tag             =   "1"
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label lblSign 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         ForeColor       =   &H80000011&
         Height          =   330
         Index           =   1
         Left            =   1440
         TabIndex        =   56
         Tag             =   "1"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label lblSign 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ">>"
         ForeColor       =   &H80000011&
         Height          =   330
         Index           =   0
         Left            =   1440
         TabIndex        =   55
         Tag             =   "1"
         Top             =   1035
         Width           =   495
      End
   End
   Begin VB.TextBox txtExtra 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   13
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtExtraValue 
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   14
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   9000
      TabIndex        =   27
      Tag             =   "1"
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadSetting 
      Caption         =   "Load"
      Height          =   495
      Left            =   3960
      TabIndex        =   25
      Tag             =   "1"
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame FrameImageSetting 
      Caption         =   "Image Settings"
      Height          =   5415
      Left            =   6960
      TabIndex        =   43
      Tag             =   "1"
      Top             =   360
      Width           =   3255
      Begin VB.TextBox txtDnaTryTimes 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Tag             =   "01"
         Text            =   "20000"
         ToolTipText     =   "Must be a number. After N attempts, if the unique DNA is still not obtained, the attempt is stopped."
         Top             =   4800
         Width           =   975
      End
      Begin VB.TextBox txtStaticColor 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Tag             =   "01"
         Text            =   "FFFFFF"
         ToolTipText     =   "The background color must be a 6-character(RGB) or 8-character(ARGB) hexadecimal without a pre-pended #"
         Top             =   3960
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Tag             =   "01"
         Text            =   "800"
         ToolTipText     =   "The image width must be a number."
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtHeight 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Tag             =   "01"
         Text            =   "800"
         ToolTipText     =   "The image height must be a number."
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkSmoothing 
         Caption         =   "Smoothing"
         Height          =   375
         Left            =   360
         TabIndex        =   15
         Tag             =   "1"
         Top             =   360
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.TextBox txtLightness 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   20
         Tag             =   "01"
         Text            =   "80"
         ToolTipText     =   "The image background color lightness must be a 0-100 number, , 100 is all white"
         Top             =   3120
         Width           =   975
      End
      Begin VB.CheckBox chkBackground 
         Caption         =   "Generate Background"
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Tag             =   "1"
         Top             =   2640
         Width           =   2535
      End
      Begin VB.CheckBox chkStaticColor 
         Caption         =   "Static"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   21
         Tag             =   "1"
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox chkResize 
         Caption         =   "Resize Image "
         Height          =   375
         Left            =   360
         TabIndex        =   16
         Tag             =   "1"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblDNATryTimes 
         Alignment       =   1  'Right Justify
         Caption         =   "DNA Try Times"
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Tag             =   "1"
         Top             =   4920
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   5
         X1              =   240
         X2              =   2880
         Y1              =   4560
         Y2              =   4560
      End
      Begin VB.Label lblDefault 
         Alignment       =   1  'Right Justify
         Caption         =   "Default"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   360
         TabIndex        =   44
         Tag             =   "1"
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label lblWidth 
         Alignment       =   1  'Right Justify
         Caption         =   "Width"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Tag             =   "1"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lblHeight 
         Alignment       =   1  'Right Justify
         Caption         =   "Height"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Tag             =   "1"
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label lblLightness 
         Alignment       =   1  'Right Justify
         Caption         =   "Lightness"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   360
         TabIndex        =   45
         Tag             =   "1"
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   4
         X1              =   240
         X2              =   2880
         Y1              =   840
         Y2              =   840
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   6
         X1              =   240
         X2              =   2880
         Y1              =   2520
         Y2              =   2520
      End
   End
   Begin VB.Frame FrameMetadataSetting 
      Caption         =   "Metadata Settings"
      Height          =   8415
      Left            =   360
      TabIndex        =   32
      Tag             =   "1"
      Top             =   360
      Width           =   6255
      Begin VB.CheckBox chkIgnoreNONE 
         Caption         =   "Ignore NONE"
         Height          =   375
         Left            =   4560
         TabIndex        =   31
         Tag             =   "11"
         ToolTipText     =   "The None attribute is ignored in the metadata"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.CheckBox chkWhiteSpace 
         Caption         =   "Format JSON"
         Height          =   375
         Left            =   2880
         TabIndex        =   30
         Tag             =   "11"
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton OptionNetwork 
         Caption         =   "Solana"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   29
         Tag             =   "01"
         ToolTipText     =   "Solana Network"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtExtra 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   11
         Top             =   7200
         Width           =   1935
      End
      Begin VB.TextBox txtExtraValue 
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   12
         Top             =   7200
         Width           =   1935
      End
      Begin VB.TextBox txtExtraValue 
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   10
         Tag             =   "01"
         ToolTipText     =   "Input Value"
         Top             =   6720
         Width           =   1935
      End
      Begin VB.TextBox txtExtra 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   9
         Tag             =   "01"
         ToolTipText     =   "Input Key"
         Top             =   6720
         Width           =   1935
      End
      Begin VB.TextBox txtAnimation_url 
         Height          =   375
         Left            =   1800
         TabIndex        =   8
         Tag             =   "01"
         ToolTipText     =   $"frmSetting.frx":7F6A
         Top             =   6240
         Width           =   3975
      End
      Begin VB.TextBox txtExternal_url 
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         Tag             =   "01"
         ToolTipText     =   "Display this URL in the NFT information"
         Top             =   5760
         Width           =   3975
      End
      Begin VB.TextBox txtNamePrefix 
         Height          =   375
         Left            =   1800
         TabIndex        =   0
         Tag             =   "01"
         Text            =   "Your Collection"
         ToolTipText     =   "Your collection name"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtImageBaseURL 
         Height          =   375
         Left            =   1800
         TabIndex        =   2
         Tag             =   "01"
         Text            =   "ipfs://YourImagesCID/"
         ToolTipText     =   "The URL or ipfs CID of the images folder. At the end is / "
         Top             =   2640
         Width           =   3975
      End
      Begin VB.TextBox txtDescription 
         Height          =   975
         Left            =   1800
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         Tag             =   "01"
         ToolTipText     =   $"frmSetting.frx":8002
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtSolSymbol 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Tag             =   "01"
         ToolTipText     =   "When Solana is selected, the symbol cannot be empty"
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox txtSolFee 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Tag             =   "01"
         Text            =   "500"
         ToolTipText     =   "Define how much % you want from secondary market sales, 1000 = 10%.The fee must be a number and less than 10000 "
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtSolCreatorsAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Tag             =   "01"
         ToolTipText     =   "The address of the wallet for the copyright fee collection, usually 44 characters"
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox txtSolCreatorsShare 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   5
         Tag             =   "01"
         Text            =   "100"
         ToolTipText     =   "Creators share %. Default 100"
         Top             =   4440
         Width           =   1215
      End
      Begin VB.OptionButton OptionNetwork 
         Caption         =   "Ethereum"
         Height          =   375
         Index           =   0
         Left            =   360
         TabIndex        =   28
         Tag             =   "01"
         ToolTipText     =   "Ethereum Network"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label lblExtendedinfo 
         Caption         =   "Extended information"
         Height          =   375
         Left            =   480
         TabIndex        =   49
         Tag             =   "1"
         Top             =   5280
         Width           =   4185
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   3
         X1              =   360
         X2              =   5880
         Y1              =   5040
         Y2              =   5040
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   1
         X1              =   360
         X2              =   5880
         Y1              =   3240
         Y2              =   3240
      End
      Begin VB.Label lblAnimation_url 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "animation_url"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   195
         TabIndex        =   48
         Tag             =   "1"
         Top             =   6360
         Width           =   1425
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   255
         Left            =   5880
         TabIndex        =   42
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label lblExternal_url 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "external_url"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   41
         Tag             =   "1"
         Top             =   5880
         Width           =   1425
      End
      Begin VB.Label lblNamePrefix 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "namePrefix"
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Tag             =   "1"
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label lblImageBaseURL 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "imageBaseURL"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Tag             =   "1"
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label lblDescription 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "description"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Tag             =   "1"
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label lblExtraMetadata 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "extraMetadata"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Tag             =   "1"
         Top             =   6840
         Width           =   1425
      End
      Begin VB.Label lblSysbol 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "symbol"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Tag             =   "1"
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label lblFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "fee"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   2880
         TabIndex        =   35
         Tag             =   "1"
         Top             =   4560
         Width           =   1425
      End
      Begin VB.Label lblCreatorsAdd 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "creatorsAdd."
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   315
         TabIndex        =   34
         Tag             =   "1"
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label lblCreatorsShare 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "creatorsShare"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   33
         Tag             =   "1"
         Top             =   4560
         Width           =   1425
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   2
         X1              =   360
         X2              =   5880
         Y1              =   840
         Y2              =   840
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      Height          =   495
      Left            =   360
      TabIndex        =   24
      Tag             =   "1"
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5400
      TabIndex        =   26
      Tag             =   "1"
      Top             =   9000
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Copyright 2023 LXDAO

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
Private Sub Form_Load()
    TranslateForm Me
    cmdLoadSetting_Click
End Sub

Private Sub cmdClear_Click()
        Dim i As Integer
        Dim o As Object
        For Each o In Me.Controls
            If typeName(o) = "TextBox" Then o.BackColor = &H80000005
        Next
   If MsgBox(Language.Item("Tips35"), vbQuestion + vbYesNo) = vbYes Then
        OptionNetwork(0).Value = True
        OptionNetwork(1).Value = False
        chkWhiteSpace.Value = Checked
        chkIgnoreNONE.Value = Checked
        txtNamePrefix.Text = "Your Collection"
        txtDescription.Text = ""
        txtImageBaseURL.Text = "ipfs://YourImagesCID/"
        txtSolSymbol.Text = ""
        txtSolCreatorsAddress.Text = ""
        txtSolCreatorsShare.Text = 100
        txtSolFee.Text = 500
        txtExternal_url.Text = ""
        txtAnimation_url.Text = ""
        For i = 0 To txtExtra.UBound
            txtExtra(i).Text = ""
            txtExtraValue(i).Text = ""
        Next i
        chkSmoothing.Value = Checked
        chkResize.Value = Unchecked
        txtWidth.Text = 800
        txtHeight.Text = 800
        chkBackground.Value = Unchecked
        txtLightness.Text = 80
        chkStaticColor.Value = Unchecked
        txtStaticColor.Text = "FFFFFF"
        txtDnaTryTimes.Text = 20000
        chkReplace.Value = Checked
        txtReplace(0) = ":"
        txtReplace(1) = "/"
        txtReplace(2) = ""
        txtSign(0) = "_COLONS_"
        txtSign(1) = "_SLASH_"
        txtSign(2) = ""
   End If
End Sub

Private Sub cmdLoadSetting_Click()
    If Dir(layersDir & "\Setting.json") = "" Then Exit Sub
    Dim settingJB As JsonBag, fn As Integer, i As Integer
    Dim o As Object, sKey As String, ctrlIndex As Integer
    Set settingJB = New JsonBag
    settingJB.Whitespace = True
    fn = FreeFile
    On Error Resume Next
    Open layersDir & "\Setting.json" For Input As #fn
    settingJB.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
    Close #fn
    Err.Clear
    If settingJB Is Nothing Then Exit Sub
    For Each o In Me.Controls
        sKey = o.Name
        On Error Resume Next
        ctrlIndex = o.Index
        If Err.Number = 0 Then sKey = sKey & "-" & ctrlIndex
        Err.Clear
        If settingJB.Exists(sKey & ".Value") Then o.Value = settingJB.Item(sKey & ".Value")
        If settingJB.Exists(sKey & ".Text") Then o.Text = settingJB.Item(sKey & ".Text")
    Next
End Sub

Private Sub cmdSave_Click()
    If saveSetting = True Then
        MsgBox Language.Item("Tips33"), vbInformation
    End If
End Sub

Private Function saveSetting() As Boolean
    Dim settingJB As JsonBag, fn As Integer, i As Integer
    Dim o As Object, errCount As Long
    Dim sKey As String
    Dim ctrlIndex As Integer
    errCount = 0
    saveSetting = False
    For Each o In Me.Controls
        If typeName(o) = "TextBox" And o <> txtDescription Then
            o.BackColor = &H80000005
            o.Text = Trim(o.Text)
        End If
    Next
    If Right(txtImageBaseURL.Text, 1) <> "/" Then
        txtImageBaseURL.BackColor = &HC0FFFF
        errCount = errCount + 1
    End If
    If OptionNetwork(1) = True Then
        If txtSolSymbol = "" Then
            txtSolSymbol.BackColor = &HC0FFFF
            errCount = errCount + 1
        End If
        If txtSolFee <> "" Then
            If (Not IsNumeric(txtSolFee.Text)) Or (Val(txtSolFee) > 10000) Then
                txtSolFee.BackColor = &HC0FFFF
                errCount = errCount + 1
            Else
                If txtSolCreatorsAddress.Text = "" Then
                    txtSolCreatorsAddress.BackColor = &HC0FFFF
                    errCount = errCount + 1
                ElseIf Len(txtSolCreatorsAddress.Text) <> 44 Then
                    txtSolCreatorsAddress.BackColor = &H80000018
                End If
                If txtSolCreatorsShare = "" Or Not IsNumeric(txtSolCreatorsShare.Text) Then
                    txtSolCreatorsShare.BackColor = &HC0FFFF
                    errCount = errCount + 1
                End If
            End If
        End If
    End If
    If chkResize.Value = Checked Then
        If Not IsNumeric(txtWidth) Then
            txtWidth.BackColor = &HC0FFFF
            errCount = errCount + 1
        End If
        If Not IsNumeric(txtHeight) Then
            txtHeight.BackColor = &HC0FFFF
            errCount = errCount + 1
        End If
    End If
    If chkBackground.Value = Checked Then
        If (Not IsNumeric(txtLightness)) Or Val(txtLightness) > 100 Or Val(txtLightness) < 0 Then
            txtLightness.BackColor = &HC0FFFF
            errCount = errCount + 1
        End If
        If chkStaticColor.Value = Checked Then
            If Len(txtStaticColor) <> 6 And Len(txtStaticColor) <> 8 Then
                txtStaticColor.BackColor = &HC0FFFF
                errCount = errCount + 1
            End If
        End If
    End If
    If Not IsNumeric(txtDnaTryTimes) Then
        txtDnaTryTimes.BackColor = &HC0FFFF
        errCount = errCount + 1
    End If
    If errCount > 0 Then
        MsgBox Language.Item("Tips34"), vbCritical
        Exit Function
    End If
    Set settingJB = New JsonBag
    settingJB.Whitespace = True
    With settingJB
        .Clear
        .IsArray = False
        For Each o In Me.Controls
            sKey = o.Name
            On Error Resume Next
            ctrlIndex = o.Index
            If Err.Number = 0 Then sKey = sKey & "-" & ctrlIndex
            Err.Clear
            
            If typeName(o) = "TextBox" Then
                .Item(sKey & ".Text") = o.Text
            ElseIf typeName(o) = "OptionButton" Or typeName(o) = "CheckBox" Then
                .Item(sKey & ".Value") = o.Value
            End If
        Next
    End With
    fn = FreeFile
    Open layersDir & "\Setting.json" For Output As #fn
    Print #fn, settingJB.JSON
    Close #fn
    saveSetting = True
End Function

Private Sub cmdBack_Click()
    If saveSetting = True Then Me.Hide
End Sub

Private Sub OptionNetwork_Click(Index As Integer)
    If Index = 0 Then
        txtSolSymbol.BackColor = &H80000005
        txtSolFee.BackColor = &H80000005
        txtSolCreatorsAddress.BackColor = &H80000005
        txtSolCreatorsShare.BackColor = &H80000005
        txtSolSymbol.Enabled = False
        txtSolFee.Enabled = False
        txtSolCreatorsAddress.Enabled = False
        txtSolCreatorsShare.Enabled = False
        lblSysbol.ForeColor = &H80000011
        lblCreatorsAdd.ForeColor = &H80000011
        lblCreatorsShare.ForeColor = &H80000011
        lblFee.ForeColor = &H80000011
    ElseIf Index = 1 Then
        txtSolSymbol.Enabled = True
        txtSolFee.Enabled = True
        txtSolCreatorsAddress.Enabled = True
        txtSolCreatorsShare.Enabled = True
        lblSysbol.ForeColor = &H80000012
        lblCreatorsAdd.ForeColor = &H80000012
        lblCreatorsShare.ForeColor = &H80000012
        lblFee.ForeColor = &H80000012
    End If
End Sub

Private Sub chkResize_Click()
    If chkResize.Value = Unchecked Then
        lblWidth.ForeColor = &H80000011
        lblHeight.ForeColor = &H80000011
        txtWidth.Enabled = False
        txtHeight.Enabled = False
    ElseIf chkResize.Value = Checked Then
        lblWidth.ForeColor = &H80000012
        lblHeight.ForeColor = &H80000012
        txtWidth.Enabled = True
        txtHeight.Enabled = True
    End If
End Sub

Private Sub chkBackground_Click()
    If chkBackground.Value = Unchecked Then
        lblLightness.ForeColor = &H80000011
        lblDefault.ForeColor = &H80000011
        txtLightness.Enabled = False
        txtStaticColor.Enabled = False
        chkStaticColor.Enabled = False
    ElseIf chkBackground.Value = Checked Then
        lblLightness.ForeColor = &H80000012
        lblDefault.ForeColor = &H80000012
        txtLightness.Enabled = True
        txtStaticColor.Enabled = True
        chkStaticColor.Enabled = True
    End If
End Sub

Private Sub chkReplace_Click()
    Dim i As Integer
    If chkReplace.Value = Unchecked Then
        For i = 0 To txtReplace.UBound
                lblSign(i).ForeColor = &H80000011
                txtReplace(i).Enabled = False
                txtSign(i).Enabled = False
        Next i
    ElseIf chkReplace.Value = Checked Then
        For i = 0 To txtReplace.UBound
            lblSign(i).ForeColor = &H80000012
            txtReplace(i).Enabled = True
            txtSign(i).Enabled = True
        Next i
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If saveSetting = True Then
        Cancel = True
        Me.Hide
    Else
        Cancel = 2
    End If
End Sub




