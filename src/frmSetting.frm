VERSION 5.00
Begin VB.Form frmSetting 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Setting"
   ClientHeight    =   9795
   ClientLeft      =   3570
   ClientTop       =   3450
   ClientWidth     =   10665
   Icon            =   "frmSetting.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9795
   ScaleWidth      =   10665
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtExtra 
      Height          =   375
      Index           =   2
      Left            =   2160
      TabIndex        =   16
      Top             =   8040
      Width           =   1935
   End
   Begin VB.TextBox txtExtraValue 
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   17
      Top             =   8040
      Width           =   1935
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "Back"
      Height          =   495
      Left            =   9000
      TabIndex        =   30
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdLoadSetting 
      Caption         =   "Load"
      Height          =   495
      Left            =   3960
      TabIndex        =   27
      Top             =   9000
      Width           =   1215
   End
   Begin VB.Frame FrameStep3 
      Caption         =   "Image Settings"
      Height          =   8415
      Left            =   6960
      TabIndex        =   42
      Top             =   360
      Width           =   3255
      Begin VB.TextBox txtDnaTryTimes 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1680
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   20
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
         TabIndex        =   21
         Text            =   "800"
         ToolTipText     =   "The image height must be a number."
         Top             =   1920
         Width           =   975
      End
      Begin VB.CheckBox chkSmoothing 
         Caption         =   "Smoothing"
         Height          =   375
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.TextBox txtLightness 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   23
         Text            =   "80"
         ToolTipText     =   "The image background color lightness must be a 0-100 number."
         Top             =   3120
         Width           =   975
      End
      Begin VB.CheckBox chkBackground 
         Caption         =   "Generate Background"
         Height          =   375
         Left            =   360
         TabIndex        =   22
         Top             =   2640
         Width           =   2295
      End
      Begin VB.CheckBox chkStaticColor 
         Caption         =   "Static"
         Enabled         =   0   'False
         Height          =   255
         Left            =   720
         TabIndex        =   24
         Top             =   3600
         Width           =   1455
      End
      Begin VB.CheckBox chkResize 
         Caption         =   "Resize Image "
         Height          =   375
         Left            =   360
         TabIndex        =   19
         Top             =   960
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000000&
         Index           =   0
         X1              =   240
         X2              =   2880
         Y1              =   5400
         Y2              =   5400
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "DNA Try Times"
         Height          =   255
         Left            =   120
         TabIndex        =   50
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
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Default"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   360
         TabIndex        =   43
         Top             =   4080
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Reserved parameters zone"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   360
         TabIndex        =   49
         Top             =   5640
         Width           =   2655
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Width"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Height"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "Lightness"
         ForeColor       =   &H80000011&
         Height          =   255
         Left            =   360
         TabIndex        =   44
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
   Begin VB.Frame FrameStep2 
      Caption         =   "Metadata Settings"
      Height          =   8415
      Left            =   360
      TabIndex        =   31
      Top             =   360
      Width           =   6255
      Begin VB.CheckBox chkIgnoreNONE 
         Caption         =   "Ignore NONE"
         Height          =   375
         Left            =   4560
         TabIndex        =   51
         Top             =   360
         Value           =   1  'Checked
         Width           =   1455
      End
      Begin VB.CheckBox chkWhiteSpace 
         Caption         =   "AddWhitespace"
         Height          =   375
         Left            =   2880
         TabIndex        =   2
         Top             =   360
         Value           =   1  'Checked
         Width           =   1575
      End
      Begin VB.OptionButton OptionNetwork 
         Caption         =   "Solana"
         Height          =   375
         Index           =   1
         Left            =   1680
         TabIndex        =   1
         ToolTipText     =   "Solana Network"
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtExtra 
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   14
         Top             =   7200
         Width           =   1935
      End
      Begin VB.TextBox txtExtraValue 
         Height          =   375
         Index           =   1
         Left            =   3840
         TabIndex        =   15
         Top             =   7200
         Width           =   1935
      End
      Begin VB.TextBox txtExtraValue 
         Height          =   375
         Index           =   0
         Left            =   3840
         TabIndex        =   13
         Top             =   6720
         Width           =   1935
      End
      Begin VB.TextBox txtExtra 
         Height          =   375
         Index           =   0
         Left            =   1800
         TabIndex        =   12
         Top             =   6720
         Width           =   1935
      End
      Begin VB.TextBox txtAnimation_url 
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   6240
         Width           =   3975
      End
      Begin VB.TextBox txtExternal_url 
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   5760
         Width           =   3975
      End
      Begin VB.TextBox txtNamePrefix 
         Height          =   375
         Left            =   1800
         TabIndex        =   3
         Text            =   "Your Collection"
         ToolTipText     =   "Your collection name"
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox txtImageBaseURL 
         Height          =   375
         Left            =   1800
         TabIndex        =   5
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
         TabIndex        =   4
         ToolTipText     =   $"frmSetting.frx":7E6A
         Top             =   1560
         Width           =   3975
      End
      Begin VB.TextBox txtSolSymbol 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         ToolTipText     =   "When Solana is selected, the symbol cannot be empty."
         Top             =   3480
         Width           =   3975
      End
      Begin VB.TextBox txtSolFee 
         Enabled         =   0   'False
         Height          =   375
         Left            =   4440
         TabIndex        =   9
         Text            =   "500"
         ToolTipText     =   "Define how much % you want from secondary market sales, 1000 = 10%.The fee must be a number and less than 10000. "
         Top             =   4440
         Width           =   1335
      End
      Begin VB.TextBox txtSolCreatorsAddress 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   7
         ToolTipText     =   "The wallet address is generally 44 characters."
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox txtSolCreatorsShare 
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         TabIndex        =   8
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
         TabIndex        =   0
         ToolTipText     =   "Ethereum Network"
         Top             =   360
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         Caption         =   "Extended information"
         Height          =   375
         Left            =   360
         TabIndex        =   48
         Top             =   5280
         Width           =   2025
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "animation_url"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   195
         TabIndex        =   47
         ToolTipText     =   $"frmSetting.frx":7E8B
         Top             =   6360
         Width           =   1425
      End
      Begin VB.Label Label77 
         BackStyle       =   0  'Transparent
         Caption         =   "*"
         Height          =   255
         Left            =   5880
         TabIndex        =   41
         Top             =   2760
         Width           =   255
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "external_url"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   40
         Top             =   5880
         Width           =   1425
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "namePrefix"
         Height          =   375
         Left            =   240
         TabIndex        =   39
         Top             =   1200
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "imageBaseURL"
         Height          =   375
         Left            =   240
         TabIndex        =   38
         Top             =   2760
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "description"
         Height          =   375
         Left            =   240
         TabIndex        =   37
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "extraMetadata"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   36
         Top             =   6840
         Width           =   1425
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "symbol"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   35
         Top             =   3600
         Width           =   1425
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "fee"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   2880
         TabIndex        =   34
         Top             =   4560
         Width           =   1425
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "creatorsAdd."
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   315
         TabIndex        =   33
         Top             =   4080
         Width           =   1425
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "creatorsShare"
         ForeColor       =   &H80000011&
         Height          =   375
         Left            =   240
         TabIndex        =   32
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
      TabIndex        =   29
      Top             =   9000
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   495
      Left            =   5400
      TabIndex        =   28
      Top             =   9000
      Width           =   1215
   End
End
Attribute VB_Name = "frmSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBack_Click()
    If saveSetting = True Then Me.Hide
End Sub

Private Sub cmdClear_Click()
        Dim i As Integer
        Dim o As Object
        For Each o In Me.Controls
            If typeName(o) = "TextBox" Then o.BackColor = &H80000005
        Next
   If MsgBox("Are you sure to clear?", vbQuestion + vbYesNo) = vbYes Then
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
        txtExtra(0).Text = ""
        txtExtra(1).Text = ""
        txtExtra(2).Text = ""
        txtExtraValue(0).Text = ""
        txtExtraValue(1).Text = ""
        txtExtraValue(2).Text = ""
        chkSmoothing.Value = Checked
        chkResize.Value = Unchecked
        txtWidth.Text = 800
        txtHeight.Text = 800
        chkBackground.Value = Unchecked
        txtLightness.Text = 80
        chkStaticColor.Value = Unchecked
        txtStaticColor.Text = "FFFFFF"
        txtDnaTryTimes.Text = 20000
   End If
End Sub

Private Sub cmdLoadSetting_Click()
    If Dir(layersDir & "\Setting.json") = "" Then Exit Sub
    Dim settingJB As JsonBag, fn As Integer, i As Integer
    Dim o As Object
    Set settingJB = New JsonBag
    settingJB.Whitespace = True
    fn = FreeFile
    Open layersDir & "\Setting.json" For Input As #fn
    settingJB.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
    Close #fn
    With settingJB
        For Each o In Me.Controls
            If typeName(o) = "TextBox" Then
                If VarType(CallByName(Me, o.Name, VbGet)) = vbObject Then
                    If .Exists(o.Name & "-" & o.Index) Then o.Text = .Item(o.Name & "-" & o.Index)
                Else
                   If .Exists(o.Name) Then o.Text = .Item(o.Name)
                End If
            ElseIf typeName(o) = "OptionButton" Or typeName(o) = "CheckBox" Then
                If VarType(CallByName(Me, o.Name, VbGet)) = vbObject Then
                    If .Exists(o.Name & "-" & o.Index) Then o.Value = .Item(o.Name & "-" & o.Index)
                Else
                    If .Exists(o.Name) Then o.Value = .Item(o.Name)
                End If
            End If
        Next
    End With
End Sub

Private Sub Form_Load()
    cmdLoadSetting_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If saveSetting = True Then
        Cancel = True
        Me.Hide
    Else
        Cancel = 2
    End If
End Sub

Private Sub cmdSave_Click()
    If saveSetting = True Then
        MsgBox "The setting has saved.", vbInformation
    End If
End Sub

Private Function saveSetting() As Boolean
    Dim settingJB As JsonBag, fn As Integer, i As Integer
    Dim o As Object, errCount As Long
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
        MsgBox "Pleae fix errors.", vbCritical
        Exit Function
    End If
    Set settingJB = New JsonBag
    settingJB.Whitespace = True
    With settingJB
        .Clear
        .IsArray = False
        For Each o In Me.Controls
            If typeName(o) = "TextBox" Then
                If VarType(CallByName(Me, o.Name, VbGet)) = vbObject Then .Item(o.Name & "-" & o.Index) = o.Text Else .Item(o.Name) = o.Text
            ElseIf typeName(o) = "OptionButton" Or typeName(o) = "CheckBox" Then
                .Item(o.Name) = o.Value
                If VarType(CallByName(Me, o.Name, VbGet)) = vbObject Then .Item(o.Name & "-" & o.Index) = o.Value Else .Item(o.Name) = o.Value
            End If
        Next
    End With
    fn = FreeFile
    Open layersDir & "\Setting.json" For Output As #fn
    Print #fn, settingJB.JSON
    Close #fn
    saveSetting = True
End Function

Private Sub OptionNetwork_Click(Index As Integer)
    If Index = 0 Then
        txtSolSymbol.Enabled = False
        txtSolFee.Enabled = False
        txtSolCreatorsAddress.Enabled = False
        txtSolCreatorsShare.Enabled = False
        Label7.ForeColor = &H80000011
        Label8.ForeColor = &H80000011
        Label10.ForeColor = &H80000011
        Label11.ForeColor = &H80000011
    ElseIf Index = 1 Then
        txtSolSymbol.Enabled = True
        txtSolFee.Enabled = True
        txtSolCreatorsAddress.Enabled = True
        txtSolCreatorsShare.Enabled = True
        Label7.ForeColor = &H80000012
        Label8.ForeColor = &H80000012
        Label10.ForeColor = &H80000012
        Label11.ForeColor = &H80000012
    End If
End Sub

Private Sub chkBackground_Click()
    If chkBackground.Value = Unchecked Then
        Label15.ForeColor = &H80000011
        Label16.ForeColor = &H80000011
        txtLightness.Enabled = False
        txtStaticColor.Enabled = False
        chkStaticColor.Enabled = False
    ElseIf chkBackground.Value = Checked Then
        Label15.ForeColor = &H80000012
        Label16.ForeColor = &H80000012
        txtLightness.Enabled = True
        txtStaticColor.Enabled = True
        chkStaticColor.Enabled = True
    End If
End Sub

Private Sub chkResize_Click()
    If chkResize.Value = Unchecked Then
        Label13.ForeColor = &H80000011
        Label14.ForeColor = &H80000011
        txtWidth.Enabled = False
        txtHeight.Enabled = False
    ElseIf chkResize.Value = Checked Then
        Label13.ForeColor = &H80000012
        Label14.ForeColor = &H80000012
        txtWidth.Enabled = True
        txtHeight.Enabled = True
    End If
End Sub

