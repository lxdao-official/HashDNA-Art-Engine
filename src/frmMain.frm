VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "HashDNA Art Eengine"
   ClientHeight    =   8310
   ClientLeft      =   2940
   ClientTop       =   2835
   ClientWidth     =   13905
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8310
   ScaleWidth      =   13905
   Begin VB.CommandButton cmdTop 
      Caption         =   "--"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7440
      TabIndex        =   38
      Tag             =   "01"
      ToolTipText     =   "Move to the top"
      Top             =   840
      Width           =   735
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   7920
      Top             =   7080
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   8640
      TabIndex        =   34
      Top             =   6000
      Width           =   2295
      Begin VB.CommandButton cmdTools 
         Caption         =   "Tools"
         Height          =   495
         Left            =   600
         TabIndex        =   35
         Tag             =   "1"
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   11400
      TabIndex        =   33
      Tag             =   "1"
      Top             =   7080
      Width           =   2175
   End
   Begin VB.CommandButton cmdAbout 
      Caption         =   "About"
      Height          =   495
      Left            =   11400
      TabIndex        =   32
      Tag             =   "1"
      Top             =   6480
      Width           =   2175
   End
   Begin VB.ComboBox cboLanguages 
      Height          =   300
      ItemData        =   "frmMain.frx":7F6A
      Left            =   11400
      List            =   "frmMain.frx":7F6C
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Tag             =   "01"
      ToolTipText     =   "Select the interface language"
      Top             =   6000
      Width           =   2175
   End
   Begin VB.PictureBox picTips 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      FontTransparent =   0   'False
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   13845
      TabIndex        =   23
      Top             =   7935
      Width           =   13905
   End
   Begin VB.Frame FrameStep3 
      Caption         =   "Step 3: Select Images"
      Height          =   1575
      Left            =   4200
      TabIndex        =   21
      Tag             =   "1"
      Top             =   6000
      Width           =   3495
      Begin VB.TextBox txtStartNumber 
         Alignment       =   2  'Center
         Height          =   435
         Left            =   360
         TabIndex        =   9
         Tag             =   "01"
         Text            =   "1"
         ToolTipText     =   "Picture Start number"
         Top             =   500
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   495
         Left            =   1920
         TabIndex        =   11
         Tag             =   "11"
         ToolTipText     =   "After deleting the selected images, click this button to start renumbering the images and sync the JSON files"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkShuffle 
         Caption         =   "Shuffle"
         Height          =   375
         Left            =   480
         TabIndex        =   10
         Tag             =   "11"
         ToolTipText     =   "Shuffle the order of the pictures"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2415
      End
   End
   Begin VB.Frame FrameStep2 
      Caption         =   "Step 2: Generate images"
      Height          =   1575
      Left            =   240
      TabIndex        =   19
      Tag             =   "1"
      Top             =   6000
      Width           =   3495
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H0080FF80&
         Caption         =   "Start"
         Height          =   495
         Left            =   1920
         TabIndex        =   7
         Tag             =   "1"
         Top             =   480
         Width           =   1215
      End
      Begin VB.CheckBox chkClean 
         Caption         =   "Clean Folder"
         Height          =   375
         Left            =   360
         TabIndex        =   8
         Tag             =   "11"
         ToolTipText     =   "Clean IMAGES folder and JSON folder"
         Top             =   1080
         Value           =   1  'Checked
         Width           =   2775
      End
      Begin VB.CommandButton cmdSetting 
         Caption         =   "Setting"
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Tag             =   "11"
         ToolTipText     =   "Set metadata parameters and image parameters"
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Frame FrameStep1 
      Caption         =   "Step 1: Layer Configurations"
      Height          =   5415
      Left            =   240
      TabIndex        =   13
      Tag             =   "1"
      Top             =   240
      Width           =   13335
      Begin VB.CommandButton cmdDelWeight 
         Caption         =   "#10"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   -1  'True
         EndProperty
         Height          =   375
         Left            =   5880
         TabIndex        =   39
         Tag             =   "01"
         ToolTipText     =   "Delete the elements weight of the selected layer"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.CommandButton cmdBottom 
         Caption         =   "--"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   37
         Tag             =   "01"
         ToolTipText     =   "Move to the bottom"
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton cmdAddWeight 
         Caption         =   "#10"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4320
         TabIndex        =   36
         Tag             =   "01"
         ToolTipText     =   "Set the elements weight of the selected layer to 10"
         Top             =   4200
         Width           =   1215
      End
      Begin VB.PictureBox picPreview 
         FillColor       =   &H008080FF&
         Height          =   4575
         Left            =   8400
         ScaleHeight     =   4515
         ScaleWidth      =   4515
         TabIndex        =   30
         Top             =   600
         Width           =   4575
      End
      Begin VB.TextBox txtTypeDisplayName 
         Height          =   375
         Left            =   1200
         TabIndex        =   1
         Tag             =   "01"
         ToolTipText     =   "Modify the display name of type"
         Top             =   4200
         Width           =   1935
      End
      Begin VB.CommandButton cmdDownLayer 
         Caption         =   "Down"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   29
         Tag             =   "01"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdUpLayer 
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   28
         Tag             =   "01"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdDownType 
         Caption         =   "Down"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   27
         Tag             =   "01"
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton cmdUpType 
         Caption         =   "Up"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Tag             =   "01"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton cmdDelType 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   25
         Tag             =   "01"
         ToolTipText     =   "Delete one type"
         Top             =   3240
         Width           =   735
      End
      Begin VB.CommandButton cmdCopyType 
         Caption         =   "Copy"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Tag             =   "01"
         ToolTipText     =   "Copy the selected type"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdReload 
         Caption         =   "Reload"
         Height          =   495
         Left            =   1920
         TabIndex        =   4
         Tag             =   "11"
         ToolTipText     =   "Load layer configurations by folder structure"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdBypassDNA 
         Caption         =   "Pass"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   20
         Tag             =   "01"
         ToolTipText     =   "bypass DNA"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdImport 
         Caption         =   "Import"
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Tag             =   "11"
         ToolTipText     =   "Import configurations from Config.json file"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.ListBox ListLayer 
         Height          =   3480
         ItemData        =   "frmMain.frx":7F6E
         Left            =   4320
         List            =   "frmMain.frx":7F70
         TabIndex        =   16
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdDelLayer 
         Caption         =   "Del"
         BeginProperty Font 
            Name            =   "Consolas"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7200
         TabIndex        =   15
         Tag             =   "01"
         ToolTipText     =   "Delete one layer"
         Top             =   3240
         Width           =   735
      End
      Begin VB.ListBox ListType 
         Height          =   3480
         ItemData        =   "frmMain.frx":7F72
         Left            =   360
         List            =   "frmMain.frx":7F74
         TabIndex        =   14
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton cmdReloadLayer 
         Caption         =   "Reload"
         Height          =   495
         Left            =   4320
         TabIndex        =   2
         Tag             =   "11"
         ToolTipText     =   "Load layer name by folder structure"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.CommandButton cmdSaveOrder 
         Caption         =   "Save"
         Height          =   495
         Left            =   5880
         TabIndex        =   3
         Tag             =   "11"
         ToolTipText     =   "Save the current configuration to Config.json file"
         Top             =   4680
         Width           =   1215
      End
      Begin VB.TextBox txtNumType 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   360
         TabIndex        =   0
         Tag             =   "01"
         Text            =   "0"
         ToolTipText     =   "Modify the number of type"
         Top             =   4200
         Width           =   735
      End
      Begin VB.Label lblPicture 
         Caption         =   "Preview Zone"
         Height          =   255
         Left            =   8400
         TabIndex        =   31
         Tag             =   "1"
         Top             =   360
         Width           =   4755
      End
      Begin VB.Label lblType 
         BackStyle       =   0  'Transparent
         Caption         =   "Type"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Tag             =   "1"
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label lblTotalEditions 
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   405
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
      Begin VB.Label lblTopLayer 
         BackStyle       =   0  'Transparent
         Caption         =   "Top Layer"
         Height          =   255
         Left            =   4320
         TabIndex        =   17
         Tag             =   "1"
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "frmMain"
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


'********************************************************************
'
'                   HashDNA Art Engine v2.1.0
'                          2022-12-11
'
'   This software is designed to help artists freely and easily
'   generate 10K images without programming knowledge. It references
'   the design of the HashLips art engine, and the example material
'   comes from cryptopunksnotdead, thanks!
'
'   References:
'   http://hashdna.art
'   https://lxdao.io
'
'                                                     by LXDAO
'
'********************************************************************

Option Explicit
Const oneArtFolderName As String = "1of1" 'lowercase
Const signSkip As String = "_SKIP_"
Const signOnly As String = "_ONLY_"
Const rarityDelimiter As String = "#"
Const nameDelimiter As String = "@"
Dim config As JsonBag
Dim layerConfigurations() As layerConfig 'All layers configuration information
Dim layers() As layer 'All layers infomation in a type
Dim elements() As element 'All elements information in a layer
Dim DNA As Collection '
Dim newDNA() As Long  'Each element number of a DNA
Dim totalEditions As Long

'after each DoEvents line .If IsUnloading Then Exit Sub  '(or "Exit Function", etc)
Private IsUnloading As Boolean

Private Enum Execution_State
    ES_SYSTEM_REQUIRED = &H1
    ES_DISPLAY_REQUIRED = &H2
    ES_USER_PRESENT = &H4
    ES_CONTINUOUS = &H80000000
End Enum
Private Declare Sub SetThreadExecutionState Lib "kernel32" (ByVal esFlags As Long)
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long



Private Sub Form_Load()
    Me.Caption = App.Title & " V" & App.Major & "." & App.Minor & "." & App.Revision
    picPreview.ScaleMode = vbPixels
    InitGDIPlus
    BuildSetup          'public sub in Public.bas
    GetLanguagesList    'public sub in Public.bas
    If checkLayers = True Then getLayerConfigurations
End Sub
'Select Language
Private Sub cboLanguages_Click()
    Dim frm As Form
    Set Language = New JsonBag
    With Language
        .Whitespace = True
        .WhitespaceIndent = 2
        .DecimalMode = False
        .JSON = LoadResString(cboLanguages.ItemData(cboLanguages.ListIndex))
    End With
    For Each frm In Forms
        TranslateForm frm
    Next
    showTips Language.Item("Tips4")
End Sub

Private Sub cmdAbout_Click()
    frmAbout.Show 1
End Sub

Private Sub cmdTools_Click()
 frmTools.Show 1
End Sub

Private Sub cmdSetting_Click()
    frmSetting.Show 1
End Sub

'**********************************************************
'
'               Setting layer Configurations
'
'**********************************************************
Private Function checkLayers() As Boolean
    layersDir = basePath & "\layers"
    If Dir(layersDir, vbDirectory) = "" Then
        MkDir layersDir
        MkDir layersDir & "\1of1"
        MkDir layersDir & "\YourNFT"
        showTips Language.Item("Tips1")
        Shell "explorer " & layersDir, vbNormalFocus
        checkLayers = False
        Exit Function
    End If
    If Dir(layersDir & "\1of1", vbDirectory) = "" Then MkDir layersDir & "\1of1"
    checkLayers = True
End Function

'Get all combined layer configuration from layers folder structure or config.txt file, including layer order, 1/1 special folders, etc.
Private Sub getLayerConfigurations()
    If Dir(layersDir & "\Config.json") <> "" Then
        importConfig
    Else
        reloadDir
    End If
End Sub

'Get all combined layer configuration from Config.json file, including layer order, 1/1 special folders, etc.
Private Sub importConfig()
    Dim configFile As String
    Dim foldername As String
    Dim fn As Integer
    Dim i As Long, j As Long, k As Long
    configFile = layersDir & "\Config.json"
    
    If Dir(configFile) = "" Then
        reloadDir
        Exit Sub
    End If
    
    k = 0
    ListType.Clear
    
    On Error GoTo ErrorHandler
    Set config = New JsonBag
    config.Whitespace = True
    fn = FreeFile
    Open configFile For Input As #fn
    config.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
    Close #fn
    ReDim layerConfigurations(config.Count - 1)
    For i = 1 To config.Count
        With config.Item(i)
            layerConfigurations(k).typeName = .Item("typeName")
            layerConfigurations(k).displayName = .Item("displayName")
            layerConfigurations(k).typeSize = .Item("typeSize")
            layerConfigurations(k).layersSize = .Item("layersSize")
            If layerConfigurations(k).layersSize > 0 Then
                ReDim layerConfigurations(k).layersOrder(.Item("layersOrder").Count - 1)
                For j = 1 To .Item("layersOrder").Count
                    layerConfigurations(k).layersOrder(j - 1) = .Item("layersOrder")(j)
                Next j
                'List configurations on form.
                ListType.AddItem layerConfigurations(k).typeSize & vbTab & layerConfigurations(k).displayName
                k = k + 1
            End If
        End With
    Next i
    
    'Close the error trap.
    On Error GoTo 0
    'catch error
        If k = 0 Then GoTo ErrorHandler
    
    'Select the first type and the first layer
    If ListType.ListCount <> 0 Then ListType.ListIndex = 0
    If ListLayer.ListCount <> 0 Then
        ListLayer.ListIndex = 0
        'Delay previewing picture
        Timer1.Enabled = True
    End If
    Call getTotalEditions
    showTips Language.Item("Tips4")
    Exit Sub
ErrorHandler:
    Close
    showTips Language.Item("Tips5")
    Shell "explorer " & configFile, vbNormalFocus
End Sub

'Get all combined layer configuration from layers folder structure, 1/1 special folders, etc.
Private Sub reloadDir()
    Dim foldername As String
    Dim k As Long
    k = 0
    ListType.Clear
    
    foldername = Dir(layersDir & "\", vbDirectory)
    Do While foldername <> ""
        If foldername <> "." And foldername <> ".." And (GetAttr(layersDir & "\" & foldername) And vbDirectory) = vbDirectory Then
            ReDim Preserve layerConfigurations(k)
            layerConfigurations(k).typeName = foldername
            layerConfigurations(k).displayName = foldername
            layerConfigurations(k).typeSize = 0
            k = k + 1
        End If
        foldername = Dir()
    Loop
    If k = 0 Then
        showTips Language.Item("Tips2")
        Shell "explorer " & layersDir, vbNormalFocus
        Exit Sub
    End If
    
    'Get layers order information:
    For k = 0 To UBound(layerConfigurations)
        foldername = layerConfigurations(k).typeName
        'layer folders or 1/1 files
        If LCase(foldername) = oneArtFolderName Then layerConfigurations(k).layersOrder = getSpecialFiles(layersDir & "\" & foldername) _
        Else layerConfigurations(k).layersOrder = getLayersOrder(layersDir & "\" & foldername)
        'Is the type folder empty?
        If (CStr(Join(layerConfigurations(k).layersOrder, ""))) = "" Then
            showTips foldername & " " & Language.Item("Tips3")
            layerConfigurations(k).layersSize = 0
        Else
            layerConfigurations(k).layersSize = UBound(layerConfigurations(k).layersOrder) + 1
        End If
        'the number of 1/1 editions = the number of 1/1 files
        If LCase(foldername) = oneArtFolderName Then layerConfigurations(k).typeSize = layerConfigurations(k).layersSize
        'List configurations on form.
        ListType.AddItem layerConfigurations(k).typeSize & vbTab & layerConfigurations(k).displayName
    Next k

    'Select the first type and the first layer
    If ListType.ListCount <> 0 Then ListType.ListIndex = 0
    If ListLayer.ListCount <> 0 Then
        ListLayer.ListIndex = 0
        'Delay previewing picture
        Timer1.Enabled = True
    End If
    Call getTotalEditions
    showTips Language.Item("Tips4")
End Sub

'Get 1/1 files.
Private Function getSpecialFiles(ByVal typeDir As String) As String()
    Dim k As Long
    Dim specialName As String
    Dim specialFiles() As String
    k = 0
    specialName = Dir(typeDir & "\")
    Do While specialName <> ""
        ReDim Preserve specialFiles(k)
        specialFiles(k) = specialName
        k = k + 1
        specialName = Dir()
    Loop
    getSpecialFiles = specialFiles
End Function

'Get the layers order from directory structure.
Private Function getLayersOrder(ByVal typeDir As String) As String()
    Dim foldername As String
    Dim layersOrder() As String
    Dim fn As Integer
    Dim k As Long
    k = 0
    foldername = Dir(typeDir & "\", vbDirectory)
    Do While foldername <> ""
        If foldername <> "." And foldername <> ".." And (GetAttr(typeDir & "\" & foldername) And vbDirectory) = vbDirectory Then
            ReDim Preserve layersOrder(k)
            layersOrder(k) = foldername
            k = k + 1
        End If
        foldername = Dir()
    Loop
    getLayersOrder = layersOrder
End Function

'Get the total weight of elements in a layer
Private Sub getTotalEditions()
    Dim i As Long
    totalEditions = 0
    On Error Resume Next
    For i = 0 To ListType.ListCount - 1
        totalEditions = totalEditions + layerConfigurations(i).typeSize
    Next i
    If Err.Number = 0 Then
        lblTotalEditions.Caption = totalEditions
    Else
        lblTotalEditions.Caption = Language.Item("Tips36")
        totalEditions = -1
        Err.Clear
    End If
End Sub

'Select a type in the types listbox.
Private Sub ListType_Click()
    If ListType.SelCount = 0 Then Exit Sub
    Dim i As Long
    txtTypeDisplayName = layerConfigurations(ListType.ListIndex).displayName
    ListLayer.Clear
    For i = 0 To layerConfigurations(ListType.ListIndex).layersSize - 1
        ListLayer.AddItem layerConfigurations(ListType.ListIndex).layersOrder(i)
    Next i
    txtNumType = Split(ListType.Text, vbTab)(0)
    If ListLayer.ListCount <> 0 Then ListLayer.ListIndex = 0
    DoEvents
    previewPic
End Sub

Private Sub Timer1_Timer()
    previewPic
    Timer1.Enabled = False
End Sub

'Input how many images to generate for this type
Private Sub txtNumType_Change()
    If ListType.SelCount = 0 Then Exit Sub
    If Val(txtNumType.Text) > 2100000000 Then
        lblTotalEditions.Caption = Language.Item("Tips36")
        totalEditions = -1
        Exit Sub
    End If
    
    If IsNumeric(txtNumType.Text) Then
        layerConfigurations(ListType.ListIndex).typeSize = txtNumType
        DoEvents
        ListType.list(ListType.ListIndex) = txtNumType & vbTab & Split(ListType.Text, vbTab)(1)
        Call getTotalEditions
    End If
End Sub

Private Sub txtTypeDisplayName_Change()
    If ListType.SelCount = 0 Then Exit Sub
    layerConfigurations(ListType.ListIndex).displayName = txtTypeDisplayName
    DoEvents
    ListType.list(ListType.ListIndex) = layerConfigurations(ListType.ListIndex).typeSize & vbTab & txtTypeDisplayName
End Sub

Private Sub cmdUpType_Click()
   If ListType.ListIndex > 0 Then
        moveType -1
    End If
End Sub

Private Sub cmdDownType_Click()
    If ListType.ListIndex <> -1 And ListType.ListIndex < ListType.ListCount - 1 Then
        moveType 1
    End If
End Sub

'Change the order of types
Private Sub moveType(k As Long)
        Dim s As String, oldIndex As Long, newIndex As Long, temp As layerConfig
        oldIndex = ListType.ListIndex
        newIndex = oldIndex + k
        temp = layerConfigurations(oldIndex)
        layerConfigurations(oldIndex) = layerConfigurations(newIndex)
        layerConfigurations(newIndex) = temp
        s = ListType.Text
        ListType.RemoveItem oldIndex
        ListType.AddItem s, newIndex
        DoEvents
        ListType.ListIndex = newIndex
End Sub
Private Sub cmdTop_Click()
    If ListLayer.ListIndex > 0 Then
        moveLayer -ListLayer.ListIndex
    End If
End Sub

Private Sub cmdUpLayer_Click()
    If ListLayer.ListIndex > 0 Then
        moveLayer -1
    End If
End Sub

Private Sub cmdDownLayer_Click()
    If ListLayer.ListIndex <> -1 And ListLayer.ListIndex < ListLayer.ListCount - 1 Then
        moveLayer 1
    End If
End Sub

Private Sub cmdBottom_Click()
    If ListLayer.ListIndex <> -1 And ListLayer.ListIndex < ListLayer.ListCount - 1 Then
        moveLayer ListLayer.ListCount - ListLayer.ListIndex - 1
    End If
End Sub

'Change the order of layers
Private Sub moveLayer(k As Long)
    Dim i As Long, oldIndex As Long, newIndex As Long, typeIndex As Long, temp As String
    typeIndex = ListType.ListIndex
    oldIndex = ListLayer.ListIndex
    newIndex = oldIndex + k
    
'    temp = layerConfigurations(typeIndex).layersOrder(oldIndex)
'    layerConfigurations(typeIndex).layersOrder(oldIndex) = layerConfigurations(typeIndex).layersOrder(newIndex)
'    layerConfigurations(typeIndex).layersOrder(newIndex) = temp
    
    temp = ListLayer.Text
    ListLayer.RemoveItem oldIndex
    ListLayer.AddItem temp, newIndex
    ListLayer.ListIndex = newIndex
    
    For i = 0 To ListLayer.ListCount - 1
        layerConfigurations(typeIndex).layersOrder(i) = ListLayer.list(i)
    Next i
    
    DoEvents
    previewPic
End Sub

'Double click to preview the generated image.
Private Sub ListLayer_DblClick()
    previewPic
End Sub

'Copy a type
Private Sub cmdCopyType_Click()
    Dim i As Long, k As Long
    If ListType.SelCount = 0 Then Exit Sub
    k = ListType.ListIndex
    If layerConfigurations(k).typeName = oneArtFolderName Then Exit Sub
    ListType.AddItem ListType.Text, k
    ReDim Preserve layerConfigurations(UBound(layerConfigurations) + 1)
    For i = UBound(layerConfigurations) To k + 1 Step -1
        layerConfigurations(i) = layerConfigurations(i - 1)
    Next i
    ListType.ListIndex = k + 1
    Call getTotalEditions
End Sub

'Delete the selected type
Private Sub cmdDelType_Click()
    Dim i As Long, k As Long
    If ListType.SelCount = 0 Then Exit Sub
    If ListType.ListCount = 1 Then
        showTips Language.Item("Tips6")
        Exit Sub
    End If
    k = ListType.ListIndex
    ListType.RemoveItem ListType.ListIndex
    For i = k To ListType.ListCount - 1
        layerConfigurations(i) = layerConfigurations(i + 1)
    Next i
    ReDim Preserve layerConfigurations(ListType.ListCount - 1)
    DoEvents
    If k = 0 Then ListType.ListIndex = k Else ListType.ListIndex = k - 1
    Call getTotalEditions
End Sub

'Delete the selected layer
Private Sub cmdDelLayer_Click()
    Dim i As Long, k As Long
    If ListLayer.SelCount = 0 Then Exit Sub
    If ListLayer.ListCount = 1 Then
        showTips Language.Item("Tips7")
        Exit Sub
    End If
    k = ListLayer.ListIndex
    ListLayer.RemoveItem k
    ReDim layerConfigurations(ListType.ListIndex).layersOrder(ListLayer.ListCount - 1)
    layerConfigurations(ListType.ListIndex).layersSize = ListLayer.ListCount
    For i = 0 To ListLayer.ListCount - 1
        layerConfigurations(ListType.ListIndex).layersOrder(i) = ListLayer.list(i)
    Next i
    If k = 0 Then ListLayer.ListIndex = k Else ListLayer.ListIndex = k - 1
    DoEvents
    previewPic
End Sub

'Delete the selected layer if the DELETE key is pressed
Private Sub ListLayer_KeyDown(KeyCode As Integer, Shift As Integer)
    If ListLayer.SelCount = 0 Then Exit Sub
    If KeyCode = vbKeyDelete Then
        cmdDelLayer_Click
    End If
End Sub

'Flag Bypass DNA
Private Sub cmdBypassDNA_Click()
    If ListLayer.SelCount = 0 Or LCase(layerConfigurations(ListType.ListIndex).typeName) = oneArtFolderName Then Exit Sub
    If Right(ListLayer.Text, 1) <> "*" Then
        ListLayer.list(ListLayer.ListIndex) = ListLayer.Text & "*"
    Else
        ListLayer.list(ListLayer.ListIndex) = Left(ListLayer.Text, Len(ListLayer.Text) - 1)
    End If
    layerConfigurations(ListType.ListIndex).layersOrder(ListLayer.ListIndex) = ListLayer.Text
End Sub

'Set files weight in one layer
Private Sub cmdAddWeight_Click()
    If ListLayer.SelCount = 0 Then Exit Sub
    
    Dim k As Long, a As Variant
    Dim path As String, iName As String
    k = 0
    path = layersDir & "\" & layerConfigurations(ListType.ListIndex).typeName & "\" & Split(ListLayer.Text, "*")(0) & "\"

    iName = Dir(path)
    Do While iName <> ""
        a = Split(GetFileName(iName), rarityDelimiter)
        'InStr(iName, "#") = 0
        If LCase(Right(iName, 4)) = ".png" And Not (UBound(a) <> 0 And IsNumeric(a(UBound(a)))) Then
            DoEvents
            Name path & iName As path & GetFileName(iName) & rarityDelimiter & "10.png"
            showTips Language.Item("Tips26")
            k = k + 1
        End If
        iName = Dir()
    Loop
    If k = 0 Then showTips Language.Item("Tips38")
End Sub

Private Sub cmdDelWeight_Click()
    If ListLayer.SelCount = 0 Then Exit Sub
    
    Dim k As Long, a As Variant
    Dim path As String, iName As String
    k = 0
    path = layersDir & "\" & layerConfigurations(ListType.ListIndex).typeName & "\" & Split(ListLayer.Text, "*")(0) & "\"

    iName = Dir(path)
    Do While iName <> ""
        a = Split(GetFileName(iName), rarityDelimiter)
        'InStr(iName, "#") = 0
        If LCase(Right(iName, 4)) = ".png" And UBound(a) <> 0 And IsNumeric(a(UBound(a))) Then
            DoEvents
            Name path & iName As path & a(0) & ".png"
            showTips Language.Item("Tips26")
            k = k + 1
        End If
        iName = Dir()
    Loop
    If k = 0 Then showTips Language.Item("Tips38")
End Sub

'Import configurations from Config.json file
Private Sub cmdImport_Click()
    If checkLayers = True Then importConfig
End Sub
Private Sub cmdReload_Click()
    If checkLayers = True Then reloadDir
End Sub

'Preview the image after setting the layer order
Private Sub previewPic()
    If ListType.ListCount = 0 Or ListLayer.ListCount = 0 Then
        'showTips Language.Item("Tips9")
        Exit Sub
    ElseIf ListType.ListIndex < 0 Then
        'showTips Language.Item("Tips10")
        Exit Sub
    End If
    Dim i As Long
    Dim fn As Integer
    Dim pngName As String
    Dim tempS As String
    On Error GoTo ErrorTips:
    'showTips ""
    picPreview.Cls
    If LCase(layerConfigurations(ListType.ListIndex).typeName) = oneArtFolderName Then
        DrawPng layersDir & "\" & layerConfigurations(ListType.ListIndex).typeName & "\" & ListLayer.Text
    Else
        For i = ListLayer.ListCount - 1 To 0 Step -1
            pngName = Dir(layersDir & "\" & layerConfigurations(ListType.ListIndex).typeName & "\" & Split(ListLayer.list(i), "*")(0) & "\")
            If LCase(pngName) = "none.png" Then
                tempS = Dir()
                If tempS <> "" Then pngName = tempS
            End If
            DrawPng layersDir & "\" & layerConfigurations(ListType.ListIndex).typeName & "\" & Split(ListLayer.list(i), "*")(0) & "\" & pngName
        Next i
    End If
    Exit Sub
ErrorTips:
    Close
    showTips Language.Item("Tips11")
End Sub

Private Sub cmdReloadLayer_Click()
    If ListType.SelCount = 0 Then Exit Sub
    
    Dim configFile As String
    Dim foldername As String
    Dim tempS() As String
    Dim fn As Integer
    Dim i As Long, j As Long, k As Long, m As Long, n As Long
    Dim nameExist As Boolean
   
    'Get layers order information:
    k = ListType.ListIndex
    tempS = layerConfigurations(k).layersOrder
    foldername = layerConfigurations(k).typeName
    'layer folders or 1/1 files
    If LCase(foldername) = oneArtFolderName Then layerConfigurations(k).layersOrder = getSpecialFiles(layersDir & "\" & foldername) _
    Else layerConfigurations(k).layersOrder = getLayersOrder(layersDir & "\" & foldername)
    'Is the type folder empty?
    If (CStr(Join(layerConfigurations(k).layersOrder, ""))) = "" Then
        showTips foldername & " " & Language.Item("Tips3")
        layerConfigurations(k).layersSize = 0
    Else
        layerConfigurations(k).layersSize = UBound(layerConfigurations(k).layersOrder) + 1
    End If
    'the number of 1/1 editions = the number of 1/1 files
    If LCase(foldername) = oneArtFolderName Then layerConfigurations(k).typeSize = layerConfigurations(k).layersSize
    
    n = 0
    For i = 0 To ListLayer.ListCount - 1
        If i > ListLayer.ListCount - 1 - n Then Exit For
        nameExist = False
        For j = 0 To layerConfigurations(k).layersSize - 1
            If Split(ListLayer.list(i), "*")(0) = layerConfigurations(k).layersOrder(j) Then
                nameExist = True
                Exit For
            End If
        Next j
        If nameExist = False Then
            ListLayer.RemoveItem i
            n = n + 1
            i = i - 1
        End If
    Next i
    
    For i = 0 To layerConfigurations(k).layersSize - 1
        nameExist = False
        For j = 0 To ListLayer.ListCount - 1
            If layerConfigurations(k).layersOrder(i) = Split(ListLayer.list(j), "*")(0) Then
                nameExist = True
                Exit For
            End If
        Next j
        If nameExist = False Then ListLayer.AddItem layerConfigurations(k).layersOrder(i)
    Next i
    
    For i = 0 To ListLayer.ListCount - 1
        layerConfigurations(k).layersOrder(i) = ListLayer.list(i)
    Next i
End Sub

'Save the current configuration to the config.txt and order.txt files for each layer.
Private Sub cmdSaveOrder_Click()
    Dim i As Long, j As Long, fn As Integer
    Close
    If ListType.ListCount = 0 Then
        showTips Language.Item("Tips12")
        Exit Sub
    End If

    Set config = New JsonBag
    With config
        .Whitespace = frmSetting.chkWhiteSpace.Value = Checked
        .WhitespaceIndent = 2
        .DecimalMode = False
        .Clear
        .IsArray = True
    End With
    
    For i = 0 To ListType.ListCount - 1
        With config.AddNewObject()
            .Item("typeName") = layerConfigurations(i).typeName
            .Item("displayName") = layerConfigurations(i).displayName
            .Item("typeSize") = layerConfigurations(i).typeSize
            .Item("layersSize") = layerConfigurations(i).layersSize
            If layerConfigurations(i).layersSize > 0 Then 'And LCase(layerConfigurations(i).typeName) <> oneArtFolderName
                With .AddNewArray("layersOrder")
                    For j = 0 To layerConfigurations(i).layersSize - 1
                        .Item = layerConfigurations(i).layersOrder(j)
                    Next j
                End With
            End If
        End With
    Next i
    fn = FreeFile
    Open layersDir & "\Config.json" For Output As #fn
    Print #fn, config.JSON
    Close #fn
    showTips Language.Item("Tips13")
End Sub

'**********************************************************
'
'                    Generate images
'
'**********************************************************

Private Sub cmdStart_Click()
    cmdSaveOrder_Click
    If totalEditions = -1 Then
        showTips Language.Item("Tips37")
        Exit Sub
    End If
    
    If checkElements() > 0 Then
        Shell "explorer " & buildDir & "\Error.txt", vbNormalFocus
        Exit Sub
    End If
    
    Dim layerConfigIndex As Long
    Dim editionCount As Long
    Dim fileName As String
    Dim i As Long, j As Long, k As Long, maxSize As Double
    Dim failedCount As Long
    Dim allErrInfo As String
    Dim startNumber As Long
    Dim fn As Integer

    Dim graphics As Long
    Dim bitmap As Long
    Dim graphicsCache As Long
    Dim bitmapCache As Long
    Dim Image As Long
    Dim imgWidth As Long
    Dim imgHeight As Long
    Dim picGraphics As Long
    Dim picWidth As Long
    Dim picHeight As Long
    Dim backgroundColor As Long
    Dim ColorBTNFACE As Long
    ColorBTNFACE = CLng("&HFF" & Hex(GetSysColor(15)))
    
    'Disable hibernation
    SetThreadExecutionState Execution_State.ES_SYSTEM_REQUIRED Or Execution_State.ES_DISPLAY_REQUIRED Or Execution_State.ES_CONTINUOUS
    
    If chkClean.Value = Checked Then
        If Dir(buildDir & "\images\*.*") <> "" Then Kill buildDir & "\images\*.*"
        If Dir(buildDir & "\json\*.*") <> "" Then Kill buildDir & "\json\*.*"
    End If
    allErrInfo = ""
    Set DNA = New Collection
    layerConfigIndex = 0
    'eth net start from 1, sol start from 0
    If frmSetting.OptionNetwork(0).Value = True Then startNumber = 1 Else startNumber = 0
    editionCount = startNumber
    Randomize

    'Get image size from settings.
    If frmSetting.chkResize.Value = Checked Then
        imgWidth = Val(frmSetting.txtWidth)
        imgHeight = Val(frmSetting.txtHeight)
    End If
    'Get background color according to settings or randomly generated or background transparent.
    backgroundColor = getColor()
    'Calling public function from Public.bas. Get the Metadata information from the settings and create a Metadata template (JsonBag class).
    Call GetTemplateJB

    'Traverse the layer configurations
    For layerConfigIndex = 0 To UBound(layerConfigurations)
    
        'If type size is 0 , skip to the next type.
        If layerConfigurations(layerConfigIndex).typeSize = 0 Then GoTo NextType
        
        'If there is no layer folder, skip to the next type.
        If layerConfigurations(layerConfigIndex).layersSize = 0 Then
            showTips Language.Item("Tips14") & " " & layerConfigurations(layerConfigIndex).typeName
            GoTo NextType
        End If
        If ListType.ListCount <> 0 Then ListType.ListIndex = layerConfigIndex
        If ListLayer.ListCount <> 0 Then ListLayer.ListIndex = 0
        'If it's a 1/1 folder, do special processing.
        If LCase(layerConfigurations(layerConfigIndex).typeName) = LCase(oneArtFolderName) Then
            For i = 0 To layerConfigurations(layerConfigIndex).layersSize - 1
                DoEvents
                If IsUnloading Then Exit Sub
                showTips Language.Item("Tips15") & editionCount - startNumber + 1 & "/" & totalEditions
                fileName = layerConfigurations(layerConfigIndex).layersOrder(i)
                If Dir(layersDir & "\" & oneArtFolderName & "\" & fileName) <> "" Then
                    isDnaUnique fileName
                    FileCopy layersDir & "\" & oneArtFolderName & "\" & fileName, buildDir & "\images\" & editionCount & GetExtensionName(fileName)
                    DrawPng buildDir & "\images\" & editionCount & GetExtensionName(fileName)
                    creatMetadata editionCount, fileName
                    saveMetadataFile editionCount
                    editionCount = editionCount + 1
                End If
            Next i
            GoTo NextType
        End If
        'If all layer folders for this type are empty, skip to the next type.
        If layersSetup(layerConfigIndex) = False Then
            showTips Language.Item("Tips16") & " " & layerConfigurations(layerConfigIndex).typeName
            GoTo NextType
        End If
        'Calculate the maximum number of combinations this type can have, and take the minimum value of it and typeSize.
        maxSize = 1
        For k = 0 To UBound(layers)
            If layers(k).bypassDNA = False Then
                maxSize = maxSize * (UBound(layers(k).elements) + 1)
            End If
        Next k
        If maxSize > layerConfigurations(layerConfigIndex).typeSize Then maxSize = layerConfigurations(layerConfigIndex).typeSize

        'Prepare for drawing memory bitmaps and previews.
        'If no image size is set, read the size of an element png and use it as the image size.
        If frmSetting.chkResize.Value = Unchecked Then
            GdipLoadImageFromFile StrPtr(layers(0).elements(0).path), Image
            GdipGetImageWidth Image, imgWidth
            GdipGetImageHeight Image, imgHeight
            GdipDisposeImage Image
        End If
        'Set the preview image size and make sure that the image aspect ratio is unchanged¡£
        If imgWidth / imgHeight >= picPreview.ScaleWidth / picPreview.ScaleHeight Then
            picHeight = imgHeight * picPreview.ScaleWidth / imgWidth
            picWidth = picPreview.ScaleWidth
        Else
            picWidth = imgWidth * picPreview.ScaleHeight / imgHeight
            picHeight = picPreview.ScaleHeight
        End If
        'Create memory bitmap, canvas as graphics
        CreateBitmapWithGraphics bitmap, graphics, imgWidth, imgHeight
        'Image smoothing settings
        If frmSetting.chkSmoothing.Value = Checked Then
            GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
        End If
        'Set the preview image as canvas picGraphics
        GdipCreateFromHDC picPreview.hDC, picGraphics

        'Inner loop, generate maxSize images of this type
        i = 0
        failedCount = 0
        Do While i < maxSize
            If isDnaUnique(createDNA(failedCount)) Then
                DoEvents
                If IsUnloading Then Exit Sub
                showTips Language.Item("Tips15") & editionCount - startNumber + 1 & "/" & totalEditions
                'Clear memory bitmap and set background (Hexadecimal ARGB format)
                GdipGraphicsClear graphics, getColor
                'Draw each element png to the memory bitmap in turn, from bottom layer to top layer
                For k = UBound(newDNA) To 0 Step -1
                    If newDNA(k) <> -1 Then
                        GdipLoadImageFromFile StrPtr(layers(k).elements(newDNA(k)).path), Image
                        GdipDrawImageRect graphics, Image, 0, 0, imgWidth, imgHeight
                        GdipDisposeImage Image
                        layers(k).elements(newDNA(k)).usedCount = layers(k).elements(newDNA(k)).usedCount + 1
                    End If
                Next k
                SaveImageToPNG bitmap, buildDir & "\images\" & editionCount & ".png"
                
                'Preview the generated image. To solve the problem of picture.cls splash screen, I introduced image caching.
                CreateBitmapWithGraphics bitmapCache, graphicsCache, imgWidth, imgHeight
                GdipGraphicsClear graphicsCache, ColorBTNFACE
                GdipDrawImageRect graphicsCache, bitmap, 0, 0, imgWidth, imgHeight
                GdipDrawImageRect picGraphics, bitmapCache, 0, 0, picWidth, picHeight
                GdipDeleteGraphics graphicsCache
                GdipDisposeImage bitmapCache
                
                'Write json file
                creatMetadata editionCount
                saveMetadataFile editionCount
                editionCount = editionCount + 1
                i = i + 1
            Else
                failedCount = failedCount + 1
                If failedCount > Val(frmSetting.txtDnaTryTimes) Then
                    showTips Language.Item("Tips17") & layerConfigurations(layerConfigIndex).typeSize & " " & layerConfigurations(layerConfigIndex).typeName
                    Exit Do
                End If
            End If
        Loop
        'Delete the memory bitmap
        GdipDeleteGraphics graphics
        GdipDisposeImage bitmap
        GdipDeleteGraphics picGraphics
NextType:
    Next layerConfigIndex
    If editionCount = 1 Then
        showTips Language.Item("Tips18")
    Else
        fn = FreeFile
        Open buildDir & "\DNAList.txt" For Output As #fn
        For i = 1 To DNA.Count
            Print #fn, i + startNumber - 1 & " -> " & DNA(i)
        Next i
        Close #fn
        showTips Language.Item("Tips19")
        Shell "explorer " & buildDir & "\images", vbNormalFocus
    End If
    If Dir(buildDir & "\Error.txt") <> "" Then Shell "explorer " & buildDir & "\Error.txt", vbNormalFocus
    'Undisable computer hibernation
    SetThreadExecutionState Execution_State.ES_CONTINUOUS
End Sub

'Check the element image format
Private Function checkElements() As Long
    Dim layerConfigIndex As Long
    Dim layersOrder() As String
    Dim typeName As String
    Dim tempPath As String
    Dim tempName As String
    Dim Image As Long
    Dim i As Long, j As Long, k As Long
    Dim errInfo() As String
    Dim pngErrCount As Long
    ReDim errInfo(0)
    showTips Language.Item("Tips20")
    DoEvents
    If IsUnloading Then Exit Function
    errInfo(0) = "-------------------------- Error information --------------------------" & vbCrLf
    pngErrCount = 0
    For layerConfigIndex = 0 To UBound(layerConfigurations)
        layersOrder = layerConfigurations(layerConfigIndex).layersOrder
        typeName = layerConfigurations(layerConfigIndex).typeName
        'no layer folder
        If layerConfigurations(layerConfigIndex).layersSize = 0 Then
            ReDim Preserve errInfo(UBound(errInfo) + 1)
            errInfo(UBound(errInfo)) = "Empty folder: " & layersDir & "\" & typeName & "\"
            GoTo NextType
        End If
        If LCase(layerConfigurations(layerConfigIndex).typeName) = LCase(oneArtFolderName) Then GoTo NextType
        For i = 0 To UBound(layersOrder)
            tempPath = layersDir & "\" & typeName & "\" & Split(layersOrder(i), "*")(0) & "\"
            tempName = Dir(tempPath)
            j = 0: k = 0
            Do While tempName <> ""
                If LCase(Right(tempName, 4)) = ".png" Then
                    If GdipLoadImageFromFile(StrPtr(tempPath & tempName), Image) <> Ok Then
                        ReDim Preserve errInfo(UBound(errInfo) + 1)
                        errInfo(UBound(errInfo)) = " ! Error file: " & tempPath & tempName
                        pngErrCount = pngErrCount + 1
                        j = j + 1
                    Else
                        k = k + 1
                    End If
                    GdipDisposeImage Image
                End If
                tempName = Dir()
            Loop
            If k = 0 Then
                ReDim Preserve errInfo(UBound(errInfo) + 1)
                If j = 0 Then errInfo(UBound(errInfo)) = "Empty folder: " & tempPath Else errInfo(UBound(errInfo)) = "No valid png: " & tempPath
            End If
        Next i
NextType:
    Next layerConfigIndex
    If UBound(errInfo) > 0 Then
        Open buildDir & "\Error.txt" For Output As #1
        For i = 0 To UBound(errInfo)
            Print #1, errInfo(i)
            Print #1,
        Next i
        Close #1
        showTips UBound(errInfo) & " " & Language.Item("Tips21")
    Else
        If Dir(buildDir & "\Error.txt") <> "" Then Kill buildDir & "\Error.txt"
    End If
    checkElements = pngErrCount
End Function

'Get background color from settings or randomly generate, transparent if none
Private Function getColor() As Long
    'If not checked, return transparent color
    If frmSetting.chkBackground = Unchecked Then
        getColor = 0
        Exit Function
    End If
    'If the color is 6-bit hexadecimal RGB format, add FF in front as A (Opacity).
    If frmSetting.chkStaticColor.Value = Checked Then
        getColor = CLng("&H" & Right("FF" & Replace(frmSetting.txtStaticColor, "#", ""), 8))
    Else
        getColor = genColor()
    End If
End Function

'Randomly generate HSL color and convert to RGB color.
Private Function genColor() As Long
    Dim hue As Integer, rgb() As String, h As String
    Randomize
    hue = Int(Rnd() * 360)
    rgb = Split(HSL2RGB(hue, 100, Val(frmSetting.txtLightness)), ",")
    h = Right("0" & Hex(rgb(0)), 2) & Right("0" & Hex(rgb(1)), 2) & Right("0" & Hex(rgb(2)), 2)
    genColor = CLng("&H" & "FF" & h)
End Function

'Get the metadata value from the current layer name, editionCount and newDNA() information, then added to the metadata template.
Private Sub creatMetadata(ByVal editionCount As Long, Optional ByVal specialFileName As String = "")
    Dim namePrefix As String
    Dim imageBaseURL As String
    Dim extensionName As String
    Dim thisValue As String
    Dim ignoreNONE As Boolean
    Dim i As Long
    If frmSetting.chkIgnoreNONE.Value = Checked Then ignoreNONE = True Else ignoreNONE = False
    namePrefix = frmSetting.txtNamePrefix.Text & " #"
    imageBaseURL = frmSetting.txtImageBaseURL.Text
    If specialFileName <> "" Then extensionName = GetExtensionName(specialFileName) Else extensionName = ".png"
    JB.Item("name") = namePrefix & editionCount
    JB.Item("image") = imageBaseURL & editionCount & extensionName
    If frmSetting.OptionNetwork(1).Value = True Then
        JB.Item("properties").Item("files")(1).Item("uri") = imageBaseURL & editionCount & extensionName
    End If
    JB.ItemJSON("attributes") = "[]"
    If specialFileName <> "" Then
        With JB.Item("attributes")
            With .AddNewObject()
                .Item("trait_type") = "1/1"
                thisValue = cleanName(specialFileName)
                If IsNumeric(thisValue) Then .Item("value") = Val(thisValue) Else .Item("value") = thisValue
            End With
        End With
    Else
        With JB.Item("attributes")
            For i = 0 To UBound(newDNA)
                If newDNA(i) = -1 Then thisValue = "NONE" Else thisValue = layers(i).elements(newDNA(i)).trait_value
                If Not (UCase(thisValue) = "NONE" And ignoreNONE = True) Then
                    With .AddNewObject()
                        .Item("trait_type") = layers(i).trait_type
                        If IsNumeric(thisValue) Then .Item("value") = Val(thisValue) Else .Item("value") = thisValue
                    End With
                End If
            Next i
        End With
    End If
End Sub

'Save json file.
Private Sub saveMetadataFile(ByVal editionCount As Long)
    Dim fn As Integer
    fn = FreeFile
    Open buildDir & "\json\" & editionCount & ".json" For Output As #fn
        Print #fn, JB.JSON
    Close #fn
End Sub

'Get the layers configuration infomation of a type, including elements,
Private Function layersSetup(ByVal layerConfigIndex As Long) As Boolean
    Dim i As Long, j As Long, k As Long, n As Long, maxWeight As Long, maxWeightIndex As Long
    Dim layersOrder() As String
    Dim typeName As String
    Dim tempName As String
    layersOrder = layerConfigurations(layerConfigIndex).layersOrder
    typeName = layerConfigurations(layerConfigIndex).typeName
    k = 0
    For i = 0 To UBound(layersOrder)
        tempName = Split(layersOrder(i), "*")(0)
        'Check the element files in the layer folder.
        'Is the layer folder empty?
        If getElements(layersDir & "\" & typeName & "\" & tempName & "\") = False Then
            showTips tempName & " " & Language.Item("Tips3")
        Else
            ReDim Preserve layers(k)
            layers(k).id = k
            layers(k).Name = tempName
            layers(k).trait_type = replaceSign(Split(tempName, nameDelimiter)(0))
            If Right(layersOrder(i), 1) = "*" Then layers(k).bypassDNA = True Else layers(k).bypassDNA = False
            layers(k).elements = elements
            layers(k).totalWeight = 0
            maxWeight = 0
            For j = 0 To UBound(elements)
                layers(k).totalWeight = layers(k).totalWeight + elements(j).weight
                If maxWeight < elements(j).weight Then
                    maxWeight = elements(j).weight
                    maxWeightIndex = j
                End If
            Next j
            n = 0
            For j = 0 To UBound(elements)
                layers(k).elements(j).usableMax = Int(layerConfigurations(layerConfigIndex).typeSize * elements(j).weight / layers(k).totalWeight - 0.09) + 1
                n = n + layers(k).elements(j).usableMax
                layers(k).elements(j).usedCount = 0
            Next j
            If layerConfigurations(layerConfigIndex).typeSize > n Then layers(k).elements(maxWeightIndex).usableMax = layers(k).elements(maxWeightIndex).usableMax + layerConfigurations(layerConfigIndex).typeSize - n
            k = k + 1
        End If
    Next i
    If k = 0 Then layersSetup = False Else layersSetup = True
End Function

'Get all elements infomation of a layer
Private Function getElements(ByVal path As String) As Boolean
    Dim i As Long
    Dim iName As String
    i = 0
    iName = Dir(path)
    Do While iName <> ""
        If LCase(Right(iName, 4)) = ".png" Then
            ReDim Preserve elements(i)
            With elements(i)
                .id = i
                .trait_value = cleanName(iName)
                .fileName = iName
                .skipLayer = getSkipLayer(iName)
                .path = path & iName
                .weight = getRarityWeight(iName)
            End With
            i = i + 1
        End If
        iName = Dir()
    Loop
    If i = 0 Then getElements = False Else getElements = True
End Function

'Remove weight from filename, leaving only the clean filenamea as the metadata attribute value.
Private Function cleanName(ByVal Str As String) As String
    Str = Split(GetFileName(Str), rarityDelimiter)(0)
    Str = Split(Str, signSkip)(0)
    'Str = Split(Str, signOnly)(0)
    Str = replaceSign(Str)
    cleanName = Str
End Function

'Under Windows system, the file name cannot contain characters such as ":" , "/" , these characters can be replaced with signs,
'which are replaced by replaceSign() function when writing json.replace, eg. _COLONS_ -> : , _SLASH_ -> /
Private Function replaceSign(ByVal Str As String) As String
    If frmSetting.chkReplace = Checked Then
        Dim i As Integer
        For i = 0 To frmSetting.txtSign.UBound
            If frmSetting.txtSign(i).Text <> "" Then Str = Replace(Str, frmSetting.txtSign(i).Text, frmSetting.txtReplace(i).Text)
        Next i
    End If
    replaceSign = Str
End Function

'A certain element may be mutually exclusive with a certain layer, this getSkipLayer() function finds out the layer name.
Private Function getSkipLayer(ByVal Str As String) As String()
    'the 0 is name, skip layer start from 1
    Str = Split(GetFileName(Str), rarityDelimiter)(0)
    getSkipLayer = Split(Str, signSkip)
End Function

'Detach weight from filename
Private Function getRarityWeight(ByVal Str As String) As Long
    Dim nameWithoutExtension As String
    Dim a As Variant
    nameWithoutExtension = GetFileName(Str)
    a = Split(nameWithoutExtension, rarityDelimiter)
    If UBound(a) <> 0 And IsNumeric(a(UBound(a))) Then
        getRarityWeight = a(UBound(a))
    Else
        getRarityWeight = 1
    End If
End Function

'Preview in picPreview
Private Sub DrawPng(ByVal pngfile As String)
    Dim Graphic As Long
    Dim Image As Long
    Dim imgWidth As Long
    Dim imgHeight As Long
    GdipCreateFromHDC picPreview.hDC, Graphic
    GdipSetSmoothingMode Graphic, SmoothingModeAntiAlias
    GdipLoadImageFromFile StrPtr(pngfile), Image
    GdipGetImageWidth Image, imgWidth
    GdipGetImageHeight Image, imgHeight
    If imgWidth / imgHeight >= picPreview.ScaleWidth / picPreview.ScaleHeight Then
        imgHeight = imgHeight * picPreview.ScaleWidth / imgWidth
        imgWidth = picPreview.ScaleWidth
    Else
        imgWidth = imgWidth * picPreview.ScaleHeight / imgHeight
        imgHeight = picPreview.ScaleHeight
    End If
    GdipDrawImageRect Graphic, Image, 0, 0, imgWidth, imgHeight
    GdipDisposeImage Image
    GdipDeleteGraphics Graphic
End Sub

'Randomly create a DNA based on the current layers() content
Private Function createDNA(ByVal failedCount As Long) As String
    Dim thisDNA As String
    Dim i As Long, j As Long, k As Long
    Dim maxTryTimes As Long
    Dim random As Long
    Dim skipLayerName() As String
    'Get a random DNA
    thisDNA = ""
    maxTryTimes = Val(frmSetting.txtDnaTryTimes) / 2
    ReDim newDNA(UBound(layers))
    For i = 0 To UBound(layers)
        If newDNA(i) = -1 Then GoTo NEXTI
        k = 0
        Do While True
            'number between 0 - totalWeight
            random = Int(Rnd() * layers(i).totalWeight)
            For j = 0 To UBound(layers(i).elements)
                'subtract the current weight from the random weight until we reach a sub zero value.
                random = random - layers(i).elements(j).weight
                If random < 0 Then Exit For
            Next j
           'When an element is used enough times (the number of NFTs * the weight of the element/total weight), it is no longer used
           'and the element is re-extracted. Unless the number of failures to generate independent DNA is greater than 10000.
            If layers(i).elements(j).usedCount < layers(i).elements(j).usableMax Or failedCount > maxTryTimes Or k > maxTryTimes Then
                If layers(i).bypassDNA = False Then
                    If thisDNA = "" Then thisDNA = layers(i).elements(j).trait_value Else thisDNA = thisDNA & "-" & layers(i).elements(j).trait_value
                End If
                newDNA(i) = j
                Exit Do
            End If
            k = k + 1
        Loop
        'Skip some layers, start from array 1
        skipLayerName() = layers(i).elements(newDNA(i)).skipLayer
        If UBound(skipLayerName) <> 0 Then
            For k = 1 To UBound(skipLayerName)
                For j = 0 To UBound(layers)
                    If LCase(layers(j).Name) = LCase(skipLayerName(k)) Or LCase(Split(layers(j).Name, nameDelimiter)(0)) = LCase(skipLayerName(k)) Then newDNA(j) = -1
                Next j
            Next k
        End If
NEXTI:
    Next i
    createDNA = thisDNA
End Function

'Determine whether the current DNA exists
Function isDnaUnique(ByVal thisDNA As String) As Boolean
    Dim tempS As String
    Err.Clear
    On Error Resume Next
    tempS = DNA(thisDNA)
    If Err.Number <> 0 Then
        isDnaUnique = True
        DNA.Add thisDNA, thisDNA
    Else
        isDnaUnique = False
    End If
End Function

'******************************************************************************************
'
'   Update the files number and JSON content after manually selecting the images
'
'******************************************************************************************
Private Sub cmdUpdate_Click()
    Dim imagesDir As String, jsonDir As String, tempImagesDir As String, tempJsonDir As String
    Dim i As Long, j As Long, k As Long
    Dim nameArray() As Long, abstractedIndexes() As Long
    Dim tempName As String, extensionName As String
    Dim fn As Integer
    Dim IsSolana As Boolean
    
    imagesDir = buildDir & "\images"
    jsonDir = buildDir & "\json"
    tempImagesDir = buildDir & "\tempimages"
    tempJsonDir = buildDir & "\tempjson"

    If Dir(imagesDir, vbDirectory) = "" Then
        showTips Language.Item("Tips22")
        Exit Sub
    ElseIf Dir(jsonDir, vbDirectory) = "" Then
        showTips Language.Item("Tips23")
        Exit Sub
    End If
    If Dir(tempImagesDir, vbDirectory) <> "" Then
        If Dir(tempImagesDir & "\*.*") <> "" Then Kill tempImagesDir & "\*.*"
        RmDir tempImagesDir
    End If
    If Dir(tempJsonDir, vbDirectory) <> "" Then
        If Dir(tempJsonDir & "\*.*") <> "" Then Kill tempJsonDir & "\*.*"
        RmDir tempJsonDir
    End If

    'Read all png files in the images folder and put the number part of the name into an array
    k = 0
    tempName = Dir(imagesDir & "\")
    Do While tempName <> ""
        DoEvents
        If IsUnloading Then Exit Sub
        showTips Language.Item("Tips24")
        If LCase(Right(tempName, 4)) = ".png" Then
            If IsNumeric(Left(tempName, Len(tempName) - 4)) Then
                ReDim Preserve nameArray(k)
                nameArray(k) = Val(Left(tempName, Len(tempName) - 4))
                k = k + 1
            End If
        End If
        tempName = Dir()
    Loop
    If k = 0 Then
        showTips Language.Item("Tips25")
        Exit Sub
    End If
    'Reorder the array of png filenames from smallest to largest
    ArraySort nameArray

    'Initialize a new numbered array.
    k = UBound(nameArray)
    ReDim abstractedIndexes(k)
    If chkShuffle.Value = Checked Then
        abstractedIndexes = Shuffle(txtStartNumber, txtStartNumber + k)
    Else
        For i = 0 To k
            abstractedIndexes(i) = txtStartNumber + i
        Next i
    End If

    Name imagesDir As tempImagesDir
    MkDir imagesDir
    Name jsonDir As tempJsonDir
    MkDir jsonDir

    Close
    Set JB = New JsonBag
    JB.Whitespace = frmSetting.chkWhiteSpace.Value = Checked
    JB.WhitespaceIndent = 2
    JB.DecimalMode = False
    For i = 0 To k
        DoEvents
        If IsUnloading Then Exit Sub
        showTips Language.Item("Tips24") & i + 1 & "/" & k + 1
        fn = FreeFile
        Open tempJsonDir & "\" & nameArray(i) & ".json" For Input As #fn
        JB.Clear
        JB.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
        Close #fn

        If i = 0 Then IsSolana = JB.Exists("properties")
        extensionName = GetExtensionName(JB.Item("image"))
        Name tempImagesDir & "\" & nameArray(i) & ".png" As imagesDir & "\" & abstractedIndexes(i) & extensionName
        With JB
            .Item("name") = Split(.Item("name"), "#")(0) & "#" & abstractedIndexes(i)
            .Item("image") = ParsePath(.Item("image")) & abstractedIndexes(i) & extensionName
            If IsSolana Then .Item("properties").Item("files")(1).Item("uri") = .Item("image")
        End With
        fn = FreeFile
        Open jsonDir & "\" & abstractedIndexes(i) & ".json" For Output As #fn
        Print #fn, JB.JSON
        Close #fn
    Next i
    
    If Dir(tempImagesDir, vbDirectory) <> "" Then
        DoEvents
        If IsUnloading Then Exit Sub
        If Dir(tempImagesDir & "\*.*") <> "" Then Kill tempImagesDir & "\*.*"
        RmDir tempImagesDir
    End If
    If Dir(tempJsonDir, vbDirectory) <> "" Then
        DoEvents
        If IsUnloading Then Exit Sub
        If Dir(tempJsonDir & "\*.*") <> "" Then Kill tempJsonDir & "\*.*"
        RmDir tempJsonDir
    End If
    showTips Language.Item("Tips26") & "  " & k + 1 & "/" & k + 1
End Sub

'Status bar information
Private Sub showTips(ByVal Str As String)
    picTips.Cls
    picTips.CurrentX = 0
    picTips.CurrentY = (picTips.ScaleHeight - picTips.TextHeight(Str)) / 2
    picTips.Print Space(2) & Str
End Sub

'safe exit
Private Sub cmdExit_Click()
    Dim objForm As Form
    'unload all forms except this one
    For Each objForm In Forms
        If objForm.hwnd <> Me.hwnd Then  'only the hWnd property is guaranteed to be unique
            Unload objForm
            Set objForm = Nothing
        End If
    Next objForm
    'unload this form
    Unload Me
    End
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    cmdExit_Click
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Call TerminateGDIPlus
    'in Form_Unload
    IsUnloading = True
End Sub
