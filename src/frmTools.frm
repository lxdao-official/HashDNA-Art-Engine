VERSION 5.00
Begin VB.Form frmTools 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tools"
   ClientHeight    =   3450
   ClientLeft      =   5175
   ClientTop       =   4755
   ClientWidth     =   14100
   Icon            =   "frmTools.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   14100
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picTips 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      FontTransparent =   0   'False
      Height          =   375
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   14040
      TabIndex        =   23
      Top             =   3075
      Width           =   14100
   End
   Begin VB.Frame Frame3 
      Caption         =   "NFT Name #Start"
      Height          =   2295
      Left            =   3720
      TabIndex        =   20
      Top             =   360
      Width           =   2175
      Begin VB.CommandButton cmdFixNameNumber 
         Caption         =   "Modify"
         Height          =   495
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.TextBox txtStartNumber 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1080
         TabIndex        =   3
         Text            =   "1"
         ToolTipText     =   "Must be a number. After N attempts, if the unique DNA is still not obtained, the attempt is stopped."
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "From #"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Resize Images"
      Height          =   2295
      Left            =   6240
      TabIndex        =   19
      Top             =   360
      Width           =   2175
      Begin VB.CommandButton cmdSetting 
         Caption         =   "Setting"
         Height          =   495
         Index           =   1
         Left            =   480
         TabIndex        =   24
         Top             =   600
         Width           =   1215
      End
      Begin VB.CommandButton cmdResize 
         Caption         =   "Resize"
         Height          =   495
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Update Metadata"
      Height          =   2295
      Left            =   360
      TabIndex        =   18
      Top             =   360
      Width           =   3015
      Begin VB.CommandButton cmdSetting 
         Caption         =   "Setting"
         Height          =   495
         Index           =   0
         Left            =   840
         TabIndex        =   22
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton Option2 
         Caption         =   "All"
         Height          =   375
         Index           =   1
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptionUpdate 
         Caption         =   "BaseURL"
         Height          =   375
         Index           =   0
         Left            =   480
         TabIndex        =   0
         Top             =   360
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.CommandButton cmdUpdateMeta 
         Caption         =   "Update"
         Height          =   495
         Left            =   840
         TabIndex        =   2
         Top             =   1440
         Width           =   1215
      End
   End
   Begin VB.Frame FrameOthers 
      Caption         =   "Signature (building...)"
      Height          =   2295
      Left            =   8760
      TabIndex        =   17
      Top             =   360
      Width           =   4935
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   11
         Text            =   "20"
         ToolTipText     =   "The background color must be a 6-character(RGB) or 8-character(ARGB) hexadecimal without a pre-pended #"
         Top             =   1080
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "I"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1080
         Width           =   375
      End
      Begin VB.CheckBox Check2 
         Caption         =   "B"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1560
         Width           =   375
      End
      Begin VB.TextBox txtStaticColor 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   1560
         TabIndex        =   12
         Text            =   "FF000000"
         ToolTipText     =   "The background color must be a 6-character(RGB) or 8-character(ARGB) hexadecimal without a pre-pended #"
         Top             =   1560
         Width           =   1095
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Number"
         Height          =   495
         Left            =   3360
         TabIndex        =   15
         Top             =   960
         Width           =   1455
      End
      Begin VB.OptionButton OptSign 
         Height          =   255
         Index           =   3
         Left            =   960
         TabIndex        =   7
         ToolTipText     =   "In the lower-right corner of the picture."
         Top             =   1560
         Value           =   -1  'True
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   8
         Top             =   1560
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   1
         Left            =   960
         TabIndex        =   10
         Top             =   1200
         Width           =   255
      End
      Begin VB.OptionButton Option1 
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   9
         Top             =   1200
         Width           =   255
      End
      Begin VB.TextBox txtSign 
         Height          =   375
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   4095
      End
      Begin VB.CommandButton cmdSign 
         Caption         =   "Sign"
         Enabled         =   0   'False
         Height          =   495
         Left            =   3360
         TabIndex        =   16
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   855
         Left            =   480
         Top             =   1080
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub cmdSetting_Click(Index As Integer)
    frmSetting.Show 1
End Sub

'Update metadata json files information
Private Sub cmdUpdateMeta_Click()
    Dim jsonDir As String
    jsonDir = buildDir & "\json"
    If Dir(jsonDir, vbDirectory) = "" Then
        showTips "The json folder was not found."
        Exit Sub
    End If
    If OptionUpdate(0).Value = True Then
        updateBaseURL jsonDir
    Else
        updateAll jsonDir
    End If
End Sub

'Update the baseURL of the images (the URL or ipfs CID of the images folder)
Private Sub updateBaseURL(jsonDir As String)
    Dim k As Long, fn As Integer
    Dim tempName As String
    Dim IsSolana As Boolean
    'Public.bas Sub
    InitJB
    '读json文件下所有.json文件
    k = 0
    tempName = Dir(jsonDir & "\")
    Do While tempName <> ""
        On Error GoTo nextJson
        If LCase(Right(tempName, 5)) = ".json" Then
            DoEvents
            showTips "Updating..." & tempName
            
            fn = FreeFile
            Open jsonDir & "\" & tempName For Input As #fn
            JB.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
            Close #fn
            If k = 0 Then IsSolana = JB.Exists("properties")
            If JB.Exists("image") Then JB.Item("image") = frmSetting.txtImageBaseURL & ParseFileName(JB.Item("image"))
            If IsSolana Then JB.Item("properties").Item("files")(1).Item("uri") = JB.Item("image")
            
            fn = FreeFile
            Open jsonDir & "\" & tempName For Output As #fn
            Print #fn, JB.JSON
            Close #fn
            
            k = k + 1
        End If
        tempName = Dir()
nextJson:
    Loop
    showTips "Great! The update is complete. Total number of JSON files: " & k
    If k = 0 Then
        showTips "The json file was not found."
    End If
End Sub

'Update the metadata information in json files according to the setting content.
Private Sub updateAll(jsonDir As String)
    Dim i As Long, j As Long, k As Long, fn As Integer
    Dim tempName As String
    Dim tempJB As JsonBag
    Dim IsSolana As Boolean
    'Publick.bas Sub, read the information from the settings and initialize the Metedata JSON template.
    GetTemplateJB
    Set tempJB = New JsonBag
    tempJB.Whitespace = frmSetting.chkWhiteSpace.Value = Checked
    tempJB.WhitespaceIndent = 2
    tempJB.DecimalMode = False
    IsSolana = frmSetting.OptionNetwork(1).Value
    'Read all .json files under the build\json file
    k = 0
    tempName = Dir(jsonDir & "\")
    Do While tempName <> ""
        On Error GoTo nextJson
        If LCase(Right(tempName, 5)) = ".json" Then
            DoEvents
            showTips "Updating..." & tempName
        
            fn = FreeFile
            Open jsonDir & "\" & tempName For Input As #fn
            tempJB.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
            Close #fn
            With JB
                If tempJB.Exists("name") Then .Item("name") = frmSetting.txtNamePrefix.Text & " #" & Split(tempJB.Item("name"), "#")(1)
                If tempJB.Exists("image") Then .Item("image") = frmSetting.txtImageBaseURL & ParseFileName(tempJB.Item("image"))
                If tempJB.Exists("attributes") Then .Item("attributes") = tempJB.Item("attributes")
                If IsSolana Then .Item("properties").Item("files")(1).Item("uri") = .Item("image")
            End With

            fn = FreeFile
            Open jsonDir & "\" & tempName For Output As #fn
            Print #fn, JB.JSON
            Close #fn
            
            k = k + 1
        End If
        tempName = Dir()
nextJson:
    Loop
    showTips "Great! The update is complete. Total number of JSON files: " & k
    If k = 0 Then
        showTips "The json file was not found."
    End If
End Sub

'In Solana network, it needs to start from 0.json and 0.png, but you can make the metadata NFT name in the json file start from #1.
Private Sub cmdFixNameNumber_Click()
    Dim jsonDir As String
    Dim i As Long, j As Long, k As Long
    Dim nameArray() As Long
    Dim tempName As String
    Dim fn As Integer
    Dim difference As Long
    
    jsonDir = buildDir & "\json"
    If Dir(jsonDir, vbDirectory) = "" Then
        showTips "The json folder was not found."
        Exit Sub
    End If
    'Public.bas Sub
    InitJB
    'Read all json files in the json folder and put the number part of the name into an array.
    k = 0
    tempName = Dir(jsonDir & "\")
    Do While tempName <> ""
        If LCase(Right(tempName, 5)) = ".json" Then
            If IsNumeric(Left(tempName, Len(tempName) - 5)) Then
                ReDim Preserve nameArray(k)
                nameArray(k) = Val(Left(tempName, Len(tempName) - 5))
                k = k + 1
            End If
        End If
        tempName = Dir()
    Loop
    If k = 0 Then
        showTips "The json file was not found."
        Exit Sub
    End If
    ''Reorder the array of json filenames from smallest to largest
    ArraySort nameArray
    k = UBound(nameArray)

    Close

    For i = 0 To k
        DoEvents
        On Error GoTo nextJson
        showTips "Updating..." & i + 1 & "/" & k + 1
        fn = FreeFile
        Open jsonDir & "\" & nameArray(i) & ".json" For Input As #fn
        JB.JSON = StrConv(InputB(LOF(fn), fn), vbUnicode)
        Close #fn
        
        If JB.Exists("name") Then
            If i = 0 Then difference = Val(txtStartNumber) - Val(Split(JB.Item("name"), "#")(1))
            JB.Item("name") = Split(JB.Item("name"), "#")(0) & "#" & Val(Split(JB.Item("name"), "#")(1)) + difference
        End If
            
        fn = FreeFile
        Open jsonDir & "\" & nameArray(i) & ".json" For Output As #fn
        Print #fn, JB.JSON
        Close #fn
nextJson:
    Next i
    showTips "Great! The update is complete. Total number of JSON files: " & k + 1
End Sub

'Resize images
Private Sub cmdResize_Click()
    Dim imagesDir As String
    Dim tempName As String
    Dim smoothing As Boolean
    Dim k As Long
    
    imagesDir = buildDir & "\images"
    If Dir(imagesDir, vbDirectory) = "" Then
        showTips "The images folder was not found."
        Exit Sub
    End If
    If frmSetting.chkSmoothing.Value = Checked Then smoothing = True Else smoothing = False
    k = 0
    tempName = Dir(imagesDir & "\")
    Do While tempName <> ""
        If LCase(Right(tempName, 4)) = ".png" Then
            DoEvents
            showTips "Resizing... " & tempName
            'call the public function Resize() in the Public.bas
            If Resize(imagesDir & "\" & tempName, Val(frmSetting.txtWidth), Val(frmSetting.txtHeight), smoothing) = True Then k = k + 1
        End If
        tempName = Dir()
    Loop
    If k = 0 Then
        showTips "The image file was not found."
    Else
        showTips "Great! Resizing is complete. Total number of image files: " & k
    End If
End Sub

Private Sub showTips(Str As String)
    picTips.Cls
    picTips.CurrentX = 0
    picTips.CurrentY = (picTips.ScaleHeight - picTips.TextHeight(Str)) / 2
    picTips.Print Space(2) & Str
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = False
End Sub
