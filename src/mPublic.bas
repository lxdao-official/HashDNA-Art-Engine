Attribute VB_Name = "Public"
Option Explicit

Type element
    id As Long
    Name As String
    fileName As String
    path As String
    weight As Long
    usableMax As Long
    usedCount As Long
End Type

Type layer
    id As Long
    Name As String
    bypassDNA As Boolean
    totalWeight As Long
    elements() As element
End Type

Type layerConfig
    typeSize As Long
    typeName As String
    layersOrder() As String
    layersSize As Long
End Type

Public basePath As String
Public buildDir As String
Public layersDir As String
Public JB As JsonBag
Public Language As JsonBag
Public Const Compiler As String = "HashDNA Art Engine"
Public Const ENS As String = "HashDNA.eth"
Public Const WalletAddress As String = "0x033d845894Fbed101BaE4916Ec8F872B5a5b06cd"

'Build BUILD, IMAGES and JSON folder in app.path.
Public Sub BuildSetup()
    On Error Resume Next
    'Determine if the app is in the root folder or in the normal directory
    If Right(App.path, 1) = "\" Then basePath = Left(App.path, Len(App.path) - 1) Else basePath = App.path
    buildDir = basePath & "\build"
    If Dir(buildDir, vbDirectory) = "" Then MkDir buildDir
    If Dir(buildDir & "\images", vbDirectory) = "" Then MkDir buildDir & "\images"
    If Dir(buildDir & "\json", vbDirectory) = "" Then MkDir buildDir & "\json"
    On Error GoTo 0
End Sub

'initialize the metedata JSON JB.
Public Sub InitJB()
    Set JB = New JsonBag
    JB.Whitespace = frmSetting.chkWhiteSpace.Value = Checked
    JB.WhitespaceIndent = 2
    JB.DecimalMode = False
End Sub

'Get the Metadata information from the settings and create a Metadata template (JsonBag class).
Public Sub GetTemplateJB()
    Dim i As Integer
    InitJB
    If frmSetting.OptionNetwork(0).Value = True Then
        With JB
            .Clear
            .IsArray = False    'Actually the default after Clear.
            .Item("name") = frmSetting.txtNamePrefix.Text
            If frmSetting.txtDescription <> "" Then .Item("description") = frmSetting.txtDescription.Text
            .Item("image") = frmSetting.txtImageBaseURL.Text
            If frmSetting.txtExternal_url <> "" Then .Item("external_url") = frmSetting.txtExternal_url.Text
            If frmSetting.txtAnimation_url <> "" Then .Item("animation_url") = frmSetting.txtAnimation_url.Text
            For i = 0 To 2
                If frmSetting.txtExtra(i) <> "" And frmSetting.txtExtraValue(i) <> "" Then
                    If IsNumeric(frmSetting.txtExtraValue(i)) Then
                        .Item(frmSetting.txtExtra(i)) = Val(frmSetting.txtExtraValue(i).Text)
                    Else
                        .Item(frmSetting.txtExtra(i)) = frmSetting.txtExtraValue(i).Text
                    End If
                End If
            Next i
            .AddNewArray ("attributes")
            .Item("compiler") = Compiler
        End With
    Else
        With JB
            .Clear
            .IsArray = False
            .Item("name") = frmSetting.txtNamePrefix.Text
            .Item("symbol") = frmSetting.txtSolSymbol.Text
            .Item("description") = frmSetting.txtDescription.Text
            If frmSetting.txtSolFee <> "" Then .Item("seller_fee_basis_points") = Val(frmSetting.txtSolFee.Text)
            .Item("image") = frmSetting.txtImageBaseURL.Text
            If frmSetting.txtExternal_url <> "" Then .Item("external_url") = frmSetting.txtExternal_url.Text
            If frmSetting.txtAnimation_url <> "" Then .Item("animation_url") = frmSetting.txtAnimation_url.Text
            For i = 0 To 2
                If frmSetting.txtExtra(i) <> "" And frmSetting.txtExtraValue(i) <> "" Then
                    If IsNumeric(frmSetting.txtExtraValue(i)) Then
                        .Item(frmSetting.txtExtra(i)) = Val(frmSetting.txtExtraValue(i).Text)
                    Else
                        .Item(frmSetting.txtExtra(i)) = frmSetting.txtExtraValue(i).Text
                    End If
                End If
            Next i
            .AddNewArray ("attributes")
            With .AddNewObject("properties")
                With .AddNewArray("files")
                    With .AddNewObject()
                        .Item("uri") = "need replace"
                        .Item("type") = "image/png"
                    End With
                End With
                .Item("category") = "image"
                If frmSetting.txtSolFee <> "" Then
                    With .AddNewArray("creators")
                        With .AddNewObject()
                            .Item("address") = frmSetting.txtSolCreatorsAddress.Text
                            .Item("share") = Val(frmSetting.txtSolCreatorsShare.Text)
                        End With
                    End With
                End If
            End With
            .Item("compiler") = Compiler
        End With
    End If
End Sub

'HSL format color to RGB format color
Public Function HSL2RGB(h As Integer, s As Integer, l As Integer) As String
    Dim i As Integer
    Dim r As Integer, g As Integer, b As Integer
    Dim rc As Single, Gc As Single, Bc As Single
    Dim Hk As Single, Sc As Single, Lc As Single
    Dim p, q
    Dim tRGB() As Single, RGBc() As Single
    ReDim tRGB(3)
    ReDim RGBc(3)
    Hk = h / 360: Sc = s / 100: Lc = l / 100
    If Sc = 0 Then
        rc = Lc: Gc = Lc: Bc = Lc
    Else
        If Lc < 0.5 Then
            q = Lc * (1 + Sc)
        Else
            q = Lc + Sc - (Lc * Sc)
        End If
        p = 2 * Lc - q
        tRGB(1) = Hk + 1 / 3
        tRGB(2) = Hk
        tRGB(3) = Hk - 1 / 3
        For i = 1 To 3
            If tRGB(i) < 0 Then tRGB(i) = tRGB(i) + 1
            If tRGB(i) > 1 Then tRGB(i) = tRGB(i) - 1
        Next i
        For i = 1 To 3
            If tRGB(i) < (1 / 6) Then
                RGBc(i) = p + ((q - p) * 6 * tRGB(i))
            ElseIf tRGB(i) >= (1 / 6) And tRGB(i) < 0.5 Then
                RGBc(i) = q
            ElseIf tRGB(i) >= 0.5 And tRGB(i) < (2 / 3) Then
                RGBc(i) = p + ((q - p) * 6 * (2 / 3 - tRGB(i)))
            Else
                RGBc(i) = p
            End If
        Next i
        rc = RGBc(1): Gc = RGBc(2): Bc = RGBc(3)
    End If
    r = Round(rc * 255)
    g = Round(Gc * 255)
    b = Round(Bc * 255)
    HSL2RGB = r & "," & g & "," & b
End Function

'Sort an array
Public Function ArraySort(ByRef a, Optional UP As Boolean = True) As Boolean
    Dim i As Long, j As Long, temp
    On Error GoTo ErrHandler
    If UP = True Then
        For i = LBound(a) To UBound(a) - 1
            For j = i + 1 To UBound(a)
                If a(i) > a(j) Then
                    temp = a(i)
                    a(i) = a(j)
                    a(j) = temp
                End If
            Next j
        Next i
    Else
        For i = LBound(a) To UBound(a) - 1
            For j = i + 1 To UBound(a)
                If a(i) < a(j) Then
                    temp = a(i)
                    a(i) = a(j)
                    a(j) = temp
                End If
            Next j
        Next i
    End If
    ArraySort = True
    Exit Function
ErrHandler:
    ArraySort = False
End Function

'Returns a array sorted randomly from min to max.
Public Function Shuffle(min As Long, max As Long) As Long()
    Dim i As Long, j As Long, tmp As Long
    Dim X() As Long
    ReDim X(max - min)
    For i = 0 To max - min
        X(i) = min + i
    Next
    Randomize
    For i = max To min Step -1
        j = Int(Rnd * (i - min)) + min
        tmp = X(j - min)
        X(j - min) = X(i - min)
        X(i - min) = tmp
    Next
    Shuffle = X
End Function

'Resize images
Public Function Resize(imagePath As String, Optional saveWidth As Long = 0&, Optional saveHeight As Long = 0&, Optional SmoothingMode As Boolean = True) As Boolean
    Dim graphics As Long
    Dim bitmap As Long
    Dim Image As Long
    Dim imgWidth As Long
    Dim imgHeight As Long
    
    If GdipLoadImageFromFile(StrPtr(imagePath), Image) <> Ok Then
        Resize = False
        Exit Function
    Else
        Resize = True
    End If
    GdipGetImageWidth Image, imgWidth
    GdipGetImageHeight Image, imgHeight
    
    'If a width and height value has been specified, it will be used, and if it is not specified,
    'it will be superimposed according to the size of the original image.
    If saveWidth > 1 And saveHeight > 1 Then
        imgWidth = saveWidth
        imgHeight = saveHeight
    ElseIf saveWidth > 1 And saveHeight = 0 Then
        saveHeight = saveWidth * imgHeight / imgWidth
        imgWidth = saveWidth
        imgHeight = saveHeight
    ElseIf saveWidth = 0 And saveHeight > 1 Then
        saveWidth = saveHeight * imgWidth / imgHeight
        imgWidth = saveWidth
        imgHeight = saveHeight
    End If
    
    CreateBitmapWithGraphics bitmap, graphics, imgWidth, imgHeight
    If SmoothingMode = True Then GdipSetSmoothingMode graphics, SmoothingModeAntiAlias
    
    GdipDrawImageRect graphics, Image, 0, 0, imgWidth, imgHeight
    GdipDisposeImage Image
    SaveImageToPNG bitmap, imagePath

    GdipDeleteGraphics graphics
    GdipDisposeImage bitmap
End Function

'Parse the path from a string (includ / or \)
Public Function ParsePath(sPath As String) As String
    Dim i As Integer
    For i = Len(sPath) To 1 Step -1
        If InStr(":\", Mid$(sPath, i, 1)) Or InStr(":/", Mid$(sPath, i, 1)) Then Exit For
    Next
    ParsePath = Left$(sPath, i)
End Function

'Parse the file name from a string (include extension)
Public Function ParseFileName(sFileIn As String) As String
    Dim i As Integer
    For i = Len(sFileIn) To 1 Step -1
        If InStr("\", Mid$(sFileIn, i, 1)) Or InStr("/", Mid$(sFileIn, i, 1)) Then Exit For
    Next
    ParseFileName = Mid$(sFileIn, i + 1, Len(sFileIn) - i)
End Function

'Parse the prefix from the file name (remove the extension)
Public Function GetFileName(ByVal fileName As String) As String
      Dim DotIndex As Long
      DotIndex = InStrRev(fileName, ".")
      If DotIndex = 0 Then
        GetFileName = fileName
      Else
        GetFileName = Left(fileName, DotIndex - 1)
      End If
End Function

'Parse the extension from the file name (eg. .png .txt ....)
Public Function GetExtensionName(ByVal fileName As String) As String
    Dim DotIndex As Long
    DotIndex = InStrRev(fileName, ".")
    If DotIndex = 0 Or (Len(fileName) - DotIndex) > 6 Then
        GetExtensionName = ""
    Else
        GetExtensionName = Right(fileName, Len(fileName) - DotIndex + 1)
    End If
End Function

'Export the default language information based on tag information
'including Caption, ToolTipText, and Text
Public Sub SetLanguageTemplate()
    Dim languagesJB As JsonBag, fn As Integer, i As Integer
    Dim frm As Form, o As Object
    Dim sMainKey As String, sKey As String, sCaption As String, sToolTipText As String, sText As String
    Dim ctrlIndex As Integer
    Load frmMain
    Load frmAbout
    Load frmSetting
    Load frmTools

    Set languagesJB = New JsonBag
    languagesJB.Whitespace = True
    With languagesJB
        .Clear
        .IsArray = False
        For i = 0 To Forms.Count - 1
            sMainKey = Forms(i).Name
            .Item(sMainKey & ".Caption") = Forms(i).Caption
            For Each o In Forms(i).Controls
                sKey = o.Name
                On Error Resume Next
                ctrlIndex = o.Index
                If Err.Number = 0 Then sKey = sKey & "-" & ctrlIndex
                Err.Clear
                If Mid(o.Tag, 1, 1) = 1 Then .Item(sMainKey & "." & sKey & ".Caption") = o.Caption
                If Mid(o.Tag, 2, 1) = 1 Then .Item(sMainKey & "." & sKey & ".ToolTipText") = o.ToolTipText
                If Mid(o.Tag, 3, 1) = 1 And typeName(o) = "TextBox" Then .Item(sMainKey & "." & sKey & ".Text") = o.Text
            Next o
        Next i
    End With
    fn = FreeFile
    Open basePath & "\Default.json" For Output As #fn
    Print #fn, languagesJB.JSON
    Close #fn
End Sub

'Get the language list from the resource file and put it in the languages listbox
Public Sub GetLanguagesList()
    With frmMain.cboLanguages
        .AddItem "English"
        .ItemData(.newIndex) = 1033
        .AddItem "Simplified Chinese"
        .ItemData(.newIndex) = 2052
        .AddItem "Traditional Chinese"
        .ItemData(.newIndex) = 1028
        .ListIndex = 0
    End With
End Sub

'Translate a form
Public Sub TranslateForm(oForm As Form)
  Dim oCNTRL As control
  ' The collection holding the Key, Text pairs
  If Language Is Nothing Then Exit Sub
  On Error Resume Next
  ' The caption of the form
  oForm.Caption = Language.Item(oForm.Name & ".caption")
  ' Get the caption and tooltiptext for the controls
  For Each oCNTRL In oForm.Controls
    TranslateControl oCNTRL, oForm.Name
  Next
End Sub

'Translate the properties of a control based on the language information in the resource file,
'including Caption, ToolTipText, and Text
Public Sub TranslateControl(oCNTRL As control, sMainKey As String)
  Dim sKey As String, sValue As String
  Dim ctrlIndex As Integer
  ' The collection holding the Key, Text pairs
  If Language Is Nothing Then Exit Sub
  On Error Resume Next
  sKey = oCNTRL.Name
  ctrlIndex = oCNTRL.Index
  If Err.Number = 0 Then sKey = sKey & "-" & ctrlIndex
  ' The Caption
  sValue = ""
  sValue = Language.Item(sMainKey & "." & sKey & ".Caption")
  If Len(sValue) > 0 Then oCNTRL.Caption = sValue
  ' The ToolTipText
  sValue = ""
  sValue = Language.Item(sMainKey & "." & sKey & ".ToolTipText")
  If Len(sValue) > 0 Then oCNTRL.ToolTipText = sValue
    ' The Text
  sValue = ""
  sValue = Language.Item(sMainKey & "." & sKey & ".Text")
  If Len(sValue) > 0 Then oCNTRL.Text = sValue
End Sub
