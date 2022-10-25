Attribute VB_Name = "Gdip"
Option Explicit

'vIstaswx GDI+ ����ģ��
'vIstaswx GDI+ Declare Module

'vIstaswx ������չ
'Extended by vIstaswx

'===========================================
'����޸ģ�2011/2/8
'Latest edit: 2011/2/8
'
'2011-2-8
'1.����Gdi+1.1�ĺ���,�ṹ��,ö�ٺͳ���������
'2.����GdipSetImageAttributesCachedBackground
'  ��GdipTestControl��������
'3.�޸�InitGdiPlus(To)�Ĳ���
'4.����һЩbug
'5.��ʽ����API�����ͽṹ��ʹ֮���׶�
'6.Enum ImageType -> Enum GdipImageType
'7.���� NewPointF,NewPointL,NewPointsF,NewPointsL,NewColors ����
'8.���� Zero(Point/Rect)(F/L) 0����
'
'2011-2-7
'1.����GdipSetLinePresetBlend��4���������������Ĵ���
'
'2010-6-5:
'1.����ͼƬ�����Ż�
'2.InitGDIPlus(To) ����ʱ��ѡ��ʾ����Ի����˳�����
'  ֧���Զ������Ի������ݣ����ӷ���ֵ�������Ѿ���ʼ�����ж�
'3.TerminateGDIPlus(From) �����Ѿ��رյ��ж�
'4.ɾ��RtlMoveMemory(CopyMemory)�������޸�CLSIDFromString����ΪPrivate��
'===========================================

'http://vIstaswx..com
'QQ     : 490241327

#Const GdipVersion = 1#

'===================================================================================
'  ��������
'===================================================================================

'=================================
'== Structures                  ==
'=================================

'=================================
'Point Structure
Public Type POINTL
    X As Long
    Y As Long
End Type

Public Type POINTF
    X As Single
    Y As Single
End Type

'=================================
'Rectange Structure
Public Type RECTL
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Public Type RECTF
    Left As Single
    Top As Single
    Right As Single
    Bottom As Single
End Type

'=================================
'Size Structure
Public Type SIZEL
    cx As Long
    cy As Long
End Type

Public Type SIZEF
    cx As Single
    cy As Single
End Type

'=================================
'Bitmap Structure
Public Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Public Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Public Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Public Type BitmapData
    Width As Long
    Height As Long
    stride As Long
    PixelFormat As GpPixelFormat
    scan0 As Long
    Reserved As Long
End Type

'=================================
'Color Structure
Public Type COLORBYTES
    BlueByte As Byte
    GreenByte As Byte
    RedByte As Byte
    AlphaByte As Byte
End Type

Public Type COLORLONG
    longval As Long
End Type

Public Type ColorMap
    oldColor As Long
    newColor As Long
End Type

Public Type ColorMatrix
    m(0 To 4, 0 To 4) As Single
End Type

'=================================
'Path
Public Type PathData
    Count As Long
    pPoints As Long
    pTypes As Long
End Type

'=================================
'Encoder
Public Type CLSID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Public Type EncoderParameter
    GUID As CLSID
    NumberOfValues As Long

Type As EncoderParameterValueType
    Value As Long
End Type

Public Type EncoderParameters
    Count As Long
    Parameter As EncoderParameter
End Type

'=================================
'== Enums                       ==
'=================================

'=================================
'Pixel
Public Enum GpPixelFormat
    PixelFormat1bppIndexed = &H30101
    PixelFormat4bppIndexed = &H30402
    PixelFormat8bppIndexed = &H30803
    PixelFormat16bppGreyScale = &H101004
    PixelFormat16bppRGB555 = &H21005
    PixelFormat16bppRGB565 = &H21006
    PixelFormat16bppARGB1555 = &H61007
    PixelFormat24bpprgb = &H21808
    PixelFormat32bppRGB = &H22009
    PixelFormat32bppARGB = &H26200A
    PixelFormat32bppPARGB = &HE200B
    PixelFormat48bppRGB = &H10300C
    PixelFormat64bppARGB = &H34400D
    PixelFormat64bppPARGB = &H1C400E
End Enum

'=================================
'Unit
Public Enum GpUnit
    UnitWorld = 0
    UnitDisplay
    UnitPixel
    UnitPoint
    UnitInch
    UnitDocument
    UnitMillimeter
End Enum

'=================================
'Path
Public Enum PathPointType  'GdipGetPathTypes,GdipCreatePath2,GdipCreatePath2I
    PathPointTypeStart = 0
    PathPointTypeLine = 1
    PathPointTypeBezier = 3
    PathPointtypeDirTypeMask = &H7
    PathPointtypeDirDashMode = &H10
    PathPointtypeDirMarker = &H20
    PathPointTypeCloseSubpath = &H80
    PathPointTypeBezier3 = 3
End Enum

'=================================
'Font / String
Public Enum GenericFontFamily
    GenericFontFamilySerif = 0
    GenericFontFamilySansSerif
    GenericFontFamilyMonospace
End Enum

Public Enum FontStyle
    FontStyleRegular = 0
    FontStyleBold = 1
    FontStyleItalic = 2
    FontStyleBoldItalic = 3
    FontStyleUnderline = 4
    FontStyleStrikeout = 8
End Enum

Public Enum StringAlignment
    StringAlignmentNear = 0
    StringAlignmentCenter = 1
    StringAlignmentFar = 2
End Enum

'=================================
'Fill / Wrap
Public Enum FillMode
    FillModeAlternate = 0
    FillModeWinding
End Enum

Public Enum WrapMode
    WrapModeTile = 0
    WrapModeTileFlipX
    WrapModeTileFlipY
    WrapModeTileFlipXY
    WrapModeClamp
End Enum

Public Enum LinearGradientMode
    LinearGradientModeHorizontal = 0
    LinearGradientModeVertical
    LinearGradientModeForwardDiagonal
    LinearGradientModeBackwardDiagonal
End Enum

'=================================
'Quality
Public Enum QualityMode
    QualityModeInvalid = -1
    QualityModeDefault = 0
    QualityModeLow = 1
    QualityModeHigh = 2
End Enum

Public Enum CompositingMode
    CompositingModeSourceOver = 0
    CompositingModeSourceCopy
End Enum

Public Enum CompositingQuality
    CompositingQualityInvalid = QualityModeInvalid
    CompositingQualityDefault = QualityModeDefault
    CompositingQualityHighSpeed = QualityModeLow
    CompositingQualityHighQuality = QualityModeHigh
    CompositingQualityGammaCorrected
    CompositingQualityAssumeLinear
End Enum

Public Enum SmoothingMode
    SmoothingModeInvalid = QualityModeInvalid
    SmoothingModeDefault = QualityModeDefault
    SmoothingModeHighSpeed = QualityModeLow
    SmoothingModeHighQuality = QualityModeHigh
    SmoothingModeNone
    SmoothingModeAntiAlias
    #If GdipVersion >= 1.1 Then
    SmoothingModeAntiAlias8x4 = SmoothingModeAntiAlias
    SmoothingModeAntiAlias8x8
    #End If
End Enum

Public Enum InterpolationMode
    InterpolationModeInvalid = QualityModeInvalid
    InterpolationModeDefault = QualityModeDefault
    InterpolationModeLowQuality = QualityModeLow
    InterpolationModeHighQuality = QualityModeHigh
    InterpolationModeBilinear
    InterpolationModeBicubic
    InterpolationModeNearestNeighbor
    InterpolationModeHighQualityBilinear
    InterpolationModeHighQualityBicubic
End Enum

Public Enum PixelOffsetMode
    PixelOffsetModeInvalid = QualityModeInvalid
    PixelOffsetModeDefault = QualityModeDefault
    PixelOffsetModeHighSpeed = QualityModeLow
    PixelOffsetModeHighQuality = QualityModeHigh
    PixelOffsetModeNone
    PixelOffsetModeHalf
End Enum

Public Enum TextRenderingHint
    TextRenderingHintSystemDefault = 0            ' Glyph with system default rendering hint
    TextRenderingHintSingleBitPerPixelGridFit     ' Glyph bitmap with hinting
    TextRenderingHintSingleBitPerPixel            ' Glyph bitmap without hinting
    TextRenderingHintAntiAliasGridFit             ' Glyph anti-alias bitmap with hinting
    TextRenderingHintAntiAlias                    ' Glyph anti-alias bitmap without hinting
    TextRenderingHintClearTypeGridFit             ' Glyph CT bitmap with hinting
End Enum

'=================================
'Color Matrix
Public Enum MatrixOrder
    MatrixOrderPrepend = 0
    MatrixOrderAppend = 1
End Enum

Public Enum ColorAdjustType
    ColorAdjustTypeDefault = 0
    ColorAdjustTypeBitmap
    ColorAdjustTypeBrush
    ColorAdjustTypePen
    ColorAdjustTypeText
    ColorAdjustTypeCount
    ColorAdjustTypeAny
End Enum

Public Enum ColorMatrixFlags
    ColorMatrixFlagsDefault = 0
    ColorMatrixFlagsSkipGrays = 1
    ColorMatrixFlagsAltGray = 2
End Enum

Public Enum WarpMode
    WarpModePerspective = 0
    WarpModeBilinear
End Enum

Public Enum CombineMode
    CombineModeReplace = 0
    CombineModeIntersect
    CombineModeUnion
    CombineModeXor
    CombineModeExclude
    CombineModeComplement
End Enum

Public Enum ImageLockMode
    ImageLockModeRead = 1
    ImageLockModeWrite = 2
    ImageLockModeUserInputBuf = 4
End Enum

Public Declare Function GdipGetDC _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        hDC As Long) As GpStatus
Attribute GdipGetDC.VB_UserMemId = 1879048192
Public Declare Function GdipReleaseDC _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal hDC As Long) As GpStatus
Attribute GdipReleaseDC.VB_UserMemId = 1879048224

'==================================================

Public Declare Function GdipCreateFromHDC _
                         Lib "gdiplus" (ByVal hDC As Long, _
                                        graphics As Long) As GpStatus
Attribute GdipCreateFromHDC.VB_UserMemId = 1879048260
Public Declare Function GdipCreateFromHWND _
                         Lib "gdiplus" (ByVal hwnd As Long, _
                                        graphics As Long) As GpStatus
Attribute GdipCreateFromHWND.VB_UserMemId = 1879048300
Public Declare Function GdipGetImageGraphicsContext _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        graphics As Long) As GpStatus
Attribute GdipGetImageGraphicsContext.VB_UserMemId = 1879048340
Public Declare Function GdipDeleteGraphics _
                         Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Attribute GdipDeleteGraphics.VB_UserMemId = 1879048388

Public Declare Function GdipGraphicsClear _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal lColor As Long) As GpStatus
Attribute GdipGraphicsClear.VB_UserMemId = 1879048428

Public Declare Function GdipSetCompositingMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal CompositingMd As CompositingMode) As GpStatus
Attribute GdipSetCompositingMode.VB_UserMemId = 1879048468
Public Declare Function GdipGetCompositingMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        CompositingMd As CompositingMode) As GpStatus
Attribute GdipGetCompositingMode.VB_UserMemId = 1879048512
Public Declare Function GdipSetRenderingOrigin _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long) As GpStatus
Attribute GdipSetRenderingOrigin.VB_UserMemId = 1879048556
Public Declare Function GdipGetRenderingOrigin _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                      X As Long, _
                                      Y As Long) As GpStatus
Attribute GdipGetRenderingOrigin.VB_UserMemId = 1879048600
Public Declare Function GdipSetCompositingQuality _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal CompositingQlty As CompositingQuality) As GpStatus
Attribute GdipSetCompositingQuality.VB_UserMemId = 1879048644
Public Declare Function GdipGetCompositingQuality _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        CompositingQlty As CompositingQuality) As GpStatus
Attribute GdipGetCompositingQuality.VB_UserMemId = 1879048692
Public Declare Function GdipSetSmoothingMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal SmoothingMd As SmoothingMode) As GpStatus
Attribute GdipSetSmoothingMode.VB_UserMemId = 1879048740
Public Declare Function GdipGetSmoothingMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        SmoothingMd As SmoothingMode) As GpStatus
Attribute GdipGetSmoothingMode.VB_UserMemId = 1879048784
Public Declare Function GdipSetPixelOffsetMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal PixOffsetMode As PixelOffsetMode) As GpStatus
Attribute GdipSetPixelOffsetMode.VB_UserMemId = 1879048828
Public Declare Function GdipGetPixelOffsetMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        PixOffsetMode As PixelOffsetMode) As GpStatus
Attribute GdipGetPixelOffsetMode.VB_UserMemId = 1879048872
Public Declare Function GdipSetTextRenderingHint _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Mode As TextRenderingHint) As GpStatus
Attribute GdipSetTextRenderingHint.VB_UserMemId = 1879048916
Public Declare Function GdipGetTextRenderingHint _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        Mode As TextRenderingHint) As GpStatus
Attribute GdipGetTextRenderingHint.VB_UserMemId = 1879048964
Public Declare Function GdipSetTextContrast _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal contrast As Long) As GpStatus
Attribute GdipSetTextContrast.VB_UserMemId = 1879049012
Public Declare Function GdipGetTextContrast _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        contrast As Long) As GpStatus
Attribute GdipGetTextContrast.VB_UserMemId = 1879049052
Public Declare Function GdipSetInterpolationMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal interpolation As InterpolationMode) As GpStatus
Attribute GdipSetInterpolationMode.VB_UserMemId = 1879049092
Public Declare Function GdipGetInterpolationMode _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        interpolation As InterpolationMode) As GpStatus
Attribute GdipGetInterpolationMode.VB_UserMemId = 1879049140

Public Declare Function GdipSetWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal matrix As Long) As GpStatus
Attribute GdipSetWorldTransform.VB_UserMemId = 1879049188
Public Declare Function GdipResetWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Attribute GdipResetWorldTransform.VB_UserMemId = 1879049232
Public Declare Function GdipMultiplyWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal matrix As Long, _
                                        ByVal order As MatrixOrder) As GpStatus
Attribute GdipMultiplyWorldTransform.VB_UserMemId = 1879049276
Public Declare Function GdipTranslateWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Attribute GdipTranslateWorldTransform.VB_UserMemId = 1879049324
Public Declare Function GdipScaleWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal sx As Single, _
                                        ByVal sy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Attribute GdipScaleWorldTransform.VB_UserMemId = 1879049372
Public Declare Function GdipRotateWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal angle As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Attribute GdipRotateWorldTransform.VB_UserMemId = 1879049416
Public Declare Function GdipGetWorldTransform _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal matrix As Long) As GpStatus
Attribute GdipGetWorldTransform.VB_UserMemId = 1879049464
Public Declare Function GdipResetPageTransform _
                         Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Attribute GdipResetPageTransform.VB_UserMemId = 1879049508
Public Declare Function GdipGetPageUnit _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        unit As GpUnit) As GpStatus
Attribute GdipGetPageUnit.VB_UserMemId = 1879049552
Public Declare Function GdipGetPageScale _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        sScale As Single) As GpStatus
Attribute GdipGetPageScale.VB_UserMemId = 1879049588
Public Declare Function GdipSetPageUnit _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal unit As GpUnit) As GpStatus
Attribute GdipSetPageUnit.VB_UserMemId = 1879049628
Public Declare Function GdipSetPageScale _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal sScale As Single) As GpStatus
Attribute GdipSetPageScale.VB_UserMemId = 1879049664
Public Declare Function GdipGetDpiX _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        dpi As Single) As GpStatus
Attribute GdipGetDpiX.VB_UserMemId = 1879049704
Public Declare Function GdipGetDpiY _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        dpi As Single) As GpStatus
Attribute GdipGetDpiY.VB_UserMemId = 1879049736
Public Declare Function GdipTransformPoints _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal destSpace As CoordinateSpace, _
                                        ByVal srcSpace As CoordinateSpace, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Attribute GdipTransformPoints.VB_UserMemId = 1879049768
Public Declare Function GdipTransformPointsI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal destSpace As CoordinateSpace, _
                                        ByVal srcSpace As CoordinateSpace, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Attribute GdipTransformPointsI.VB_UserMemId = 1879049808
Public Declare Function GdipTransformPoints_ _
                         Lib "gdiplus" _
                             Alias "GdipTransformPoints" _
                             (ByVal graphics As Long, _
                              ByVal destSpace As CoordinateSpace, _
                              ByVal srcSpace As CoordinateSpace, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Attribute GdipTransformPoints_.VB_UserMemId = 1879049852
Public Declare Function GdipTransformPointsI_ _
                         Lib "gdiplus" _
                             Alias "GdipTransformPointsI" _
                             (ByVal graphics As Long, _
                              ByVal destSpace As CoordinateSpace, _
                              ByVal srcSpace As CoordinateSpace, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Attribute GdipTransformPointsI_.VB_UserMemId = 1879049892
Public Declare Function GdipGetNearestColor _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        argb As Long) As GpStatus
Attribute GdipGetNearestColor.VB_UserMemId = 1879049936
Public Declare Function GdipCreateHalftonePalette Lib "gdiplus" () As Long
Attribute GdipCreateHalftonePalette.VB_UserMemId = 1879049976

Public Declare Function GdipSetClipGraphics _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal srcgraphics As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Attribute GdipSetClipGraphics.VB_UserMemId = 1879050024
Public Declare Function GdipSetClipRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Attribute GdipSetClipRect.VB_UserMemId = 1879050064
Public Declare Function GdipSetClipRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Attribute GdipSetClipRectI.VB_UserMemId = 1879050100
Public Declare Function GdipSetClipPath _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal path As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Attribute GdipSetClipPath.VB_UserMemId = 1879050140
Public Declare Function GdipSetClipRegion _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal region As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Attribute GdipSetClipRegion.VB_UserMemId = 1879050176
Public Declare Function GdipSetClipHrgn _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal hRgn As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Attribute GdipSetClipHrgn.VB_UserMemId = 1879050216
Public Declare Function GdipResetClip Lib "gdiplus" (ByVal graphics As Long) As GpStatus
Attribute GdipResetClip.VB_UserMemId = 1879050252

Public Declare Function GdipTranslateClip _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single) As GpStatus
Attribute GdipTranslateClip.VB_UserMemId = 1879050288
Public Declare Function GdipTranslateClipI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal dx As Long, _
                                        ByVal dy As Long) As GpStatus
Attribute GdipTranslateClipI.VB_UserMemId = 1879050328
Public Declare Function GdipGetClip Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal region As Long) As GpStatus
Public Declare Function GdipGetClipBounds _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        rect As RECTF) As GpStatus
Public Declare Function GdipGetClipBoundsI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        rect As RECTL) As GpStatus

Public Declare Function GdipIsClipEmpty _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipGetVisibleClipBounds _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        rect As RECTF) As GpStatus
Public Declare Function GdipGetVisibleClipBoundsI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        rect As RECTL) As GpStatus
Public Declare Function GdipIsVisibleClipEmpty _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        result As Long) As GpStatus

Public Declare Function GdipIsVisiblePoint _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisiblePointI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisibleRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisibleRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        result As Long) As GpStatus

Public Declare Function GdipSaveGraphics _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        state As Long) As GpStatus
Public Declare Function GdipRestoreGraphics _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal state As Long) As GpStatus
Public Declare Function GdipBeginContainer _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        dstRect As RECTF, _
                                        srcRect As RECTF, _
                                        ByVal unit As GpUnit, _
                                        state As Long) As GpStatus
Public Declare Function GdipBeginContainerI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        dstRect As RECTL, _
                                        srcRect As RECTL, _
                                        ByVal unit As GpUnit, _
                                        state As Long) As GpStatus
Public Declare Function GdipBeginContainer2 _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        state As Long) As GpStatus
Public Declare Function GdipEndContainer _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal state As Long) As GpStatus

'==================================================

Public Declare Function GdipDrawLine _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal x1 As Single, _
                                        ByVal y1 As Single, _
                                        ByVal x2 As Single, _
                                        ByVal y2 As Single) As GpStatus
Public Declare Function GdipDrawLineI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal x1 As Long, _
                                        ByVal y1 As Long, _
                                        ByVal x2 As Long, _
                                        ByVal y2 As Long) As GpStatus
Public Declare Function GdipDrawLines _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLinesI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLines_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawLines" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawLinesI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawLinesI" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
'==================================================

Public Declare Function GdipDrawArc _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawArcI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus

'==================================================

Public Declare Function GdipDrawBezier _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal x1 As Single, _
                                        ByVal y1 As Single, _
                                        ByVal x2 As Single, _
                                        ByVal y2 As Single, _
                                        ByVal x3 As Single, _
                                        ByVal y3 As Single, _
                                        ByVal x4 As Single, _
                                        ByVal y4 As Single) As GpStatus
Public Declare Function GdipDrawBezierI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal x1 As Long, _
                                        ByVal y1 As Long, _
                                        ByVal x2 As Long, _
                                        ByVal y2 As Long, _
                                        ByVal x3 As Long, _
                                        ByVal y3 As Long, _
                                        ByVal x4 As Long, _
                                        ByVal y4 As Long) As GpStatus
Public Declare Function GdipDrawBeziers _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziers_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawBeziers" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawBeziersI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawBeziersI" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
'==================================================

Public Declare Function GdipDrawRectangle _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawRectangleI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawRectangles _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        rects As RECTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawRectanglesI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        rects As RECTL, _
                                        ByVal Count As Long) As GpStatus

Public Declare Function GdipFillRectangle _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipFillRectangleI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus
Public Declare Function GdipFillRectangles _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        rects As RECTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipFillRectanglesI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        rects As RECTL, _
                                        ByVal Count As Long) As GpStatus

'==================================================

Public Declare Function GdipDrawEllipse _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawEllipseI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus

Public Declare Function GdipFillEllipse _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipFillEllipseI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus

'==================================================

Public Declare Function GdipDrawPie _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipDrawPieI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus

Public Declare Function GdipFillPie _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipFillPieI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus

'==================================================

Public Declare Function GdipDrawPolygon _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus

Public Declare Function GdipFillPolygon _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2 _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygon_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawPolygon" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawPolygonI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawPolygonI" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus

Public Declare Function GdipFillPolygon_ _
                         Lib "gdiplus" _
                             Alias "GdipFillPolygon" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygonI_ _
                         Lib "gdiplus" _
                             Alias "GdipFillPolygonI" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillPolygon2_ _
                         Lib "gdiplus" _
                             Alias "GdipFillPolygon2" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipFillPolygon2I_ _
                         Lib "gdiplus" _
                             Alias "GdipFillPolygon2I" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus

'==================================================

Public Declare Function GdipDrawPath _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        ByVal path As Long) As GpStatus
Attribute GdipDrawPath.VB_UserMemId = 1879052636

Public Declare Function GdipFillPath _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal path As Long) As GpStatus

'==================================================

Public Declare Function GdipDrawCurve _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurveI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2 _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3 _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal Offset As Long, _
                                        ByVal numberOfSegments As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal Offset As Long, _
                                        ByVal numberOfSegments As Long, _
                                        ByVal tension As Single) As GpStatus

Public Declare Function GdipDrawClosedCurve _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2 _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal pen As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus

Public Declare Function GdipFillClosedCurve _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2 _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single, _
                                        ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single, _
                                        ByVal FillMd As FillMode) As GpStatus

Public Declare Function GdipDrawCurve_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawCurve" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurveI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawCurveI" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawCurve2_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawCurve2" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve2I_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawCurve2I" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawCurve3" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal Offset As Long, _
                              ByVal numberOfSegments As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawCurve3I_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawCurve3I" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal Offset As Long, _
                              ByVal numberOfSegments As Long, _
                              ByVal tension As Single) As GpStatus

Public Declare Function GdipDrawClosedCurve_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawClosedCurve" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurveI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawClosedCurveI" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawClosedCurve2_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawClosedCurve2" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipDrawClosedCurve2I_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawClosedCurve2I" _
                             (ByVal graphics As Long, _
                              ByVal pen As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus

Public Declare Function GdipFillClosedCurve_ _
                         Lib "gdiplus" _
                             Alias "GdipFillClosedCurve" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurveI_ _
                         Lib "gdiplus" _
                             Alias "GdipFillClosedCurveI" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipFillClosedCurve2_ _
                         Lib "gdiplus" _
                             Alias "GdipFillClosedCurve2" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single, _
                              ByVal FillMd As FillMode) As GpStatus
Public Declare Function GdipFillClosedCurve2I_ _
                         Lib "gdiplus" _
                             Alias "GdipFillClosedCurve2I" _
                             (ByVal graphics As Long, _
                              ByVal brush As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single, _
                              ByVal FillMd As FillMode) As GpStatus


'==================================================

Public Declare Function GdipFillRegion _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal brush As Long, _
                                        ByVal region As Long) As GpStatus
Attribute GdipFillRegion.VB_UserMemId = 1879053828

'==================================================

Public Declare Function GdipDrawImage _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single) As GpStatus
Public Declare Function GdipDrawImageI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long) As GpStatus

Public Declare Function GdipDrawImageRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipDrawImageRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus
Public Declare Function GdipDrawImagePoints _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        dstpoints As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        dstpoints As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePoints_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawImagePoints" _
                             (ByVal graphics As Long, _
                              ByVal Image As Long, _
                              dstpoints As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointsI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawImagePointsI" _
                             (ByVal graphics As Long, _
                              ByVal Image As Long, _
                              dstpoints As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipDrawImagePointRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal srcx As Single, _
                                        ByVal srcy As Single, _
                                        ByVal srcwidth As Single, _
                                        ByVal srcheight As Single, _
                                        ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal srcx As Long, _
                                        ByVal srcy As Long, _
                                        ByVal srcwidth As Long, _
                                        ByVal srcheight As Long, _
                                        ByVal srcUnit As GpUnit) As GpStatus
Public Declare Function GdipDrawImagePointsRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal srcx As Single, _
                                        ByVal srcy As Single, _
                                        ByVal srcwidth As Single, _
                                        ByVal srcheight As Single, _
                                        ByVal srcUnit As GpUnit, _
                                        Optional ByVal imageAttributes As Long = 0, _
                                        Optional ByVal callback As Long = 0, _
                                        Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal srcx As Long, _
                                        ByVal srcy As Long, _
                                        ByVal srcwidth As Long, _
                                        ByVal srcheight As Long, _
                                        ByVal srcUnit As GpUnit, _
                                        Optional ByVal imageAttributes As Long = 0, _
                                        Optional ByVal callback As Long = 0, _
                              Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRect_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawImagePointsRect" _
                             (ByVal graphics As Long, _
                              ByVal Image As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal srcx As Single, _
                              ByVal srcy As Single, _
                              ByVal srcwidth As Single, _
                              ByVal srcheight As Single, _
                              ByVal srcUnit As GpUnit, _
                              Optional ByVal imageAttributes As Long = 0, _
                              Optional ByVal callback As Long = 0, _
                              Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipDrawImagePointsRectI_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawImagePointsRectI" _
                             (ByVal graphics As Long, _
                              ByVal Image As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal srcx As Long, _
                              ByVal srcy As Long, _
                              ByVal srcwidth As Long, _
                              ByVal srcheight As Long, _
                              ByVal srcUnit As GpUnit, _
                              Optional ByVal imageAttributes As Long = 0, _
                              Optional ByVal callback As Long = 0, _
                              Optional ByVal callbackData As Long = 0) As GpStatus

Public Declare Function GdipGetImageDecoders _
                         Lib "gdiplus" (ByVal numDecoders As Long, _
                                        ByVal Size As Long, _
                                        decoders As Any) As GpStatus
Public Declare Function GdipGetImageEncodersSize _
                         Lib "gdiplus" (numEncoders As Long, _
                                        Size As Long) As GpStatus
Public Declare Function GdipGetImageEncoders _
                         Lib "gdiplus" (ByVal numEncoders As Long, _
                                        ByVal Size As Long, _
                                        Encoders As Any) As GpStatus
Public Declare Function GdipComment _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal sizeData As Long, _
                                        Data As Any) As GpStatus

Public Declare Function GdipLoadImageFromFile _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromFileICM _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStream _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        Image As Long) As GpStatus
Public Declare Function GdipLoadImageFromStreamICM _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        Image As Long) As GpStatus
Public Declare Function GdipDisposeImage Lib "gdiplus" (ByVal Image As Long) As GpStatus
Public Declare Function GdipCloneImage _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        cloneImage As Long) As GpStatus

Public Declare Function GdipSaveImageToFile _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal fileName As Long, _
                                        clsidEncoder As CLSID, _
                                        encoderParams As Any) As GpStatus
Public Declare Function GdipSaveImageToStream _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal stream As Any, _
                                        clsidEncoder As CLSID, _
                                        encoderParams As Any) As GpStatus

Public Declare Function GdipSaveAdd _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        encoderParams As EncoderParameters) As GpStatus
Public Declare Function GdipSaveAddImage _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal newImage As Long, _
                                        encoderParams As EncoderParameters) As GpStatus

Public Declare Function GdipGetImageBounds _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        srcRect As RECTF, _
                                        srcUnit As GpUnit) As GpStatus
Public Declare Function GdipGetImageDimension _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Width As Single, _
                                        Height As Single) As GpStatus
Public Declare Function GdipGetImageType _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        itype As Image_Type) As GpStatus
Public Declare Function GdipGetImageWidth _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Width As Long) As GpStatus
Public Declare Function GdipGetImageHeight _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Height As Long) As GpStatus
Public Declare Function GdipGetImageHorizontalResolution _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Resolution As Single) As GpStatus
Public Declare Function GdipGetImageVerticalResolution _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Resolution As Single) As GpStatus
Public Declare Function GdipGetImageFlags _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Flags As Long) As GpStatus
Public Declare Function GdipGetImageRawFormat _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        format As CLSID) As GpStatus
Public Declare Function GdipGetImagePixelFormat _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        PixelFormat As GpPixelFormat) As GpStatus
Public Declare Function GdipGetImageThumbnail _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal thumbWidth As Long, _
                                        ByVal thumbHeight As Long, _
                                        thumbImage As Long, _
                                        Optional ByVal callback As Long = 0, _
                                        Optional ByVal callbackData As Long = 0) As GpStatus
Public Declare Function GdipGetEncoderParameterListSize _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        clsidEncoder As CLSID, _
                                        Size As Long) As GpStatus
Public Declare Function GdipGetEncoderParameterList _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        clsidEncoder As CLSID, _
                                        ByVal Size As Long, _
                                        Buffer As EncoderParameters) As GpStatus

Public Declare Function GdipImageGetFrameDimensionsCount _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameDimensionsList _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        dimensionIDs As CLSID, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipImageGetFrameCount _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        dimensionID As CLSID, _
                                        Count As Long) As GpStatus
Public Declare Function GdipImageSelectActiveFrame _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        dimensionID As CLSID, _
                                        ByVal frameIndex As Long) As GpStatus
Public Declare Function GdipImageRotateFlip _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal rfType As RotateFlipType) As GpStatus
Public Declare Function GdipGetImagePalette _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        palette As ColorPalette, _
                                        ByVal Size As Long) As GpStatus
Public Declare Function GdipSetImagePalette _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        palette As ColorPalette) As GpStatus
Public Declare Function GdipGetImagePaletteSize _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Size As Long) As GpStatus
Public Declare Function GdipGetPropertyCount _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        numOfProperty As Long) As GpStatus
Public Declare Function GdipGetPropertyIdList _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal numOfProperty As Long, _
                                        list As Long) As GpStatus
Public Declare Function GdipGetPropertyItemSize _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal propId As Long, _
                                        Size As Long) As GpStatus
Public Declare Function GdipGetPropertyItem _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal propId As Long, _
                                        ByVal propSize As Long, _
                                        Buffer As PropertyItem) As GpStatus
Public Declare Function GdipGetPropertySize _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        totalBufferSize As Long, _
                                        numProperties As Long) As GpStatus
Public Declare Function GdipGetAllPropertyItems _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal totalBufferSize As Long, _
                                        ByVal numProperties As Long, _
                                        allItems As PropertyItem) As GpStatus
Public Declare Function GdipRemovePropertyItem _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal propId As Long) As GpStatus
Public Declare Function GdipSetPropertyItem _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Item As PropertyItem) As GpStatus
Public Declare Function GdipImageForceValidation _
                         Lib "gdiplus" (ByVal Image As Long) As GpStatus

'==================================================

Public Declare Function GdipCreatePen1 _
                         Lib "gdiplus" (ByVal Color As Long, _
                                        ByVal Width As Single, _
                                        ByVal unit As GpUnit, _
                                        pen As Long) As GpStatus
Public Declare Function GdipCreatePen2 _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal Width As Single, _
                                        ByVal unit As GpUnit, _
                                        pen As Long) As GpStatus
Public Declare Function GdipClonePen _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        clonepen As Long) As GpStatus
Public Declare Function GdipDeletePen Lib "gdiplus" (ByVal pen As Long) As GpStatus

Public Declare Function GdipSetPenWidth _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal Width As Single) As GpStatus
Public Declare Function GdipGetPenWidth _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        Width As Single) As GpStatus
Public Declare Function GdipSetPenUnit _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal unit As GpUnit) As GpStatus
Public Declare Function GdipGetPenUnit _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        unit As GpUnit) As GpStatus

Public Declare Function GdipSetPenLineCap _
                         Lib "gdiplus" _
                             Alias "GdipSetPenLineCap197819" (ByVal pen As Long, _
                                                              ByVal startCap As LineCap, _
                                                              ByVal endCap As LineCap, _
                                                              ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenStartCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal startCap As LineCap) As GpStatus
Public Declare Function GdipSetPenEndCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal endCap As LineCap) As GpStatus
Public Declare Function GdipSetPenDashCap _
                         Lib "gdiplus" _
                             Alias "GdipSetPenDashCap197819" (ByVal pen As Long, _
                                                              ByVal dcap As DashCap) As GpStatus
Public Declare Function GdipGetPenStartCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        startCap As LineCap) As GpStatus
Public Declare Function GdipGetPenEndCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        endCap As LineCap) As GpStatus
Public Declare Function GdipGetPenDashCap _
                         Lib "gdiplus" _
                             Alias "GdipGetPenDashCap197819" (ByVal pen As Long, _
                                                              dcap As DashCap) As GpStatus
Public Declare Function GdipSetPenLineJoin _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal lnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetPenLineJoin _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        lnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetPenCustomStartCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomStartCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        customCap As Long) As GpStatus
Public Declare Function GdipSetPenCustomEndCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal customCap As Long) As GpStatus
Public Declare Function GdipGetPenCustomEndCap _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        customCap As Long) As GpStatus

Public Declare Function GdipSetPenMiterLimit _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal miterLimit As Single) As GpStatus
Public Declare Function GdipGetPenMiterLimit _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        miterLimit As Single) As GpStatus
Public Declare Function GdipSetPenMode _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal penMode As PenAlignment) As GpStatus
Public Declare Function GdipGetPenMode _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        penMode As PenAlignment) As GpStatus
Public Declare Function GdipSetPenTransform _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetPenTransform _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPenTransform _
                         Lib "gdiplus" (ByVal pen As Long) As GpStatus
Public Declare Function GdipMultiplyPenTransform _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal matrix As Long, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePenTransform _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePenTransform _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal sx As Single, _
                                        ByVal sy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePenTransform _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal angle As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipSetPenColor _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal argb As Long) As GpStatus
Public Declare Function GdipGetPenColor _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        argb As Long) As GpStatus
Public Declare Function GdipSetPenBrushFill _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal brush As Long) As GpStatus
Public Declare Function GdipGetPenBrushFill _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        brush As Long) As GpStatus
Public Declare Function GdipGetPenFillType _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ptype As PenType) As GpStatus
Public Declare Function GdipGetPenDashStyle _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        dStyle As DashStyle) As GpStatus
Public Declare Function GdipSetPenDashStyle _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal dStyle As DashStyle) As GpStatus
Public Declare Function GdipGetPenDashOffset _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        Offset As Single) As GpStatus
Public Declare Function GdipSetPenDashOffset _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        ByVal Offset As Single) As GpStatus
Public Declare Function GdipGetPenDashCount _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipSetPenDashArray _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        dash As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenDashArray _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        dash As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundCount _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipSetPenCompoundArray _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        dash As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPenCompoundArray _
                         Lib "gdiplus" (ByVal pen As Long, _
                                        dash As Single, _
                                        ByVal Count As Long) As GpStatus

Public Declare Function GdipCreateCustomLineCap _
                         Lib "gdiplus" (ByVal fillPath As Long, _
                                        ByVal strokePath As Long, _
                                        ByVal baseCap As LineCap, _
                                        ByVal baseInset As Single, _
                                        customCap As Long) As GpStatus
Public Declare Function GdipDeleteCustomLineCap _
                         Lib "gdiplus" (ByVal customCap As Long) As GpStatus
Public Declare Function GdipCloneCustomLineCap _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        clonedCap As Long) As GpStatus
Public Declare Function GdipGetCustomLineCapType _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        capType As CustomLineCapType) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeCaps _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        ByVal startCap As LineCap, _
                                        ByVal endCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeCaps _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        startCap As LineCap, _
                                        endCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapStrokeJoin _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        ByVal lnJoin As LineJoin) As GpStatus
Public Declare Function GdipGetCustomLineCapStrokeJoin _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        lnJoin As LineJoin) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseCap _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        ByVal baseCap As LineCap) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseCap _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        baseCap As LineCap) As GpStatus
Public Declare Function GdipSetCustomLineCapBaseInset _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        ByVal inset As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapBaseInset _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        inset As Single) As GpStatus
Public Declare Function GdipSetCustomLineCapWidthScale _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        ByVal widthScale As Single) As GpStatus
Public Declare Function GdipGetCustomLineCapWidthScale _
                         Lib "gdiplus" (ByVal customCap As Long, _
                                        widthScale As Single) As GpStatus

Public Declare Function GdipCreateAdjustableArrowCap _
                         Lib "gdiplus" (ByVal Height As Single, _
                                        ByVal Width As Single, _
                                        ByVal isFilled As Long, _
                                        cap As Long) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapHeight _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapHeight _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        Height As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapWidth _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        ByVal Width As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapWidth _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        Width As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapMiddleInset _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        ByVal middleInset As Single) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapMiddleInset _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        middleInset As Single) As GpStatus
Public Declare Function GdipSetAdjustableArrowCapFillState _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        ByVal bFillState As Long) As GpStatus
Public Declare Function GdipGetAdjustableArrowCapFillState _
                         Lib "gdiplus" (ByVal cap As Long, _
                                        bFillState As Long) As GpStatus

'==================================================

Public Declare Function GdipCreateBitmapFromFile _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromFileICM _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStream _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromStreamICM _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromScan0 _
                         Lib "gdiplus" (ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal stride As Long, _
                                        ByVal PixelFormat As GpPixelFormat, _
                                        scan0 As Any, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGraphics _
                         Lib "gdiplus" (ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal graphics As Long, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromGdiDib _
                         Lib "gdiplus" (gdiBitmapInfo As BITMAPINFO, _
                                        ByVal gdiBitmapData As Long, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHBITMAP _
                         Lib "gdiplus" (ByVal hbm As Long, _
                                        ByVal hPal As Long, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateHBITMAPFromBitmap _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        hbmReturn As Long, _
                                        ByVal background As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromHICON _
                         Lib "gdiplus" (ByVal hicon As Long, _
                                        bitmap As Long) As GpStatus
Public Declare Function GdipCreateHICONFromBitmap _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        hbmReturn As Long) As GpStatus
Public Declare Function GdipCreateBitmapFromResource _
                         Lib "gdiplus" (ByVal hInstance As Long, _
                                        ByVal lpBitmapName As Long, _
                                        bitmap As Long) As GpStatus

Public Declare Function GdipCloneBitmapArea _
                         Lib "gdiplus" (ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal PixelFormat As GpPixelFormat, _
                                        ByVal srcBitmap As Long, _
                                        dstBitmap As Long) As GpStatus
Public Declare Function GdipCloneBitmapAreaI _
                         Lib "gdiplus" (ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal PixelFormat As GpPixelFormat, _
                                        ByVal srcBitmap As Long, _
                                        dstBitmap As Long) As GpStatus

Public Declare Function GdipBitmapLockBits _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        rect As RECTL, _
                                        ByVal Flags As ImageLockMode, _
                                        ByVal PixelFormat As GpPixelFormat, _
                                        lockedBitmapData As BitmapData) As GpStatus
Public Declare Function GdipBitmapUnlockBits _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        lockedBitmapData As BitmapData) As GpStatus

Public Declare Function GdipBitmapGetPixel _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        Color As Long) As GpStatus
Public Declare Function GdipBitmapSetPixel _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Color As Long) As GpStatus

Public Declare Function GdipBitmapSetResolution _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        ByVal xdpi As Single, _
                                        ByVal ydpi As Single) As GpStatus

Public Declare Function GdipCreateCachedBitmap _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        ByVal graphics As Long, _
                                        cachedBitmap As Long) As GpStatus
Public Declare Function GdipDeleteCachedBitmap _
                         Lib "gdiplus" (ByVal cachedBitmap As Long) As GpStatus
Public Declare Function GdipDrawCachedBitmap _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal cachedBitmap As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long) As GpStatus

'==================================================

Public Declare Function GdipCloneBrush _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        cloneBrush As Long) As GpStatus
Public Declare Function GdipDeleteBrush Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipGetBrushType _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        brshType As BrushType) As GpStatus
Public Declare Function GdipCreateHatchBrush _
                         Lib "gdiplus" (ByVal style As HatchStyle, _
                                        ByVal forecolr As Long, _
                                        ByVal backcolr As Long, _
                                        brush As Long) As GpStatus
Public Declare Function GdipGetHatchStyle _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        style As HatchStyle) As GpStatus
Public Declare Function GdipGetHatchForegroundColor _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        forecolr As Long) As GpStatus
Public Declare Function GdipGetHatchBackgroundColor _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        backcolr As Long) As GpStatus
Public Declare Function GdipCreateSolidFill _
                         Lib "gdiplus" (ByVal argb As Long, _
                                        brush As Long) As GpStatus
Public Declare Function GdipSetSolidFillColor _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal argb As Long) As GpStatus
Public Declare Function GdipGetSolidFillColor _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        argb As Long) As GpStatus
Public Declare Function GdipCreateLineBrush _
                         Lib "gdiplus" (Point1 As POINTF, _
                                        Point2 As POINTF, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushI _
                         Lib "gdiplus" (Point1 As POINTL, _
                                        Point2 As POINTL, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRect _
                         Lib "gdiplus" (rect As RECTF, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long, _
                                        ByVal Mode As LinearGradientMode, _
                                        ByVal WrapMd As WrapMode, _
                                        lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectI _
                         Lib "gdiplus" (rect As RECTL, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long, _
                                        ByVal Mode As LinearGradientMode, _
                                        ByVal WrapMd As WrapMode, _
                                        lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngle _
                         Lib "gdiplus" (rect As RECTF, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long, _
                                        ByVal angle As Single, _
                                        ByVal isAngleScalable As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        lineGradient As Long) As GpStatus
Public Declare Function GdipCreateLineBrushFromRectWithAngleI _
                         Lib "gdiplus" (rect As RECTL, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long, _
                                        ByVal angle As Single, _
                                        ByVal isAngleScalable As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        lineGradient As Long) As GpStatus
Public Declare Function GdipSetLineColors _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal color1 As Long, _
                                        ByVal color2 As Long) As GpStatus
Public Declare Function GdipGetLineColors _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        lColors As Long) As GpStatus
Public Declare Function GdipGetLineRect _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        rect As RECTF) As GpStatus
Public Declare Function GdipGetLineRectI _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        rect As RECTL) As GpStatus
Public Declare Function GdipSetLineGammaCorrection _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineGammaCorrection _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetLineBlendCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLineBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipGetLineBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLineBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipSetLineBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlendCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetLinePresetBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipGetLinePresetBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipSetLinePresetBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipSetLinePresetBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus

Public Declare Function GdipSetLineSigmaBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal focus As Single, _
                                        ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineLinearBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal focus As Single, _
                                        ByVal theScale As Single) As GpStatus
Public Declare Function GdipSetLineWrapMode _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineWrapMode _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetLineTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        matrix As Long) As GpStatus
Public Declare Function GdipSetLineTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetLineTransform _
                         Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipMultiplyLineTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateLineTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleLineTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal sx As Single, _
                                        ByVal sy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateLineTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal angle As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipCreateTexture _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2 _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIA _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal imageAttributes As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        texture As Long) As GpStatus
Public Declare Function GdipCreateTexture2I _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        texture As Long) As GpStatus
Public Declare Function GdipCreateTextureIAI _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        ByVal imageAttributes As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        texture As Long) As GpStatus
Public Declare Function GdipGetTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipTranslateTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipMultiplyTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal sx As Single, _
                                        ByVal sy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateTextureTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal angle As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipSetTextureWrapMode _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureWrapMode _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetTextureImage _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Image As Long) As GpStatus
Public Declare Function GdipCreatePathGradient _
                         Lib "gdiplus" (Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI _
                         Lib "gdiplus" (Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal WrapMd As WrapMode, _
                                        polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradient_ _
                         Lib "gdiplus" _
                             Alias "GdipCreatePathGradient" _
                             (Points As Any, _
                              ByVal Count As Long, _
                              ByVal WrapMd As WrapMode, _
                              polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientI_ _
                         Lib "gdiplus" _
                             Alias "GdipCreatePathGradientI" _
                             (Points As Any, _
                              ByVal Count As Long, _
                              ByVal WrapMd As WrapMode, _
                              polyGradient As Long) As GpStatus
Public Declare Function GdipCreatePathGradientFromPath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        polyGradient As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterColor _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        lColors As Long) As GpStatus
Public Declare Function GdipSetPathGradientCenterColor _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal lColors As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorsWithCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        argb As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSurroundColorsWithCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        argb As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPath _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal path As Long) As GpStatus
Public Declare Function GdipSetPathGradientPath _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal path As Long) As GpStatus
Public Declare Function GdipGetPathGradientCenterPoint _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Points As POINTF) As GpStatus
Public Declare Function GdipGetPathGradientCenterPointI _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Points As POINTL) As GpStatus
Public Declare Function GdipSetPathGradientCenterPoint _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Points As POINTF) As GpStatus
Public Declare Function GdipSetPathGradientCenterPointI _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Points As POINTL) As GpStatus
Public Declare Function GdipGetPathGradientRect _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        rect As RECTF) As GpStatus
Public Declare Function GdipGetPathGradientRectI _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        rect As RECTL) As GpStatus
Public Declare Function GdipGetPathGradientPointCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientSurroundColorCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientGammaCorrection _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientGammaCorrection _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        useGammaCorrection As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlendCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipGetPathGradientBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipSetPathGradientBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlendCount _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        blend As Long, _
                                        positions As Single, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathGradientPresetBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipGetPathGradientPresetBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientPresetBlend_ _
                         Lib "gdiplus" _
                             Alias "GdipSetPathGradientPresetBlend" _
                             (ByVal brush As Long, _
                              blend As Any, _
                              positions As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipSetPathGradientSigmaBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal focus As Single, _
                                        ByVal sScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientLinearBlend _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal focus As Single, _
                                        ByVal sScale As Single) As GpStatus
Public Declare Function GdipGetPathGradientWrapMode _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        WrapMd As WrapMode) As GpStatus
Public Declare Function GdipSetPathGradientWrapMode _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal WrapMd As WrapMode) As GpStatus
Public Declare Function GdipGetPathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetPathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipResetPathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long) As GpStatus
Public Declare Function GdipMultiplyPathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal matrix As Long, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslatePathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScalePathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal sx As Single, _
                                        ByVal sy As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotatePathGradientTransform _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal angle As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipGetPathGradientFocusScales _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        xScale As Single, _
                                        yScale As Single) As GpStatus
Public Declare Function GdipSetPathGradientFocusScales _
                         Lib "gdiplus" (ByVal brush As Long, _
                                        ByVal xScale As Single, _
                                        ByVal yScale As Single) As GpStatus
Public Declare Function GdipCreatePath _
                         Lib "gdiplus" (ByVal brushmode As FillMode, _
                                        path As Long) As GpStatus
Public Declare Function GdipCreatePath2 _
                         Lib "gdiplus" (Points As POINTF, _
                                        types As Any, _
                                        ByVal Count As Long, _
                                        brushmode As FillMode, _
                                        path As Long) As GpStatus
Public Declare Function GdipCreatePath2I _
                         Lib "gdiplus" (Points As POINTL, _
                                        types As Any, _
                                        ByVal Count As Long, _
                                        brushmode As FillMode, _
                                        path As Long) As GpStatus
Public Declare Function GdipCreatePath2_ _
                         Lib "gdiplus" _
                             Alias "GdipCreatePath2" _
                             (Points As Any, _
                              types As Any, _
                              ByVal Count As Long, _
                              brushmode As FillMode, _
                              path As Long) As GpStatus
Public Declare Function GdipCreatePath2I_ _
                         Lib "gdiplus" _
                             Alias "GdipCreatePath2I" _
                             (Points As Any, _
                              types As Any, _
                              ByVal Count As Long, _
                              brushmode As FillMode, _
                              path As Long) As GpStatus
Public Declare Function GdipClonePath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        clonePath As Long) As GpStatus
Public Declare Function GdipDeletePath Lib "gdiplus" (ByVal path As Long) As GpStatus
Attribute GdipDeletePath.VB_UserMemId = 1879065532
Public Declare Function GdipResetPath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipGetPointCount _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipGetPathTypes _
                         Lib "gdiplus" (ByVal path As Long, _
                                        types As Any, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPoints _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPoints_ _
                         Lib "gdiplus" _
                             Alias "GdipGetPathPoints" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathPointsI_ _
                         Lib "gdiplus" _
                             Alias "GdipGetPathPointsI" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipGetPathFillMode _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal brushmode As FillMode) As GpStatus
Public Declare Function GdipSetPathFillMode _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal brushmode As FillMode) As GpStatus
Public Declare Function GdipGetPathData _
                         Lib "gdiplus" (ByVal path As Long, _
                                        pData As PathData) As GpStatus
Public Declare Function GdipStartPathFigure Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipClosePathFigure Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipClosePathFigures _
                         Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipSetPathMarker Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipClearPathMarkers _
                         Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipReversePath Lib "gdiplus" (ByVal path As Long) As GpStatus
Public Declare Function GdipGetPathLastPoint _
                         Lib "gdiplus" (ByVal path As Long, _
                                        lastPoint As POINTF) As GpStatus
Public Declare Function GdipAddPathLine _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal x1 As Single, _
                                        ByVal y1 As Single, _
                                        ByVal x2 As Single, _
                                        ByVal y2 As Single) As GpStatus
Public Declare Function GdipAddPathLine2 _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathLine2_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathLine2" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArc _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezier _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal x1 As Single, _
                                        ByVal y1 As Single, _
                                        ByVal x2 As Single, _
                                        ByVal y2 As Single, _
                                        ByVal x3 As Single, _
                                        ByVal y3 As Single, _
                                        ByVal x4 As Single, _
                                        ByVal y4 As Single) As GpStatus
Public Declare Function GdipAddPathBeziers _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2 _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3 _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal Offset As Long, _
                                        ByVal numberOfSegments As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2 _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathBeziers_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathBeziers" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathCurve" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathCurve2" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathCurve3_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathCurve3" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal Offset As Long, _
                              ByVal numberOfSegments As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurve_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathClosedCurve" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathClosedCurve2" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangle _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathRectangles _
                         Lib "gdiplus" (ByVal path As Long, _
                                        rect As RECTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipse _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single) As GpStatus
Public Declare Function GdipAddPathPie _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygon _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPolygon_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathPolygon" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal addingPath As Long, _
                                        ByVal bConnect As Long) As GpStatus
Public Declare Function GdipAddPathString _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal family As Long, _
                                        ByVal style As FontStyle, _
                                        ByVal emSize As Single, _
                                        layoutRect As RECTF, _
                                        ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathStringI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal family As Long, _
                                        ByVal style As FontStyle, _
                                        ByVal emSize As Single, _
                                        layoutRect As RECTL, _
                                        ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipAddPathLineI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal x1 As Long, _
                                        ByVal y1 As Long, _
                                        ByVal x2 As Long, _
                                        ByVal y2 As Long) As GpStatus
Public Declare Function GdipAddPathLine2I _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathLine2I_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathLine2I" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathArcI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathBezierI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal x1 As Long, _
                                        ByVal y1 As Long, _
                                        ByVal x2 As Long, _
                                        ByVal y2 As Long, _
                                        ByVal x3 As Long, _
                                        ByVal y3 As Long, _
                                        ByVal x4 As Long, _
                                        ByVal y4 As Long) As GpStatus
Public Declare Function GdipAddPathBeziersI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal Offset As Long, _
                                        ByVal numberOfSegments As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long, _
                                        ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathBeziersI_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathBeziersI" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurveI_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathCurveI" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathCurve2I_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathCurve2I" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Long) As GpStatus
Public Declare Function GdipAddPathCurve3I_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathCurve3I" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal Offset As Long, _
                              ByVal numberOfSegments As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathClosedCurveI_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathClosedCurveI" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathClosedCurve2I_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathClosedCurve2I" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal tension As Single) As GpStatus
Public Declare Function GdipAddPathRectangleI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathRectanglesI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        rects As RECTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathEllipseI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long) As GpStatus
Public Declare Function GdipAddPathPieI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal startAngle As Single, _
                                        ByVal sweepAngle As Single) As GpStatus
Public Declare Function GdipAddPathPolygonI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Points As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipAddPathPolygonI_ _
                         Lib "gdiplus" _
                             Alias "GdipAddPathPolygonI" _
                             (ByVal path As Long, _
                              Points As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipFlattenPath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        Optional ByVal matrix As Long = 0, _
                                        Optional ByVal flatness As Single = 0.25) As GpStatus
Public Declare Function GdipWindingModeOutline _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal matrix As Long, _
                                        ByVal flatness As Single) As GpStatus
Public Declare Function GdipWidenPath _
                         Lib "gdiplus" (ByVal NativePath As Long, _
                                        ByVal pen As Long, _
                                        ByVal matrix As Long, _
                                        ByVal flatness As Single) As GpStatus
Public Declare Function GdipWarpPath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal matrix As Long, _
                                        Points As POINTF, _
                                        ByVal Count As Long, _
                                        ByVal srcx As Single, _
                                        ByVal srcy As Single, _
                                        ByVal srcwidth As Single, _
                                        ByVal srcheight As Single, _
                                        ByVal WarpMd As WarpMode, _
                                        ByVal flatness As Single) As GpStatus
Public Declare Function GdipWarpPath_ _
                         Lib "gdiplus" _
                             Alias "GdipWarpPath" _
                             (ByVal path As Long, _
                              ByVal matrix As Long, _
                              Points As Any, _
                              ByVal Count As Long, _
                              ByVal srcx As Single, _
                              ByVal srcy As Single, _
                              ByVal srcwidth As Single, _
                              ByVal srcheight As Single, _
                              ByVal WarpMd As WarpMode, _
                              ByVal flatness As Single) As GpStatus
Public Declare Function GdipTransformPath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetPathWorldBounds _
                         Lib "gdiplus" (ByVal path As Long, _
                                        bounds As RECTF, _
                                        ByVal matrix As Long, _
                                        ByVal pen As Long) As GpStatus
Public Declare Function GdipGetPathWorldBoundsI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        bounds As RECTL, _
                                        ByVal matrix As Long, _
                                        ByVal pen As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPoint _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisiblePathPointI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPoint _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal pen As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsOutlineVisiblePathPointI _
                         Lib "gdiplus" (ByVal path As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal pen As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipCreatePathIter _
                         Lib "gdiplus" (iterator As Long, _
                                        ByVal path As Long) As GpStatus
Public Declare Function GdipDeletePathIter _
                         Lib "gdiplus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpath _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        startIndex As Long, _
                                        endIndex As Long, _
                                        isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextSubpathPath _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        ByVal path As Long, _
                                        isClosed As Long) As GpStatus
Public Declare Function GdipPathIterNextPathType _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        pathType As Any, _
                                        startIndex As Long, _
                                        endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarker _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        startIndex As Long, _
                                        endIndex As Long) As GpStatus
Public Declare Function GdipPathIterNextMarkerPath _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        ByVal path As Long) As GpStatus
Public Declare Function GdipPathIterGetCount _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipPathIterGetSubpathCount _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipPathIterIsValid _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        valid As Long) As GpStatus
Public Declare Function GdipPathIterHasCurve _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        hasCurve As Long) As GpStatus
Public Declare Function GdipPathIterRewind _
                         Lib "gdiplus" (ByVal iterator As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        Points As POINTF, _
                                        types As Any, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData _
                         Lib "gdiplus" (ByVal iterator As Long, _
                                        resultCount As Long, _
                                        Points As POINTF, _
                                        types As Any, _
                                        ByVal startIndex As Long, _
                                        ByVal endIndex As Long) As GpStatus
Public Declare Function GdipPathIterEnumerate_ _
                         Lib "gdiplus" _
                             Alias "GdipPathIterEnumerate" _
                             (ByVal iterator As Long, _
                              resultCount As Long, _
                              Points As Any, _
                              types As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipPathIterCopyData_ _
                         Lib "gdiplus" _
                             Alias "GdipPathIterCopyData" _
                             (ByVal iterator As Long, _
                              resultCount As Long, _
                              Points As Any, _
                              types As Any, _
                              ByVal startIndex As Long, _
                              ByVal endIndex As Long) As GpStatus
Public Declare Function GdipCreateMatrix Lib "gdiplus" (matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix2 _
                         Lib "gdiplus" (ByVal m11 As Single, _
                                        ByVal m12 As Single, _
                                        ByVal m21 As Single, _
                                        ByVal m22 As Single, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single, _
                                        matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3 _
                         Lib "gdiplus" (rect As RECTF, _
                                        dstplg As POINTF, _
                                        matrix As Long) As GpStatus
Public Declare Function GdipCreateMatrix3I _
                         Lib "gdiplus" (rect As RECTL, _
                                        dstplg As POINTL, _
                                        matrix As Long) As GpStatus
Public Declare Function GdipCloneMatrix _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        cloneMatrix As Long) As GpStatus
Public Declare Function GdipDeleteMatrix Lib "gdiplus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipSetMatrixElements _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal m11 As Single, _
                                        ByVal m12 As Single, _
                                        ByVal m21 As Single, _
                                        ByVal m22 As Single, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single) As GpStatus
Public Declare Function GdipMultiplyMatrix _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal matrix2 As Long, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipTranslateMatrix _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal offsetX As Single, _
                                        ByVal offsetY As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipScaleMatrix _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal scaleX As Single, _
                                        ByVal scaleY As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipRotateMatrix _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal angle As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipShearMatrix _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal shearX As Single, _
                                        ByVal shearY As Single, _
                                        ByVal order As MatrixOrder) As GpStatus
Public Declare Function GdipInvertMatrix Lib "gdiplus" (ByVal matrix As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        pts As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        pts As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        pts As POINTF, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        pts As POINTL, _
                                        ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPoints_ _
                         Lib "gdiplus" _
                             Alias "GdipTransformMatrixPoints" _
                             (ByVal matrix As Long, _
                              pts As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipTransformMatrixPointsI_ _
                         Lib "gdiplus" _
                             Alias "GdipTransformMatrixPointsI" _
                             (ByVal matrix As Long, _
                              pts As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPoints_ _
                         Lib "gdiplus" _
                             Alias "GdipVectorTransformMatrixPoints" _
                             (ByVal matrix As Long, _
                              pts As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipVectorTransformMatrixPointsI_ _
                         Lib "gdiplus" _
                             Alias "GdipVectorTransformMatrixPointsI" _
                             (ByVal matrix As Long, _
                              pts As Any, _
                              ByVal Count As Long) As GpStatus
Public Declare Function GdipGetMatrixElements _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        matrixOut As Single) As GpStatus
Public Declare Function GdipIsMatrixInvertible _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsMatrixIdentity _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsMatrixEqual _
                         Lib "gdiplus" (ByVal matrix As Long, _
                                        ByVal matrix2 As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipCreateRegion Lib "gdiplus" (region As Long) As GpStatus
Public Declare Function GdipCreateRegionRect _
                         Lib "gdiplus" (rect As RECTF, _
                                        region As Long) As GpStatus
Public Declare Function GdipCreateRegionRectI _
                         Lib "gdiplus" (rect As RECTL, _
                                        region As Long) As GpStatus
Public Declare Function GdipCreateRegionPath _
                         Lib "gdiplus" (ByVal path As Long, _
                                        region As Long) As GpStatus
Public Declare Function GdipCreateRegionRgnData _
                         Lib "gdiplus" (regionData As Any, _
                                        ByVal Size As Long, _
                                        region As Long) As GpStatus
Public Declare Function GdipCreateRegionHrgn _
                         Lib "gdiplus" (ByVal hRgn As Long, _
                                        region As Long) As GpStatus
Public Declare Function GdipCloneRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        cloneRegion As Long) As GpStatus
Public Declare Function GdipDeleteRegion Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetInfinite Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipSetEmpty Lib "gdiplus" (ByVal region As Long) As GpStatus
Public Declare Function GdipCombineRegionRect _
                         Lib "gdiplus" (ByVal region As Long, _
                                        rect As RECTF, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRectI _
                         Lib "gdiplus" (ByVal region As Long, _
                                        rect As RECTL, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionPath _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal path As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipCombineRegionRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal region2 As Long, _
                                        ByVal CombineMd As CombineMode) As GpStatus
Public Declare Function GdipTranslateRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal dx As Single, _
                                        ByVal dy As Single) As GpStatus
Public Declare Function GdipTranslateRegionI _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal dx As Long, _
                                        ByVal dy As Long) As GpStatus
Public Declare Function GdipTransformRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionBounds _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal graphics As Long, _
                                        rect As RECTF) As GpStatus
Public Declare Function GdipGetRegionBoundsI _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal graphics As Long, _
                                        rect As RECTL) As GpStatus
Public Declare Function GdipGetRegionHRgn _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal graphics As Long, _
                                        hRgn As Long) As GpStatus
Public Declare Function GdipIsEmptyRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsInfiniteRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsEqualRegion _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal region2 As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipGetRegionDataSize _
                         Lib "gdiplus" (ByVal region As Long, _
                                        bufferSize As Long) As GpStatus
Public Declare Function GdipGetRegionData _
                         Lib "gdiplus" (ByVal region As Long, _
                                        Buffer As Any, _
                                        ByVal bufferSize As Long, _
                                        sizeFilled As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPoint _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionPointI _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRect _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal X As Single, _
                                        ByVal Y As Single, _
                                        ByVal Width As Single, _
                                        ByVal Height As Single, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipIsVisibleRegionRectI _
                         Lib "gdiplus" (ByVal region As Long, _
                                        ByVal X As Long, _
                                        ByVal Y As Long, _
                                        ByVal Width As Long, _
                                        ByVal Height As Long, _
                                        ByVal graphics As Long, _
                                        result As Long) As GpStatus
Public Declare Function GdipGetRegionScansCount _
                         Lib "gdiplus" (ByVal region As Long, _
                                        Ucount As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScans _
                         Lib "gdiplus" (ByVal region As Long, _
                                        rects As RECTF, _
                                        Count As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipGetRegionScansI _
                         Lib "gdiplus" (ByVal region As Long, _
                                        rects As RECTL, _
                                        Count As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipCreateImageAttributes _
                         Lib "gdiplus" (imageattr As Long) As GpStatus
Public Declare Function GdipCloneImageAttributes _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        cloneImageattr As Long) As GpStatus
Public Declare Function GdipDisposeImageAttributes _
                         Lib "gdiplus" (ByVal imageattr As Long) As GpStatus
Public Declare Function GdipSetImageAttributesToIdentity _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipResetImageAttributes _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipSetImageAttributesColorMatrix _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        colourMatrix As Any, _
                                        grayMatrix As Any, _
                                        ByVal Flags As ColorMatrixFlags) As GpStatus
Public Declare Function GdipSetImageAttributesThreshold _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        ByVal threshold As Single) As GpStatus
Public Declare Function GdipSetImageAttributesGamma _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        ByVal gamma As Single) As GpStatus
Public Declare Function GdipSetImageAttributesNoOp _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipSetImageAttributesColorKeys _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        ByVal colorLow As Long, _
                                        ByVal colorHigh As Long) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannel _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjstType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        ByVal channelFlags As ColorChannelFlags) As GpStatus
Public Declare Function GdipSetImageAttributesOutputChannelColorProfile _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        ByVal colorProfileFilename As Long) As GpStatus
Public Declare Function GdipSetImageAttributesRemapTable _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal ClrAdjType As ColorAdjustType, _
                                        ByVal enableFlag As Long, _
                                        ByVal mapSize As Long, _
                                        map As Any) As GpStatus
Public Declare Function GdipSetImageAttributesWrapMode _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal wrap As WrapMode, _
                                        ByVal argb As Long, _
                                        ByVal bClamp As Long) As GpStatus
Public Declare Function GdipSetImageAttributesICMMode _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal bOn As Long) As GpStatus
Public Declare Function GdipGetImageAttributesAdjustedPalette _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        colorPal As ColorPalette, _
                                        ByVal ClrAdjType As ColorAdjustType) As GpStatus
Public Declare Function GdipCreateFontFamilyFromName _
                         Lib "gdiplus" (ByVal Name As Long, _
                                        ByVal fontCollection As Long, _
                                        fontFamily As Long) As GpStatus
Public Declare Function GdipDeleteFontFamily _
                         Lib "gdiplus" (ByVal fontFamily As Long) As GpStatus
Public Declare Function GdipCloneFontFamily _
                         Lib "gdiplus" (ByVal fontFamily As Long, _
                                        clonedFontFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySansSerif _
                         Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilySerif _
                         Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetGenericFontFamilyMonospace _
                         Lib "gdiplus" (nativeFamily As Long) As GpStatus
Public Declare Function GdipGetFamilyName _
                         Lib "gdiplus" (ByVal family As Long, _
                                        ByVal Name As Long, _
                                        ByVal language As Integer) As GpStatus
Public Declare Function GdipIsStyleAvailable _
                         Lib "gdiplus" (ByVal family As Long, _
                                        ByVal style As Long, _
                                        IsStyleAvailable As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerable _
                         Lib "gdiplus" (ByVal fontCollection As Long, _
                                        ByVal graphics As Long, _
                                        numFound As Long) As GpStatus
Public Declare Function GdipFontCollectionEnumerate _
                         Lib "gdiplus" (ByVal fontCollection As Long, _
                                        ByVal numSought As Long, _
                                        gpFamilies As Long, _
                                        ByVal numFound As Long, _
                                        ByVal graphics As Long) As GpStatus
Public Declare Function GdipGetEmHeight _
                         Lib "gdiplus" (ByVal family As Long, _
                                        ByVal style As FontStyle, _
                                        EmHeight As Integer) As GpStatus
Public Declare Function GdipGetCellAscent _
                         Lib "gdiplus" (ByVal family As Long, _
                                        ByVal style As FontStyle, _
                                        CellAscent As Integer) As GpStatus
Public Declare Function GdipGetCellDescent _
                         Lib "gdiplus" (ByVal family As Long, _
                                        ByVal style As FontStyle, _
                                        CellDescent As Integer) As GpStatus
Public Declare Function GdipGetLineSpacing _
                         Lib "gdiplus" (ByVal family As Long, _
                                        ByVal style As FontStyle, _
                                        LineSpacing As Integer) As GpStatus
Public Declare Function GdipCreateFontFromDC _
                         Lib "gdiplus" (ByVal hDC As Long, _
                                        createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontA _
                         Lib "gdiplus" (ByVal hDC As Long, _
                                        logfont As LOGFONTA, _
                                        createdfont As Long) As GpStatus
Public Declare Function GdipCreateFontFromLogfontW _
                         Lib "gdiplus" (ByVal hDC As Long, _
                                        logfont As LOGFONTW, _
                                        createdfont As Long) As GpStatus
Public Declare Function GdipCreateFont _
                         Lib "gdiplus" (ByVal fontFamily As Long, _
                                        ByVal emSize As Single, _
                                        ByVal style As FontStyle, _
                                        ByVal unit As GpUnit, _
                                        createdfont As Long) As GpStatus
Public Declare Function GdipCloneFont _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        cloneFont As Long) As GpStatus
Public Declare Function GdipDeleteFont Lib "gdiplus" (ByVal curFont As Long) As GpStatus
Public Declare Function GdipGetFamily _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        family As Long) As GpStatus
Public Declare Function GdipGetFontStyle _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        style As FontStyle) As GpStatus
Public Declare Function GdipGetFontSize _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        Size As Single) As GpStatus
Public Declare Function GdipGetFontUnit _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        unit As GpUnit) As GpStatus
Public Declare Function GdipGetFontHeight _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        ByVal graphics As Long, _
                                        Height As Single) As GpStatus
Public Declare Function GdipGetFontHeightGivenDPI _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        ByVal dpi As Single, _
                                        Height As Single) As GpStatus
Public Declare Function GdipGetLogFontA _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        ByVal graphics As Long, _
                                        logfont As LOGFONTA) As GpStatus
Public Declare Function GdipGetLogFontW _
                         Lib "gdiplus" (ByVal curFont As Long, _
                                        ByVal graphics As Long, _
                                        logfont As LOGFONTW) As GpStatus
Public Declare Function GdipNewInstalledFontCollection _
                         Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipNewPrivateFontCollection _
                         Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipDeletePrivateFontCollection _
                         Lib "gdiplus" (fontCollection As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyCount _
                         Lib "gdiplus" (ByVal fontCollection As Long, _
                                        numFound As Long) As GpStatus
Public Declare Function GdipGetFontCollectionFamilyList _
                         Lib "gdiplus" (ByVal fontCollection As Long, _
                                        ByVal numSought As Long, _
                                        gpFamilies As Long, _
                                        numFound As Long) As GpStatus
Public Declare Function GdipPrivateAddFontFile _
                         Lib "gdiplus" (ByVal fontCollection As Long, _
                                        ByVal fileName As Long) As GpStatus
Public Declare Function GdipPrivateAddMemoryFont _
                         Lib "gdiplus" (ByVal fontCollection As Long, _
                                        ByVal memory As Long, _
                                        ByVal Length As Long) As GpStatus
Public Declare Function GdipDrawString _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal thefont As Long, _
                                        layoutRect As RECTF, _
                                        ByVal StringFormat As Long, _
                                        ByVal brush As Long) As GpStatus
Public Declare Function GdipMeasureString _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal thefont As Long, _
                                        layoutRect As RECTF, _
                                        ByVal StringFormat As Long, _
                                        boundingBox As RECTF, _
                                        codepointsFitted As Long, _
                                        linesFilled As Long) As GpStatus
Public Declare Function GdipMeasureCharacterRanges _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal thefont As Long, _
                                        layoutRect As RECTF, _
                                        ByVal StringFormat As Long, _
                                        ByVal regionCount As Long, _
                                        regions As Long) As GpStatus
Public Declare Function GdipDrawDriverString _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal thefont As Long, _
                                        ByVal brush As Long, _
                                        positions As POINTF, _
                                        ByVal Flags As Long, _
                                        ByVal matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Str As Long, _
                                        ByVal Length As Long, _
                                        ByVal thefont As Long, _
                                        positions As POINTF, _
                                        ByVal Flags As Long, _
                                        ByVal matrix As Long, _
                                        boundingBox As RECTF) As GpStatus
Public Declare Function GdipDrawDriverString_ _
                         Lib "gdiplus" _
                             Alias "GdipDrawDriverString" _
                             (ByVal graphics As Long, _
                              ByVal Str As Long, _
                              ByVal Length As Long, _
                              ByVal thefont As Long, _
                              ByVal brush As Long, _
                              positions As Any, _
                              ByVal Flags As Long, _
                              ByVal matrix As Long) As GpStatus
Public Declare Function GdipMeasureDriverString_ _
                         Lib "gdiplus" _
                             Alias "GdipMeasureDriverString" _
                             (ByVal graphics As Long, _
                              ByVal Str As Long, _
                              ByVal Length As Long, _
                              ByVal thefont As Long, _
                              positions As Any, _
                              ByVal Flags As Long, _
                              ByVal matrix As Long, _
                              boundingBox As RECTF) As GpStatus
Public Declare Function GdipCreateStringFormat _
                         Lib "gdiplus" (ByVal formatAttributes As Long, _
                                        ByVal language As Integer, _
                                        StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericDefault _
                         Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipStringFormatGetGenericTypographic _
                         Lib "gdiplus" (StringFormat As Long) As GpStatus
Public Declare Function GdipDeleteStringFormat _
                         Lib "gdiplus" (ByVal StringFormat As Long) As GpStatus
Public Declare Function GdipCloneStringFormat _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        newFormat As Long) As GpStatus
Public Declare Function GdipSetStringFormatFlags _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal Flags As Long) As GpStatus
Public Declare Function GdipGetStringFormatFlags _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        Flags As Long) As GpStatus
Public Declare Function GdipSetStringFormatAlign _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatAlign _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatLineAlign _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal Align As StringAlignment) As GpStatus
Public Declare Function GdipGetStringFormatLineAlign _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        Align As StringAlignment) As GpStatus
Public Declare Function GdipSetStringFormatTrimming _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal trimming As StringTrimming) As GpStatus
Public Declare Function GdipGetStringFormatTrimming _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        trimming As Long) As GpStatus
Public Declare Function GdipSetStringFormatHotkeyPrefix _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipGetStringFormatHotkeyPrefix _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        hkPrefix As HotkeyPrefix) As GpStatus
Public Declare Function GdipSetStringFormatTabStops _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal firstTabOffset As Single, _
                                        ByVal Count As Long, _
                                        tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStops _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal Count As Long, _
                                        firstTabOffset As Single, _
                                        tabStops As Single) As GpStatus
Public Declare Function GdipGetStringFormatTabStopCount _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatDigitSubstitution _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal language As Integer, _
                                        ByVal substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatDigitSubstitution _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        language As Integer, _
                                        substitute As StringDigitSubstitute) As GpStatus
Public Declare Function GdipGetStringFormatMeasurableCharacterRangeCount _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        Count As Long) As GpStatus
Public Declare Function GdipSetStringFormatMeasurableCharacterRanges _
                         Lib "gdiplus" (ByVal StringFormat As Long, _
                                        ByVal rangeCount As Long, _
                                        ranges As CharacterRange) As GpStatus

'===================================================================================
'  GdiPlus 1.1 ������
'===================================================================================
#If GdipVersion >= 1.1 Then

    Public Const BlurEffectGuid As String = "{633C80A4-1843-482B-9EF2-BE2834C5FDD4}"
    Public Const BrightnessContrastEffectGuid As String = "{D3A1DBE1-8EC4-4C17-9F4C-EA97AD1C343D}"
    Public Const ColorBalanceEffectGuid As String = "{537E597D-251E-48DA-9664-29CA496B70F8}"
    Public Const ColorCurveEffectGuid As String = "{DD6A0022-58E4-4A67-9D9B-D48EB881A53D}"
    Public Const ColorLookupTableEffectGuid As String = "{A7CE72A9-0F7F-40D7-B3CC-D0C02D5C3212}"
    Public Const ColorMatrixEffectGuid As String = "{718F2615-7933-40E3-A511-5F68FE14DD74}"
    Public Const HueSaturationLightnessEffectGuid As String = "{8B2DD6C3-EB07-4D87-A5F0-7108E26A9C5F}"
    Public Const LevelsEffectGuid As String = "{99C354EC-2A31-4F3A-8C34-17A803B33A25}"
    Public Const RedEyeCorrectionEffectGuid As String = "{74D29D05-69A4-4266-9549-3CC52836B632}"
    Public Const SharpenEffectGuid As String = "{63CBF3EE-C526-402C-8F71-62C540BF5142}"
    Public Const TintEffectGuid As String = "{1077AF00-2848-4441-9489-44AD4C2D7A2C}"

Public Enum GdipEffectType
    Blur
    BrightnessContrast
    ColorBalance
    ColorCurve
    ColorLookupTable
    ColorMatrix
    HueSaturationLightness
    Levels
    RedEyeCorrection
    Sharpen
    Tint
End Enum

Public Enum HistogramFormat
    HistogramFormatARGB
    HistogramFormatPARGB
    HistogramFormatRGB
    HistogramFormatGray
    HistogramFormatB
    HistogramFormatG
    HistogramFormatR
    HistogramFormatA
End Enum

Public Enum CurveAdjustments
    AdjustExposure
    AdjustDensity
    AdjustContrast
    AdjustHighlight
    AdjustShadow
    AdjustMidtone
    AdjustWhiteSaturation
    AdjustBlackSaturation
End Enum

Public Enum CurveChannel
    CurveChannelAll
    CurveChannelRed
    CurveChannelGreen
    CurveChannelBlue
End Enum

Public Enum PaletteType
    PaletteTypeCustom = 0
    ' Optimal palette generated using a median-cut algorithm.
    PaletteTypeOptimal = 1
    ' Black and white palette.
    PaletteTypeFixedBW = 2
    ' Symmetric halftone palettes.
    ' Each of these halftone palettes will be a superset of the system palette.
    ' E.g. Halftone8 will have it's 8-color on-off primaries and the 16 system
    ' colors added. With duplicates removed, that leaves 16 colors.
    PaletteTypeFixedHalftone8 = 3   ' 8-color, on-off primaries
    PaletteTypeFixedHalftone27 = 4  ' 3 intensity levels of each color
    PaletteTypeFixedHalftone64 = 5  ' 4 intensity levels of each color
    PaletteTypeFixedHalftone125 = 6    ' 5 intensity levels of each color
    PaletteTypeFixedHalftone216 = 7    ' 6 intensity levels of each color
    ' Assymetric halftone palettes.
    ' These are somewhat less useful than the symmetric ones, but are
    ' included for completeness. These do not include all of the system
    ' colors.
    PaletteTypeFixedHalftone252 = 8    ' 6-red, 7-green, 6-blue intensities
    PaletteTypeFixedHalftone256 = 9    ' 8-red, 8-green, 4-blue intensities
End Enum

Public Enum DitherType
    DitherTypeNone = 0
    ' Solid color - picks the nearest matching color with no attempt to
    ' halftone or dither. May be used on an arbitrary palette.
    DitherTypeSolid = 1
    ' Ordered dithers and spiral dithers must be used with a fixed palette.
    ' NOTE: DitherOrdered4x4 is unique in that it may apply to 16bpp
    ' conversions also.
    DitherTypeOrdered4x4 = 2
    DitherTypeOrdered8x8 = 3
    DitherTypeOrdered16x16 = 4
    DitherTypeSpiral4x4 = 5
    DitherTypeSpiral8x8 = 6
    DitherTypeDualSpiral4x4 = 7
    DitherTypeDualSpiral8x8 = 8
    ' Error diffusion. May be used with any palette.
    DitherTypeErrorDiffusion = 9
    DitherTypeMax = 10
End Enum

Public Enum ItemDataPosition
    ItemDataPositionAfterHeader = 0
    ItemDataPositionAfterPalette = 1
    ItemDataPositionAfterBits = 2
End Enum

'struct __declspec(novtable) GdiplusAbort
'{
'    virtual HRESULT __stdcall Abort(void) = 0;
'};
Public Type GdiplusAbort
    AbortCallback As Long
End Type

Public Type ImageItemData
    Size As Long
    Position As Long
    pDesc As Long
    DescSize As Long
    pData As Long
    dataSize As Long
    Cookie As Long
End Type

Public Type SharpenParams
    radius As Single
    amount As Single
End Type

Public Type BlurParams
    radius As Single
    expandEdge As Long
End Type

Public Type BrightnessContrastParams
    brightnessLevel As Long
    contrastLevel As Long
End Type

Public Type RedEyeCorrectionParams
    numberOfAreas As Long
    areas As RECTL
End Type

Public Type HueSaturationLightnessParams
    hueLevel As Long
    saturationLevel As Long
    lightnessLevel As Long
End Type

Public Type TintParams
    hue As Long
    amount As Long
End Type

Public Type LevelsParams
    highlight As Long
    midtone As Long
    shadow As Long
End Type

Public Type ColorBalanceParams
    cyanRed As Long
    magentaGreen As Long
    yellowBlue As Long
End Type

Public Type ColorLUTParams
    lutB(0 To 255) As Byte
    lutG(0 To 255) As Byte
    lutR(0 To 255) As Byte
    lutA(0 To 255) As Byte
End Type

Public Type ColorCurveParams
    adjustment As CurveAdjustments
    channel As CurveChannel
    adjustValue As Long
End Type

Public Declare Function GdipCreateEffect _
                         Lib "gdiplus" (ByVal Guid41 As Long, _
                                        ByVal Guid42 As Long, _
                                        ByVal Guid43 As Long, _
                                        ByVal Guid44 As Long, _
                                        effect As Long) As GpStatus

Public Declare Function GdipDeleteEffect _
                         Lib "gdiplus" (ByVal effect As Long) As GpStatus
Public Declare Function GdipGetEffectParameterSize _
                         Lib "gdiplus" (ByVal effect As Long, _
                                        Size As Long) As GpStatus
Public Declare Function GdipSetEffectParameters _
                         Lib "gdiplus" (ByVal effect As Long, _
                                        Params As Any, _
                                        ByVal Size As Long) As GpStatus
Public Declare Function GdipGetEffectParameters _
                         Lib "gdiplus" (ByVal effect As Long, _
                                        Size As Long, _
                                        Params As Any) As GpStatus

Public Declare Function GdipImageSetAbort _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        IAbort As GdiplusAbort) As GpStatus
Public Declare Function GdipGraphicsSetAbort _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        IAbort As GdiplusAbort) As GpStatus
Public Declare Function GdipBitmapConvertFormat _
                         Lib "gdiplus" (ByVal InputBitmap As Long, _
                                        ByVal format As GpPixelFormat, _
                                        ByVal DitherType As DitherType, _
                                        ByVal PaletteType As PaletteType, _
                                        palette As ColorPalette, _
                                        ByVal alphaThresholdPercent As Single) As GpStatus
Public Declare Function GdipInitializePalette _
                         Lib "gdiplus" (palette As ColorPalette, _
                                        ByVal PaletteType As PaletteType, _
                                        ByVal optimalColors As Long, _
                                        ByVal useTransparentColor As Long, _
                                        Optional ByVal bitmap As Long) As GpStatus
Public Declare Function GdipBitmapApplyEffect _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        ByVal effect As Long, _
                                        roi As RECTL, _
                                        ByVal useAuxData As Long, _
                                        auxData As Any, _
                                        auxDataSize As Long) As GpStatus
Public Declare Function GdipBitmapCreateApplyEffect _
                         Lib "gdiplus" (inputBitmaps As Any, _
                                        ByVal numInputs As Long, _
                                        ByVal effect As Long, _
                                        roi As RECTL, _
                                        outputRect As RECTL, _
                                        outputBitmap As Long, _
                                        ByVal useAuxData As Long, _
                                        auxData As Any, _
                                        auxDataSize As Long) As GpStatus
Public Declare Function GdipBitmapGetHistogram _
                         Lib "gdiplus" (ByVal bitmap As Long, _
                                        ByVal format As HistogramFormat, _
                                        ByVal NumberOfEntries As Long, _
                                        channel0 As Any, _
                                        channel1 As Any, _
                                        channel2 As Any, _
                                        channel3 As Any) As GpStatus
Public Declare Function GdipBitmapGetHistogramSize _
                         Lib "gdiplus" (ByVal format As HistogramFormat, _
                                        NumberOfEntries As Long) As GpStatus

Public Declare Function GdipFindFirstImageItem _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Item As ImageItemData) As GpStatus
Public Declare Function GdipFindNextImageItem _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Item As ImageItemData) As GpStatus
Public Declare Function GdipGetImageItemData _
                         Lib "gdiplus" (ByVal Image As Long, _
                                        Item As ImageItemData) As GpStatus

Public Declare Function GdipDrawImageFX _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal Image As Long, _
                                        Source As RECTF, _
                                        ByVal xForm As Long, _
                                        ByVal effect As Long, _
                                        ByVal imageAttributes As Long, _
                                        ByVal srcUnit As GpUnit) As GpStatus

#End If

'===================================================================================
'  ����ô���õĶ���
'===================================================================================

'=================================
'== Structures                  ==
'=================================

'=================================
'Log Font Structure
Public Type LOGFONTA
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32) As Byte
End Type

Public Type LOGFONTW
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName(32) As Byte
End Type

'=================================
'Image
Public Type ImageCodecInfo
    ClassID As CLSID
    FormatID As CLSID
    CodecName As Long
    DllName As Long
    FormatDescription As Long
    FilenameExtension As Long
    MimeType As Long
    Flags As ImageCodecFlags
    Version As Long
    SigCount As Long
    SigSize As Long
    SigPattern As Long
    SigMask As Long
End Type

'=================================
'Colors
Public Type ColorPalette
    Flags As PaletteFlags
    Count As Long
    Entries(0 To 255) As Long
End Type

'=================================
'Meta File
Public Type PWMFRect16
    Left As Integer
    Top As Integer
    Right As Integer
    Bottom As Integer
End Type

Public Type WmfPlaceableFileHeader
    Key As Long        ' GDIP_WMF_PLACEABLEKEY
    Hmf As Integer     ' Metafile HANDLE number (always 0)
    boundingBox As PWMFRect16          ' Coordinates in metafile units
    Inch As Integer      ' Number of metafile units per inch
    Reserved As Long             ' Reserved (always 0)
    Checksum As Integer          ' Checksum value for previous 10 WORDs
End Type

Public Type ENHMETAHEADER3
    itype As Long     ' Record type EMR_HEADER
    nSize As Long     ' Record size in bytes.  This may be greater
    ' than the sizeof(ENHMETAHEADER).
    rclBounds As RECTL        ' Inclusive-inclusive bounds in device units
    rclFrame As RECTL       ' Inclusive-inclusive Picture Frame .01mm unit
    dSignature As Long          ' Signature.  Must be ENHMETA_SIGNATURE.
    nVersion As Long        ' Version number
    nBytes As Long      ' Size of the metafile in bytes
    nRecords As Long        ' Number of records in the metafile
    nHandles As Integer     ' Number of handles in the handle table
    ' Handle index zero is reserved.
    sReserved As Integer      ' Reserved.  Must be zero.
    nDescription As Long            ' Number of chars in the unicode desc string
    ' This is 0 if there is no description string
    offDescription As Long              ' Offset to the metafile description record.
    ' This is 0 if there is no description string
    nPalEntries As Long           ' Number of entries in the metafile palette.
    szlDevice As SIZEL        ' Size of the reference device in pels
    szlMillimeters As SIZEL             ' Size of the reference device in millimeters
End Type

Public Type METAHEADER
    mtType As Integer
    mtHeaderSize As Integer
    mtVersion As Integer
    mtSize As Long
    mtNoObjects As Integer
    mtMaxRecord As Long
    mtNoParameters As Integer
End Type

Public Type MetafileHeader
    mType As MetafileType
    Size As Long    ' Size of the metafile (in bytes)
    Version As Long    ' EMF+, EMF, or WMF version
    EmfPlusFlags As Long
    DpiX As Single
    DpiY As Single
    X As Long    ' Bounds in device units
    Y As Long
    Width As Long
    Height As Long

    EmfHeader As ENHMETAHEADER3    ' NOTE: You'll have to use CopyMemory to view the METAHEADER type
    EmfPlusHeaderSize As Long           ' size of the EMF+ header in file
    LogicalDpiX As Long     ' Logical Dpi of reference Hdc
    LogicalDpiY As Long     ' usually valid only for EMF+
End Type

'=================================
'Other
Public Type PropertyItem
    propId As Long              ' ID of this property
    Length As Long              ' Length of the property value, in bytes

    Type As Integer         ' Type of the value, as one of TAG_TYPE_XXX
        ' defined above
        Value As Long             ' property value
    End Type

Public Type CharacterRange
    First As Long
    Length As Long
End Type

'=================================
'== Enums                       ==
'=================================

'=================================
'Image
Public Enum GpImageSaveFormat
    GpSaveBMP = 0
    GpSaveJPEG = 1
    GpSaveGIF = 2
    GpSavePNG = 3
    GpSaveTIFF = 4
End Enum

Public Enum GpImageFormatIdentifiers
    GpImageFormatUndefined = 0
    GpImageFormatMemoryBMP = 1
    GpImageFormatBMP = 2
    GpImageFormatEMF = 3
    GpImageFormatWMF = 4
    GpImageFormatJPEG = 5
    GpImageFormatPNG = 6
    GpImageFormatGIF = 7
    GpImageFormatTIFF = 8
    GpImageFormatEXIF = 9
    GpImageFormatIcon = 10
End Enum

Public Enum Image_Type
    ImageTypeUnknown = 0
    ImageTypeBitmap = 1
    ImageTypeMetafile = 2
End Enum

Public Enum Image_Property_Types
    PropertyTagTypeByte = 1
    PropertyTagTypeASCII = 2
    PropertyTagTypeShort = 3
    PropertyTagTypeLong = 4
    PropertyTagTypeRational = 5
    PropertyTagTypeUndefined = 7
    PropertyTagTypeSLONG = 9
    PropertyTagTypeSRational = 10
End Enum

Public Enum ImageCodecFlags
    ImageCodecFlagsEncoder = &H1
    ImageCodecFlagsDecoder = &H2
    ImageCodecFlagsSupportBitmap = &H4
    ImageCodecFlagsSupportVector = &H8
    ImageCodecFlagsSeekableEncode = &H10
    ImageCodecFlagsBlockingDecode = &H20

    ImageCodecFlagsBuiltin = &H10000
    ImageCodecFlagsSystem = &H20000
    ImageCodecFlagsUser = &H40000
End Enum

Public Enum Image_Property_ID_Tags
    PropertyTagExifIFD = &H8769
    PropertyTagGpsIFD = &H8825

    PropertyTagNewSubfileType = &HFE
    PropertyTagSubfileType = &HFF
    PropertyTagImageWidth = &H100
    PropertyTagImageHeight = &H101
    PropertyTagBitsPerSample = &H102
    PropertyTagCompression = &H103
    PropertyTagPhotometricInterp = &H106
    PropertyTagThreshHolding = &H107
    PropertyTagCellWidth = &H108
    PropertyTagCellHeight = &H109
    PropertyTagFillOrder = &H10A
    PropertyTagDocumentName = &H10D
    PropertyTagImageDescription = &H10E
    PropertyTagEquipMake = &H10F
    PropertyTagEquipModel = &H110
    PropertyTagStripOffsets = &H111
    PropertyTagOrientation = &H112
    PropertyTagSamplesPerPixel = &H115
    PropertyTagRowsPerStrip = &H116
    PropertyTagStripBytesCount = &H117
    PropertyTagMinSampleValue = &H118
    PropertyTagMaxSampleValue = &H119
    PropertyTagXResolution = &H11A            ' Image resolution in width direction
    PropertyTagYResolution = &H11B            ' Image resolution in height direction
    PropertyTagPlanarConfig = &H11C           ' Image data arrangement
    PropertyTagPageName = &H11D
    PropertyTagXPosition = &H11E
    PropertyTagYPosition = &H11F
    PropertyTagFreeOffset = &H120
    PropertyTagFreeByteCounts = &H121
    PropertyTagGrayResponseUnit = &H122
    PropertyTagGrayResponseCurve = &H123
    PropertyTagT4Option = &H124
    PropertyTagT6Option = &H125
    PropertyTagResolutionUnit = &H128         ' Unit of X and Y resolution
    PropertyTagPageNumber = &H129
    PropertyTagTransferFuncition = &H12D
    PropertyTagSoftwareUsed = &H131
    PropertyTagDateTime = &H132
    PropertyTagArtist = &H13B
    PropertyTagHostComputer = &H13C
    PropertyTagPredictor = &H13D
    PropertyTagWhitePoint = &H13E
    PropertyTagPrimaryChromaticities = &H13F
    PropertyTagColorMap = &H140
    PropertyTagHalftoneHints = &H141
    PropertyTagTileWidth = &H142
    PropertyTagTileLength = &H143
    PropertyTagTileOffset = &H144
    PropertyTagTileByteCounts = &H145
    PropertyTagInkSet = &H14C
    PropertyTagInkNames = &H14D
    PropertyTagNumberOfInks = &H14E
    PropertyTagDotRange = &H150
    PropertyTagTargetPrinter = &H151
    PropertyTagExtraSamples = &H152
    PropertyTagSampleFormat = &H153
    PropertyTagSMinSampleValue = &H154
    PropertyTagSMaxSampleValue = &H155
    PropertyTagTransferRange = &H156

    PropertyTagJPEGProc = &H200
    PropertyTagJPEGInterFormat = &H201
    PropertyTagJPEGInterLength = &H202
    PropertyTagJPEGRestartInterval = &H203
    PropertyTagJPEGLosslessPredictors = &H205
    PropertyTagJPEGPointTransforms = &H206
    PropertyTagJPEGQTables = &H207
    PropertyTagJPEGDCTables = &H208
    PropertyTagJPEGACTables = &H209

    PropertyTagYCbCrCoefficients = &H211
    PropertyTagYCbCrSubsampling = &H212
    PropertyTagYCbCrPositioning = &H213
    PropertyTagREFBlackWhite = &H214

    PropertyTagICCProfile = &H8773            ' This TAG is defined by ICC
    ' for embedded ICC in TIFF
    PropertyTagGamma = &H301
    PropertyTagICCProfileDescriptor = &H302
    PropertyTagSRGBRenderingIntent = &H303

    PropertyTagImageTitle = &H320
    PropertyTagCopyright = &H8298

    PropertyTagResolutionXUnit = &H5001
    PropertyTagResolutionYUnit = &H5002
    PropertyTagResolutionXLengthUnit = &H5003
    PropertyTagResolutionYLengthUnit = &H5004
    PropertyTagPrintFlags = &H5005
    PropertyTagPrintFlagsVersion = &H5006
    PropertyTagPrintFlagsCrop = &H5007
    PropertyTagPrintFlagsBleedWidth = &H5008
    PropertyTagPrintFlagsBleedWidthScale = &H5009
    PropertyTagHalftoneLPI = &H500A
    PropertyTagHalftoneLPIUnit = &H500B
    PropertyTagHalftoneDegree = &H500C
    PropertyTagHalftoneShape = &H500D
    PropertyTagHalftoneMisc = &H500E
    PropertyTagHalftoneScreen = &H500F
    PropertyTagJPEGQuality = &H5010
    PropertyTagGridSize = &H5011
    PropertyTagThumbnailFormat = &H5012            ' 1 = JPEG, 0 = RAW RGB
    PropertyTagThumbnailWidth = &H5013
    PropertyTagThumbnailHeight = &H5014
    PropertyTagThumbnailColorDepth = &H5015
    PropertyTagThumbnailPlanes = &H5016
    PropertyTagThumbnailRawBytes = &H5017
    PropertyTagThumbnailSize = &H5018
    PropertyTagThumbnailCompressedSize = &H5019
    PropertyTagColorTransferFunction = &H501A
    PropertyTagThumbnailData = &H501B
    PropertyTagThumbnailImageWidth = &H5020        ' Thumbnail width
    PropertyTagThumbnailImageHeight = &H5021       ' Thumbnail height
    PropertyTagThumbnailBitsPerSample = &H5022     ' Number of bits per
    ' component
    PropertyTagThumbnailCompression = &H5023       ' Compression Scheme
    PropertyTagThumbnailPhotometricInterp = &H5024    ' Pixel composition
    PropertyTagThumbnailImageDescription = &H5025  ' Image Tile
    PropertyTagThumbnailEquipMake = &H5026         ' Manufacturer of Image
    ' Input equipment
    PropertyTagThumbnailEquipModel = &H5027        ' Model of Image input
    ' equipment
    PropertyTagThumbnailStripOffsets = &H5028      ' Image data location
    PropertyTagThumbnailOrientation = &H5029       ' Orientation of image
    PropertyTagThumbnailSamplesPerPixel = &H502A   ' Number of components
    PropertyTagThumbnailRowsPerStrip = &H502B      ' Number of rows per strip
    PropertyTagThumbnailStripBytesCount = &H502C   ' Bytes per compressed
    ' strip
    PropertyTagThumbnailResolutionX = &H502D       ' Resolution in width
    ' direction
    PropertyTagThumbnailResolutionY = &H502E       ' Resolution in height
    ' direction
    PropertyTagThumbnailPlanarConfig = &H502F      ' Image data arrangement
    PropertyTagThumbnailResolutionUnit = &H5030    ' Unit of X and Y
    ' Resolution
    PropertyTagThumbnailTransferFunction = &H5031  ' Transfer function
    PropertyTagThumbnailSoftwareUsed = &H5032      ' Software used
    PropertyTagThumbnailDateTime = &H5033          ' File change date and
    ' time
    PropertyTagThumbnailArtist = &H5034            ' Person who created the
    ' image
    PropertyTagThumbnailWhitePoint = &H5035        ' White point chromaticity
    PropertyTagThumbnailPrimaryChromaticities = &H5036
    ' Chromaticities of
    ' primaries
    PropertyTagThumbnailYCbCrCoefficients = &H5037    ' Color space transforma-
    ' tion coefficients
    PropertyTagThumbnailYCbCrSubsampling = &H5038  ' Subsampling ratio of Y
    ' to C
    PropertyTagThumbnailYCbCrPositioning = &H5039  ' Y and C position
    PropertyTagThumbnailRefBlackWhite = &H503A     ' Pair of black and white
    ' reference values
    PropertyTagThumbnailCopyRight = &H503B         ' CopyRight holder

    PropertyTagLuminanceTable = &H5090
    PropertyTagChrominanceTable = &H5091

    PropertyTagFrameDelay = &H5100
    PropertyTagLoopCount = &H5101
    #If GdipVersion >= 1.1 Then
    PropertyTagGlobalPalette = &H5102
    PropertyTagIndexBackground = &H5103
    PropertyTagIndexTransparent = &H5104
    #End If

    PropertyTagPixelUnit = &H5110          ' Unit specifier for pixel/unit
    PropertyTagPixelPerUnitX = &H5111      ' Pixels per unit in X
    PropertyTagPixelPerUnitY = &H5112      ' Pixels per unit in Y
    PropertyTagPaletteHistogram = &H5113   ' Palette histogram

    PropertyTagExifExposureTime = &H829A
    PropertyTagExifFNumber = &H829D

    PropertyTagExifExposureProg = &H8822
    PropertyTagExifSpectralSense = &H8824
    PropertyTagExifISOSpeed = &H8827
    PropertyTagExifOECF = &H8828

    PropertyTagExifVer = &H9000
    PropertyTagExifDTOrig = &H9003         ' Date & time of original
    PropertyTagExifDTDigitized = &H9004    ' Date & time of digital data generation

    PropertyTagExifCompConfig = &H9101
    PropertyTagExifCompBPP = &H9102

    PropertyTagExifShutterSpeed = &H9201
    PropertyTagExifAperture = &H9202
    PropertyTagExifBrightness = &H9203
    PropertyTagExifExposureBias = &H9204
    PropertyTagExifMaxAperture = &H9205
    PropertyTagExifSubjectDist = &H9206
    PropertyTagExifMeteringMode = &H9207
    PropertyTagExifLightSource = &H9208
    PropertyTagExifFlash = &H9209
    PropertyTagExifFocalLength = &H920A
    PropertyTagExifMakerNote = &H927C
    PropertyTagExifUserComment = &H9286
    PropertyTagExifDTSubsec = &H9290        ' Date & Time subseconds
    PropertyTagExifDTOrigSS = &H9291        ' Date & Time original subseconds
    PropertyTagExifDTDigSS = &H9292         ' Date & TIme digitized subseconds

    PropertyTagExifFPXVer = &HA000
    PropertyTagExifColorSpace = &HA001
    PropertyTagExifPixXDim = &HA002
    PropertyTagExifPixYDim = &HA003
    PropertyTagExifRelatedWav = &HA004      ' related sound file
    PropertyTagExifInterop = &HA005
    PropertyTagExifFlashEnergy = &HA20B
    PropertyTagExifSpatialFR = &HA20C       ' Spatial Frequency Response
    PropertyTagExifFocalXRes = &HA20E       ' Focal Plane X Resolution
    PropertyTagExifFocalYRes = &HA20F       ' Focal Plane Y Resolution
    PropertyTagExifFocalResUnit = &HA210    ' Focal Plane Resolution Unit
    PropertyTagExifSubjectLoc = &HA214
    PropertyTagExifExposureIndex = &HA215
    PropertyTagExifSensingMethod = &HA217
    PropertyTagExifFileSource = &HA300
    PropertyTagExifSceneType = &HA301
    PropertyTagExifCfaPattern = &HA302

    PropertyTagGpsVer = &H0
    PropertyTagGpsLatitudeRef = &H1
    PropertyTagGpsLatitude = &H2
    PropertyTagGpsLongitudeRef = &H3
    PropertyTagGpsLongitude = &H4
    PropertyTagGpsAltitudeRef = &H5
    PropertyTagGpsAltitude = &H6
    PropertyTagGpsGpsTime = &H7
    PropertyTagGpsGpsSatellites = &H8
    PropertyTagGpsGpsStatus = &H9
    PropertyTagGpsGpsMeasureMode = &HA
    PropertyTagGpsGpsDop = &HB              ' Measurement precision
    PropertyTagGpsSpeedRef = &HC
    PropertyTagGpsSpeed = &HD
    PropertyTagGpsTrackRef = &HE
    PropertyTagGpsTrack = &HF
    PropertyTagGpsImgDirRef = &H10
    PropertyTagGpsImgDir = &H11
    PropertyTagGpsMapDatum = &H12
    PropertyTagGpsDestLatRef = &H13
    PropertyTagGpsDestLat = &H14
    PropertyTagGpsDestLongRef = &H15
    PropertyTagGpsDestLong = &H16
    PropertyTagGpsDestBearRef = &H17
    PropertyTagGpsDestBear = &H18
    PropertyTagGpsDestDistRef = &H19
    PropertyTagGpsDestDist = &H1A
End Enum

'=================================
'Palette
Public Enum PaletteFlags
    PaletteFlagsHasAlpha = &H1
    PaletteFlagsGrayScale = &H2
    PaletteFlagsHalftone = &H4
End Enum

'=================================
'Rotate
Public Enum RotateFlipType
    RotateNoneFlipNone = 0
    Rotate90FlipNone = 1
    Rotate180FlipNone = 2
    Rotate270FlipNone = 3

    RotateNoneFlipX = 4
    Rotate90FlipX = 5
    Rotate180FlipX = 6
    Rotate270FlipX = 7

    RotateNoneFlipY = Rotate180FlipX
    Rotate90FlipY = Rotate270FlipX
    Rotate180FlipY = RotateNoneFlipX
    Rotate270FlipY = Rotate90FlipX

    RotateNoneFlipXY = Rotate180FlipNone
    Rotate90FlipXY = Rotate270FlipNone
    Rotate180FlipXY = RotateNoneFlipNone
    Rotate270FlipXY = Rotate90FlipNone
End Enum

'=================================
'Colors
Public Enum colors
    AliceBlue = &HFFF0F8FF
    AntiqueWhite = &HFFFAEBD7
    Aqua = &HFF00FFFF
    Aquamarine = &HFF7FFFD4
    Azure = &HFFF0FFFF
    Beige = &HFFF5F5DC
    Bisque = &HFFFFE4C4
    Black = &HFF000000
    BlanchedAlmond = &HFFFFEBCD
    Blue = &HFF0000FF
    BlueViolet = &HFF8A2BE2
    Brown = &HFFA52A2A
    BurlyWood = &HFFDEB887
    CadetBlue = &HFF5F9EA0
    Chartreuse = &HFF7FFF00
    Chocolate = &HFFD2691E
    Coral = &HFFFF7F50
    CornflowerBlue = &HFF6495ED
    Cornsilk = &HFFFFF8DC
    Crimson = &HFFDC143C
    Cyan = &HFF00FFFF
    DarkBlue = &HFF00008B
    DarkCyan = &HFF008B8B
    DarkGoldenrod = &HFFB8860B
    DarkGray = &HFFA9A9A9
    DarkGreen = &HFF006400
    DarkKhaki = &HFFBDB76B
    DarkMagenta = &HFF8B008B
    DarkOliveGreen = &HFF556B2F
    DarkOrange = &HFFFF8C00
    DarkOrchid = &HFF9932CC
    DarkRed = &HFF8B0000
    DarkSalmon = &HFFE9967A
    DarkSeaGreen = &HFF8FBC8B
    DarkSlateBlue = &HFF483D8B
    DarkSlateGray = &HFF2F4F4F
    DarkTurquoise = &HFF00CED1
    DarkViolet = &HFF9400D3
    DeepPink = &HFFFF1493
    DeepSkyBlue = &HFF00BFFF
    DimGray = &HFF696969
    DodgerBlue = &HFF1E90FF
    Firebrick = &HFFB22222
    FloralWhite = &HFFFFFAF0
    ForestGreen = &HFF228B22
    Fuchsia = &HFFFF00FF
    Gainsboro = &HFFDCDCDC
    GhostWhite = &HFFF8F8FF
    Gold = &HFFFFD700
    Goldenrod = &HFFDAA520
    Gray = &HFF808080
    Green = &HFF008000
    GreenYellow = &HFFADFF2F
    Honeydew = &HFFF0FFF0
    HotPink = &HFFFF69B4
    IndianRed = &HFFCD5C5C
    Indigo = &HFF4B0082
    Ivory = &HFFFFFFF0
    Khaki = &HFFF0E68C
    Lavender = &HFFE6E6FA
    LavenderBlush = &HFFFFF0F5
    LawnGreen = &HFF7CFC00
    LemonChiffon = &HFFFFFACD
    LightBlue = &HFFADD8E6
    LightCoral = &HFFF08080
    LightCyan = &HFFE0FFFF
    LightGoldenrodYellow = &HFFFAFAD2
    LightGray = &HFFD3D3D3
    LightGreen = &HFF90EE90
    LightPink = &HFFFFB6C1
    LightSalmon = &HFFFFA07A
    LightSeaGreen = &HFF20B2AA
    LightSkyBlue = &HFF87CEFA
    LightSlateGray = &HFF778899
    LightSteelBlue = &HFFB0C4DE
    LightYellow = &HFFFFFFE0
    Lime = &HFF00FF00
    LimeGreen = &HFF32CD32
    Linen = &HFFFAF0E6
    Magenta = &HFFFF00FF
    Maroon = &HFF800000
    MediumAquamarine = &HFF66CDAA
    MediumBlue = &HFF0000CD
    MediumOrchid = &HFFBA55D3
    MediumPurple = &HFF9370DB
    MediumSeaGreen = &HFF3CB371
    MediumSlateBlue = &HFF7B68EE
    MediumSpringGreen = &HFF00FA9A
    MediumTurquoise = &HFF48D1CC
    MediumVioletRed = &HFFC71585
    MidnightBlue = &HFF191970
    MintCream = &HFFF5FFFA
    MistyRose = &HFFFFE4E1
    Moccasin = &HFFFFE4B5
    NavajoWhite = &HFFFFDEAD
    Navy = &HFF000080
    OldLace = &HFFFDF5E6
    Olive = &HFF808000
    OliveDrab = &HFF6B8E23
    Orange = &HFFFFA500
    OrangeRed = &HFFFF4500
    Orchid = &HFFDA70D6
    PaleGoldenrod = &HFFEEE8AA
    PaleGreen = &HFF98FB98
    PaleTurquoise = &HFFAFEEEE
    PaleVioletRed = &HFFDB7093
    PapayaWhip = &HFFFFEFD5
    PeachPuff = &HFFFFDAB9
    Peru = &HFFCD853F
    Pink = &HFFFFC0CB
    Plum = &HFFDDA0DD
    PowderBlue = &HFFB0E0E6
    Purple = &HFF800080
    Red = &HFFFF0000
    RosyBrown = &HFFBC8F8F
    RoyalBlue = &HFF4169E1
    SaddleBrown = &HFF8B4513
    Salmon = &HFFFA8072
    SandyBrown = &HFFF4A460
    SeaGreen = &HFF2E8B57
    SeaShell = &HFFFFF5EE
    Sienna = &HFFA0522D
    Silver = &HFFC0C0C0
    SkyBlue = &HFF87CEEB
    SlateBlue = &HFF6A5ACD
    SlateGray = &HFF708090
    Snow = &HFFFFFAFA
    SpringGreen = &HFF00FF7F
    SteelBlue = &HFF4682B4
    Tan = &HFFD2B48C
    Teal = &HFF008080
    Thistle = &HFFD8BFD8
    Tomato = &HFFFF6347
    Transparent = &HFFFFFF
    Turquoise = &HFF40E0D0
    Violet = &HFFEE82EE
    Wheat = &HFFF5DEB3
    White = &HFFFFFFFF
    WhiteSmoke = &HFFF5F5F5
    Yellow = &HFFFFFF00
    YellowGreen = &HFF9ACD32
End Enum

Public Enum ColorMode
    ColorModeARGB32 = 0
    ColorModeARGB64 = 1
End Enum

Public Enum ColorChannelFlags
    ColorChannelFlagsC = 0
    ColorChannelFlagsM
    ColorChannelFlagsY
    ColorChannelFlagsK
    ColorChannelFlagsLast
End Enum

Public Enum ColorShiftComponents
    AlphaShift = 24
    RedShift = 16
    GreenShift = 8
    BlueShift = 0
End Enum

Public Enum ColorMaskComponents
    AlphaMask = &HFF000000
    RedMask = &HFF0000
    GreenMask = &HFF00
    BlueMask = &HFF
End Enum

'=================================
'String
Public Enum StringFormatFlags
    StringFormatFlagsDirectionRightToLeft = &H1
    StringFormatFlagsDirectionVertical = &H2
    StringFormatFlagsNoFitBlackBox = &H4
    StringFormatFlagsDisplayFormatControl = &H20
    StringFormatFlagsNoFontFallback = &H400
    StringFormatFlagsMeasureTrailingSpaces = &H800
    StringFormatFlagsNoWrap = &H1000
    StringFormatFlagsLineLimit = &H2000

    StringFormatFlagsNoClip = &H4000
End Enum

Public Enum StringTrimming
    StringTrimmingNone = 0
    StringTrimmingCharacter = 1
    StringTrimmingWord = 2
    StringTrimmingEllipsisCharacter = 3
    StringTrimmingEllipsisWord = 4
    StringTrimmingEllipsisPath = 5
End Enum

Public Enum StringDigitSubstitute
    StringDigitSubstituteUser = 0
    StringDigitSubstituteNone = 1
    StringDigitSubstituteNational = 2
    StringDigitSubstituteTraditional = 3
End Enum

'=================================
'Pen / Brush
Public Enum HatchStyle
    HatchStyleHorizontal                   ' 0
    HatchStyleVertical                     ' 1
    HatchStyleForwardDiagonal              ' 2
    HatchStyleBackwardDiagonal             ' 3
    HatchStyleCross                        ' 4
    HatchStyleDiagonalCross                ' 5
    HatchStyle05Percent                    ' 6
    HatchStyle10Percent                    ' 7
    HatchStyle20Percent                    ' 8
    HatchStyle25Percent                    ' 9
    HatchStyle30Percent                    ' 10
    HatchStyle40Percent                    ' 11
    HatchStyle50Percent                    ' 12
    HatchStyle60Percent                    ' 13
    HatchStyle70Percent                    ' 14
    HatchStyle75Percent                    ' 15
    HatchStyle80Percent                    ' 16
    HatchStyle90Percent                    ' 17
    HatchStyleLightDownwardDiagonal        ' 18
    HatchStyleLightUpwardDiagonal          ' 19
    HatchStyleDarkDownwardDiagonal         ' 20
    HatchStyleDarkUpwardDiagonal           ' 21
    HatchStyleWideDownwardDiagonal         ' 22
    HatchStyleWideUpwardDiagonal           ' 23
    HatchStyleLightVertical                ' 24
    HatchStyleLightHorizontal              ' 25
    HatchStyleNarrowVertical               ' 26
    HatchStyleNarrowHorizontal             ' 27
    HatchStyleDarkVertical                 ' 28
    HatchStyleDarkHorizontal               ' 29
    HatchStyleDashedDownwardDiagonal       ' 30
    HatchStyleDashedUpwardDiagonal         ' 31
    HatchStyleDashedHorizontal             ' 32
    HatchStyleDashedVertical               ' 33
    HatchStyleSmallConfetti                ' 34
    HatchStyleLargeConfetti                ' 35
    HatchStyleZigZag                       ' 36
    HatchStyleWave                         ' 37
    HatchStyleDiagonalBrick                ' 38
    HatchStyleHorizontalBrick              ' 39
    HatchStyleWeave                        ' 40
    HatchStylePlaid                        ' 41
    HatchStyleDivot                        ' 42
    HatchStyleDottedGrid                   ' 43
    HatchStyleDottedDiamond                ' 44
    HatchStyleShingle                      ' 45
    HatchStyleTrellis                      ' 46
    HatchStyleSphere                       ' 47
    HatchStyleSmallGrid                    ' 48
    HatchStyleSmallCheckerBoard            ' 49
    HatchStyleLargeCheckerBoard            ' 50
    HatchStyleOutlinedDiamond              ' 51
    HatchStyleSolidDiamond                 ' 52

    HatchStyleTotal
    HatchStyleLargeGrid = HatchStyleCross  ' 4

    HatchStyleMin = HatchStyleHorizontal
    HatchStyleMax = HatchStyleTotal - 1
End Enum

Public Enum PenAlignment
    PenAlignmentCenter = 0
    PenAlignmentInset = 1
End Enum

Public Enum BrushType
    BrushTypeSolidColor = 0
    BrushTypeHatchFill = 1
    BrushTypeTextureFill = 2
    BrushtypeDirGradient = 3
    BrushTypeLinearGradient = 4
End Enum

Public Enum DashStyle
    DashStyleSolid
    DashStyleDash
    DashStyleDot
    DashStyleDashDot
    DashStyleDashDotDot
    DashStyleCustom
End Enum

Public Enum DashCap
    DashCapFlat = 0
    DashCapRound = 2
    DashCapTriangle = 3
End Enum

Public Enum LineCap
    LineCapFlat = 0
    LineCapSquare = 1
    LineCapRound = 2
    LineCapTriangle = 3

    LineCapNoAnchor = &H10         ' corresponds to flat cap
    LineCapSquareAnchor = &H11     ' corresponds to square cap
    LineCapRoundAnchor = &H12      ' corresponds to round cap
    LineCapDiamondAnchor = &H13    ' corresponds to triangle cap
    LineCapArrowAnchor = &H14      ' no correspondence

    LineCapCustom = &HFF           ' custom cap

    LineCapAnchorMask = &HF0        ' mask to check for anchor or not.
End Enum

Public Enum CustomLineCapType
    CustomLineCapTypeDefault = 0
    CustomLineCapTypeAdjustableArrow = 1
End Enum

Public Enum LineJoin
    LineJoinMiter = 0
    LineJoinBevel = 1
    LineJoinRound = 2
    LineJoinMiterClipped = 3
End Enum

Public Enum PenType
    PenTypeSolidColor = BrushTypeSolidColor
    PenTypeHatchFill = BrushTypeHatchFill
    PenTypeTextureFill = BrushTypeTextureFill
    PentypeDirGradient = BrushtypeDirGradient
    PenTypeLinearGradient = BrushTypeLinearGradient
    PenTypeUnknown = -1
End Enum

'=================================
'Meta File
Public Enum MetafileType
    MetafileTypeInvalid            ' Invalid metafile
    MetafileTypeWmf                ' Standard WMF
    MetafileTypeWmfPlaceable       ' Placeable WMF
    MetafileTypeEmf                ' EMF (not EMF+)
    MetafileTypeEmfPlusOnly        ' EMF+ without dual down-level records
    MetafileTypeEmfPlusDual         ' EMF+ with dual down-level records
End Enum

Public Enum emfType
    EmfTypeEmfOnly = MetafileTypeEmf               ' no EMF+  only EMF
    EmfTypeEmfPlusOnly = MetafileTypeEmfPlusOnly   ' no EMF  only EMF+
    EmfTypeEmfPlusDual = MetafileTypeEmfPlusDual   ' both EMF+ and EMF
End Enum

Public Enum ObjectType
    ObjectTypeInvalid
    ObjectTypeBrush
    ObjectTypePen
    ObjecttypeDir
    ObjectTypeRegion
    ObjectTypeImage
    ObjectTypeFont
    ObjectTypeStringFormat
    ObjectTypeImageAttributes
    ObjectTypeCustomLineCap
    #If GdipVersion >= 1.1 Then
    ObjectTypeGraphics
    ObjectTypeMax = ObjectTypeGraphics
    #Else
    ObjectTypeMax = ObjectTypeCustomLineCap
    #End If
    ObjectTypeMin = ObjectTypeBrush
End Enum

Public Enum MetafileFrameUnit
    MetafileFrameUnitPixel = UnitPixel
    MetafileFrameUnitPoint = UnitPoint
    MetafileFrameUnitInch = UnitInch
    MetafileFrameUnitDocument = UnitDocument
    MetafileFrameUnitMillimeter = UnitMillimeter
    MetafileFrameUnitGdi                        ' GDI compatible .01 MM units
End Enum

' Coordinate space identifiers
Public Enum CoordinateSpace
    CoordinateSpaceWorld     ' 0
    CoordinateSpacePage      ' 1
    CoordinateSpaceDevice     ' 2
End Enum

Public Enum EmfPlusRecordType
    WmfRecordTypeSetBkColor = &H10201
    WmfRecordTypeSetBkMode = &H10102
    WmfRecordTypeSetMapMode = &H10103
    WmfRecordTypeSetROP2 = &H10104
    WmfRecordTypeSetRelAbs = &H10105
    WmfRecordTypeSetPolyFillMode = &H10106
    WmfRecordTypeSetStretchBltMode = &H10107
    WmfRecordTypeSetTextCharExtra = &H10108
    WmfRecordTypeSetTextColor = &H10209
    WmfRecordTypeSetTextJustification = &H1020A
    WmfRecordTypeSetWindowOrg = &H1020B
    WmfRecordTypeSetWindowExt = &H1020C
    WmfRecordTypeSetViewportOrg = &H1020D
    WmfRecordTypeSetViewportExt = &H1020E
    WmfRecordTypeOffsetWindowOrg = &H1020F
    WmfRecordTypeScaleWindowExt = &H10410
    WmfRecordTypeOffsetViewportOrg = &H10211
    WmfRecordTypeScaleViewportExt = &H10412
    WmfRecordTypeLineTo = &H10213
    WmfRecordTypeMoveTo = &H10214
    WmfRecordTypeExcludeClipRect = &H10415
    WmfRecordTypeIntersectClipRect = &H10416
    WmfRecordTypeArc = &H10817
    WmfRecordTypeEllipse = &H10418
    WmfRecordTypeFloodFill = &H10419
    WmfRecordTypePie = &H1081A
    WmfRecordTypeRectangle = &H1041B
    WmfRecordTypeRoundRect = &H1061C
    WmfRecordTypePatBlt = &H1061D
    WmfRecordTypeSaveDC = &H1001E
    WmfRecordTypeSetPixel = &H1041F
    WmfRecordTypeOffsetClipRgn = &H10220
    WmfRecordTypeTextOut = &H10521
    WmfRecordTypeBitBlt = &H10922
    WmfRecordTypeStretchBlt = &H10B23
    WmfRecordTypePolygon = &H10324
    WmfRecordTypePolyline = &H10325
    WmfRecordTypeEscape = &H10626
    WmfRecordTypeRestoreDC = &H10127
    WmfRecordTypeFillRegion = &H10228
    WmfRecordTypeFrameRegion = &H10429
    WmfRecordTypeInvertRegion = &H1012A
    WmfRecordTypePaintRegion = &H1012B
    WmfRecordTypeSelectClipRegion = &H1012C
    WmfRecordTypeSelectObject = &H1012D
    WmfRecordTypeSetTextAlign = &H1012E
    WmfRecordTypeDrawText = &H1062F
    WmfRecordTypeChord = &H10830
    WmfRecordTypeSetMapperFlags = &H10231
    WmfRecordTypeExtTextOut = &H10A32
    WmfRecordTypeSetDIBToDev = &H10D33
    WmfRecordTypeSelectPalette = &H10234
    WmfRecordTypeRealizePalette = &H10035
    WmfRecordTypeAnimatePalette = &H10436
    WmfRecordTypeSetPalEntries = &H10037
    WmfRecordTypePolyPolygon = &H10538
    WmfRecordTypeResizePalette = &H10139
    WmfRecordTypeDIBBitBlt = &H10940
    WmfRecordTypeDIBStretchBlt = &H10B41
    WmfRecordTypeDIBCreatePatternBrush = &H10142
    WmfRecordTypeStretchDIB = &H10F43
    WmfRecordTypeExtFloodFill = &H10548
    WmfRecordTypeSetLayout = &H10149
    WmfRecordTypeResetDC = &H1014C
    WmfRecordTypeStartDoc = &H1014D
    WmfRecordTypeStartPage = &H1004F
    WmfRecordTypeEndPage = &H10050
    WmfRecordTypeAbortDoc = &H10052
    WmfRecordTypeEndDoc = &H1005E
    WmfRecordTypeDeleteObject = &H101F0
    WmfRecordTypeCreatePalette = &H100F7
    WmfRecordTypeCreateBrush = &H100F8
    WmfRecordTypeCreatePatternBrush = &H101F9
    WmfRecordTypeCreatePenIndirect = &H102FA
    WmfRecordTypeCreateFontIndirect = &H102FB
    WmfRecordTypeCreateBrushIndirect = &H102FC
    WmfRecordTypeCreateBitmapIndirect = &H102FD
    WmfRecordTypeCreateBitmap = &H106FE
    WmfRecordTypeCreateRegion = &H106FF
    EmfRecordTypeHeader = 1
    EmfRecordTypePolyBezier = 2
    EmfRecordTypePolygon = 3
    EmfRecordTypePolyline = 4
    EmfRecordTypePolyBezierTo = 5
    EmfRecordTypePolyLineTo = 6
    EmfRecordTypePolyPolyline = 7
    EmfRecordTypePolyPolygon = 8
    EmfRecordTypeSetWindowExtEx = 9
    EmfRecordTypeSetWindowOrgEx = 10
    EmfRecordTypeSetViewportExtEx = 11
    EmfRecordTypeSetViewportOrgEx = 12
    EmfRecordTypeSetBrushOrgEx = 13
    EmfRecordTypeEOF = 14
    EmfRecordTypeSetPixelV = 15
    EmfRecordTypeSetMapperFlags = 16
    EmfRecordTypeSetMapMode = 17
    EmfRecordTypeSetBkMode = 18
    EmfRecordTypeSetPolyFillMode = 19
    EmfRecordTypeSetROP2 = 20
    EmfRecordTypeSetStretchBltMode = 21
    EmfRecordTypeSetTextAlign = 22
    EmfRecordTypeSetColorAdjustment = 23
    EmfRecordTypeSetTextColor = 24
    EmfRecordTypeSetBkColor = 25
    EmfRecordTypeOffsetClipRgn = 26
    EmfRecordTypeMoveToEx = 27
    EmfRecordTypeSetMetaRgn = 28
    EmfRecordTypeExcludeClipRect = 29
    EmfRecordTypeIntersectClipRect = 30
    EmfRecordTypeScaleViewportExtEx = 31
    EmfRecordTypeScaleWindowExtEx = 32
    EmfRecordTypeSaveDC = 33
    EmfRecordTypeRestoreDC = 34
    EmfRecordTypeSetWorldTransform = 35
    EmfRecordTypeModifyWorldTransform = 36
    EmfRecordTypeSelectObject = 37
    EmfRecordTypeCreatePen = 38
    EmfRecordTypeCreateBrushIndirect = 39
    EmfRecordTypeDeleteObject = 40
    EmfRecordTypeAngleArc = 41
    EmfRecordTypeEllipse = 42
    EmfRecordTypeRectangle = 43
    EmfRecordTypeRoundRect = 44
    EmfRecordTypeArc = 45
    EmfRecordTypeChord = 46
    EmfRecordTypePie = 47
    EmfRecordTypeSelectPalette = 48
    EmfRecordTypeCreatePalette = 49
    EmfRecordTypeSetPaletteEntries = 50
    EmfRecordTypeResizePalette = 51
    EmfRecordTypeRealizePalette = 52
    EmfRecordTypeExtFloodFill = 53
    EmfRecordTypeLineTo = 54
    EmfRecordTypeArcTo = 55
    EmfRecordTypePolyDraw = 56
    EmfRecordTypeSetArcDirection = 57
    EmfRecordTypeSetMiterLimit = 58
    EmfRecordTypeBeginPath = 59
    EmfRecordTypeEndPath = 60
    EmfRecordTypeCloseFigure = 61
    EmfRecordTypeFillPath = 62
    EmfRecordTypeStrokeAndFillPath = 63
    EmfRecordTypeStrokePath = 64
    EmfRecordTypeFlattenPath = 65
    EmfRecordTypeWidenPath = 66
    EmfRecordTypeSelectClipPath = 67
    EmfRecordTypeAbortPath = 68
    EmfRecordTypeReserved_069 = 69
    EmfRecordTypeGdiComment = 70
    EmfRecordTypeFillRgn = 71
    EmfRecordTypeFrameRgn = 72
    EmfRecordTypeInvertRgn = 73
    EmfRecordTypePaintRgn = 74
    EmfRecordTypeExtSelectClipRgn = 75
    EmfRecordTypeBitBlt = 76
    EmfRecordTypeStretchBlt = 77
    EmfRecordTypeMaskBlt = 78
    EmfRecordTypePlgBlt = 79
    EmfRecordTypeSetDIBitsToDevice = 80
    EmfRecordTypeStretchDIBits = 81
    EmfRecordTypeExtCreateFontIndirect = 82
    EmfRecordTypeExtTextOutA = 83
    EmfRecordTypeExtTextOutW = 84
    EmfRecordTypePolyBezier16 = 85
    EmfRecordTypePolygon16 = 86
    EmfRecordTypePolyline16 = 87
    EmfRecordTypePolyBezierTo16 = 88
    EmfRecordTypePolylineTo16 = 89
    EmfRecordTypePolyPolyline16 = 90
    EmfRecordTypePolyPolygon16 = 91
    EmfRecordTypePolyDraw16 = 92
    EmfRecordTypeCreateMonoBrush = 93
    EmfRecordTypeCreateDIBPatternBrushPt = 94
    EmfRecordTypeExtCreatePen = 95
    EmfRecordTypePolyTextOutA = 96
    EmfRecordTypePolyTextOutW = 97
    EmfRecordTypeSetICMMode = 98
    EmfRecordTypeCreateColorSpace = 99
    EmfRecordTypeSetColorSpace = 100
    EmfRecordTypeDeleteColorSpace = 101
    EmfRecordTypeGLSRecord = 102
    EmfRecordTypeGLSBoundedRecord = 103
    EmfRecordTypePixelFormat = 104
    EmfRecordTypeDrawEscape = 105
    EmfRecordTypeExtEscape = 106
    EmfRecordTypeStartDoc = 107
    EmfRecordTypeSmallTextOut = 108
    EmfRecordTypeForceUFIMapping = 109
    EmfRecordTypeNamedEscape = 110
    EmfRecordTypeColorCorrectPalette = 111
    EmfRecordTypeSetICMProfileA = 112
    EmfRecordTypeSetICMProfileW = 113
    EmfRecordTypeAlphaBlend = 114
    EmfRecordTypeSetLayout = 115
    EmfRecordTypeTransparentBlt = 116
    EmfRecordTypeReserved_117 = 117
    EmfRecordTypeGradientFill = 118
    EmfRecordTypeSetLinkedUFIs = 119
    EmfRecordTypeSetTextJustification = 120
    EmfRecordTypeColorMatchToTargetW = 121
    EmfRecordTypeCreateColorSpaceW = 122
    EmfRecordTypeMax = 122
    EmfRecordTypeMin = 1

    EmfPlusRecordTypeInvalid = 16384    '//GDIP_EMFPLUS_RECORD_BASE
    EmfPlusRecordTypeHeader = 16385
    EmfPlusRecordTypeEndOfFile = 16386
    EmfPlusRecordTypeComment = 16387
    EmfPlusRecordTypeGetDC = 16388
    EmfPlusRecordTypeMultiFormatStart = 16389
    EmfPlusRecordTypeMultiFormatSection = 16390
    EmfPlusRecordTypeMultiFormatEnd = 16391

    EmfPlusRecordTypeObject = 16392

    EmfPlusRecordTypeClear = 16393
    EmfPlusRecordTypeFillRects = 16394
    EmfPlusRecordTypeDrawRects = 16395
    EmfPlusRecordTypeFillPolygon = 16396
    EmfPlusRecordTypeDrawLines = 16397
    EmfPlusRecordTypeFillEllipse = 16398
    EmfPlusRecordTypeDrawEllipse = 16399
    EmfPlusRecordTypeFillPie = 16400
    EmfPlusRecordTypeDrawPie = 16401
    EmfPlusRecordTypeDrawArc = 16402
    EmfPlusRecordTypeFillRegion = 16403
    EmfPlusRecordTypeFillPath = 16404
    EmfPlusRecordTypeDrawPath = 16405
    EmfPlusRecordTypeFillClosedCurve = 16406
    EmfPlusRecordTypeDrawClosedCurve = 16407
    EmfPlusRecordTypeDrawCurve = 16408
    EmfPlusRecordTypeDrawBeziers = 16409
    EmfPlusRecordTypeDrawImage = 16410
    EmfPlusRecordTypeDrawImagePoints = 16411
    EmfPlusRecordTypeDrawString = 16412

    EmfPlusRecordTypeSetRenderingOrigin = 16413
    EmfPlusRecordTypeSetAntiAliasMode = 16414
    EmfPlusRecordTypeSetTextRenderingHint = 16415
    EmfPlusRecordTypeSetTextContrast = 16416
    EmfPlusRecordTypeSetInterpolationMode = 16417
    EmfPlusRecordTypeSetPixelOffsetMode = 16418
    EmfPlusRecordTypeSetCompositingMode = 16419
    EmfPlusRecordTypeSetCompositingQuality = 16420
    EmfPlusRecordTypeSave = 16421
    EmfPlusRecordTypeRestore = 16422
    EmfPlusRecordTypeBeginContainer = 16423
    EmfPlusRecordTypeBeginContainerNoParams = 16424
    EmfPlusRecordTypeEndContainer = 16425
    EmfPlusRecordTypeSetWorldTransform = 16426
    EmfPlusRecordTypeResetWorldTransform = 16427
    EmfPlusRecordTypeMultiplyWorldTransform = 16428
    EmfPlusRecordTypeTranslateWorldTransform = 16429
    EmfPlusRecordTypeScaleWorldTransform = 16430
    EmfPlusRecordTypeRotateWorldTransform = 16431
    EmfPlusRecordTypeSetPageTransform = 16432
    EmfPlusRecordTypeResetClip = 16433
    EmfPlusRecordTypeSetClipRect = 16434
    EmfPlusRecordTypeSetClipPath = 16435
    EmfPlusRecordTypeSetClipRegion = 16436
    EmfPlusRecordTypeOffsetClip = 16437
    EmfPlusRecordTypeDrawDriverString = 16438
    #If GdipVersion >= 1.1 Then
    EmfPlusRecordTypeStrokeFillPath = 16439
    EmfPlusRecordTypeSerializableObject = 16440
    EmfPlusRecordTypeSetTSGraphics = 16441
    EmfPlusRecordTypeSetTSClip = 16442
    EmfPlusRecordTotal = 16443
    #Else
    EmfPlusRecordTotal = 16439
    #End If
    EmfPlusRecordTypeMax = EmfPlusRecordTotal - 1
    EmfPlusRecordTypeMin = EmfPlusRecordTypeHeader
End Enum

'=================================
'Other
Public Enum HotkeyPrefix
    HotkeyPrefixNone = 0
    HotkeyPrefixShow = 1
    HotkeyPrefixHide = 2
End Enum

Public Enum FlushIntention
    FlushIntentionFlush = 0         ' Flush all batched rendering operations
    FlushIntentionSync = 1          ' Flush all batched rendering operations
End Enum

Public Enum EncoderParameterValueType
    EncoderParameterValueTypeByte = 1              ' 8-bit unsigned int
    EncoderParameterValueTypeASCII = 2             ' 8-bit byte containing one 7-bit ASCII
    ' code. NULL terminated.
    EncoderParameterValueTypeShort = 3             ' 16-bit unsigned int
    EncoderParameterValueTypeLong = 4              ' 32-bit unsigned int
    EncoderParameterValueTypeRational = 5          ' Two Longs. The first Long is the
    ' numerator the second Long expresses the
    ' denomintor.
    EncoderParameterValueTypeLongRange = 6         ' Two longs which specify a range of
    ' integer values. The first Long specifies
    ' the lower end and the second one
    ' specifies the higher end. All values
    ' are inclusive at both ends
    EncoderParameterValueTypeUndefined = 7         ' 8-bit byte that can take any value
    ' depending on field definition
    EncoderParameterValueTypeRationalRange = 8     ' Two Rationals. The first Rational
    ' specifies the lower end and the second
    ' specifies the higher end. All values
    ' are inclusive at both ends
    #If GdipVersion >= 1.1 Then
    EncoderParameterValueTypePointer = 9       ' a pointer to a parameter defined data.
    #End If
End Enum

Public Enum EncoderValue
    EncoderValueColorTypeCMYK = 0
    EncoderValueColorTypeYCCK
    EncoderValueCompressionLZW
    EncoderValueCompressionCCITT3
    EncoderValueCompressionCCITT4
    EncoderValueCompressionRle
    EncoderValueCompressionNone
    EncoderValueScanMethodInterlaced
    EncoderValueScanMethodNonInterlaced
    EncoderValueVersionGif87
    EncoderValueVersionGif89
    EncoderValueRenderProgressive
    EncoderValueRenderNonProgressive
    EncoderValueTransformRotate90
    EncoderValueTransformRotate180
    EncoderValueTransformRotate270
    EncoderValueTransformFlipHorizontal
    EncoderValueTransformFlipVertical
    EncoderValueMultiFrame
    EncoderValueLastFrame
    EncoderValueFlush
    EncoderValueFrameDimensionTime
    EncoderValueFrameDimensionResolution
    EncoderValueFrameDimensionPage
    #If GdipVersion >= 1.1 Then
    EncoderValueColorTypeGray
    EncoderValueColorTypeRGB
    #End If
End Enum

#If GdipVersion >= 1.1 Then
Public Enum ConvertToEmfPlusFlags
    ConvertToEmfPlusFlagsDefault = 0
    ConvertToEmfPlusFlagsRopUsed = 1
    ConvertToEmfPlusFlagsText = 2
    ConvertToEmfPlusFlagsInvalidRecord = 4
End Enum
#End If

Public Enum DebugEventLevel
    DebugEventLevelFatal = 0
    DebugEventLevelWarning
End Enum

Public Enum GpTestControlEnum
    TestControlForceBilinear = 0
    TestControlNoICM = 1
    TestControlGetBuildNumber = 2
End Enum

Public Declare Function GdipCreateFromHDC2 _
                         Lib "gdiplus" (ByVal hDC As Long, _
                                        ByVal hDevice As Long, _
                                        graphics As Long) As GpStatus
Public Declare Function GdipCreateFromHWNDICM _
                         Lib "gdiplus" (ByVal hwnd As Long, _
                                        graphics As Long) As GpStatus

Public Declare Function GdipEnumerateMetafileDestPoint _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoint As POINTF, _
                                        lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointI _
                         Lib "gdiplus" (graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoint As POINTL, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destRect As RECTF, _
                                        lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destRect As RECTL, _
                                        lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPoints _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoint As POINTF, _
                                        ByVal Count As Long, _
                                        lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileDestPointsI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoint As POINTL, _
                                        ByVal Count As Long, _
                                        lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoint _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoint As POINTF, _
                                        srcRect As RECTF, _
                                        ByVal srcUnit As GpUnit, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoint As POINTL, _
                                        srcRect As RECTL, _
                                        ByVal srcUnit As GpUnit, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRect _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destRect As RECTF, _
                                        srcRect As RECTF, _
                                        ByVal srcUnit As GpUnit, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestRectI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destRect As RECTL, _
                                        srcRect As RECTL, _
                                        ByVal srcUnit As GpUnit, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoints As POINTF, _
                                        ByVal Count As Long, _
                                        srcRect As RECTF, _
                                        ByVal srcUnit As GpUnit, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal metafile As Long, _
                                        destPoints As POINTL, _
                                        ByVal Count As Long, _
                                        srcRect As RECTL, _
                                        ByVal srcUnit As GpUnit, _
                                        ByVal lpEnumerateMetafileProc As Long, _
                                        ByVal callbackData As Long, _
                                        ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPoints_ _
                         Lib "gdiplus" _
                             Alias "GdipEnumerateMetafileSrcRectDestPoints" _
                             (ByVal graphics As Long, _
                              ByVal metafile As Long, _
                              destPoints As Any, _
                              ByVal Count As Long, _
                              srcRect As RECTF, _
                              ByVal srcUnit As GpUnit, _
                              ByVal lpEnumerateMetafileProc As Long, _
                              ByVal callbackData As Long, _
                              ByVal imageAttributes As Long) As GpStatus
Public Declare Function GdipEnumerateMetafileSrcRectDestPointsI_ _
                         Lib "gdiplus" _
                             Alias "GdipEnumerateMetafileSrcRectDestPointsI" _
                             (ByVal graphics As Long, _
                              ByVal metafile As Long, _
                              destPoints As Any, _
                              ByVal Count As Long, _
                              srcRect As RECTL, _
                              ByVal srcUnit As GpUnit, _
                              ByVal lpEnumerateMetafileProc As Long, _
                              ByVal callbackData As Long, _
                              ByVal imageAttributes As Long) As GpStatus

Public Declare Function GdipPlayMetafileRecord _
                         Lib "gdiplus" (ByVal metafile As Long, _
                                        ByVal recordType As EmfPlusRecordType, _
                                        ByVal Flags As Long, _
                                        ByVal dataSize As Long, _
                                        byteData As Any) As GpStatus

Public Declare Function GdipGetMetafileHeaderFromWmf _
                         Lib "gdiplus" (ByVal hWmf As Long, _
                                        WmfPlaceableFileHdr As WmfPlaceableFileHeader, _
                                        header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromEmf _
                         Lib "gdiplus" (ByVal hEmf As Long, _
                                        header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromFile _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromStream _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        header As MetafileHeader) As GpStatus
Public Declare Function GdipGetMetafileHeaderFromMetafile _
                         Lib "gdiplus" (ByVal metafile As Long, _
                                        header As MetafileHeader) As GpStatus
Public Declare Function GdipGetHemfFromMetafile _
                         Lib "gdiplus" (ByVal metafile As Long, _
                                        hEmf As Long) As GpStatus
Public Declare Function GdipCreateStreamOnFile Lib "gdiplus" (ByVal fileName As Long, ByVal access As Long, stream As Any) As GpStatus
Public Declare Function GdipCreateMetafileFromWmf _
                         Lib "gdiplus" (ByVal hWmf As Long, _
                                        ByVal bDeleteWmf As Long, _
                                        WmfPlaceableFileHdr As WmfPlaceableFileHeader, _
                                        ByVal metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromEmf _
                         Lib "gdiplus" (ByVal hEmf As Long, _
                                        ByVal bDeleteEmf As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromFile _
                         Lib "gdiplus" (ByVal file As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromWmfFile _
                         Lib "gdiplus" (ByVal file As Long, _
                                        WmfPlaceableFileHdr As WmfPlaceableFileHeader, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipCreateMetafileFromStream _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafile _
                         Lib "gdiplus" (ByVal referenceHdc As Long, _
                                        etype As emfType, _
                                        frameRect As RECTF, _
                                        ByVal frameUnit As MetafileFrameUnit, _
                                        ByVal Description As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileI _
                         Lib "gdiplus" (ByVal referenceHdc As Long, _
                                        etype As emfType, _
                                        frameRect As RECTL, _
                                        ByVal frameUnit As MetafileFrameUnit, _
                                        ByVal Description As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileName _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        ByVal referenceHdc As Long, _
                                        etype As emfType, _
                                        frameRect As RECTF, _
                                        ByVal frameUnit As MetafileFrameUnit, _
                                        ByVal Description As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileFileNameI _
                         Lib "gdiplus" (ByVal fileName As Long, _
                                        ByVal referenceHdc As Long, _
                                        etype As emfType, _
                                        frameRect As RECTL, _
                                        ByVal frameUnit As MetafileFrameUnit, _
                                        ByVal Description As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStream _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        ByVal referenceHdc As Long, _
                                        etype As emfType, _
                                        frameRect As RECTF, _
                                        ByVal frameUnit As MetafileFrameUnit, _
                                        ByVal Description As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipRecordMetafileStreamI _
                         Lib "gdiplus" (ByVal stream As Any, _
                                        ByVal referenceHdc As Long, _
                                        etype As emfType, _
                                        frameRect As RECTL, _
                                        ByVal frameUnit As MetafileFrameUnit, _
                                        ByVal Description As Long, _
                                        metafile As Long) As GpStatus
Public Declare Function GdipSetMetafileDownLevelRasterizationLimit _
                         Lib "gdiplus" (ByVal metafile As Long, _
                                        ByVal metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetMetafileDownLevelRasterizationLimit _
                         Lib "gdiplus" (ByVal metafile As Long, _
                                        metafileRasterizationLimitDpi As Long) As GpStatus
Public Declare Function GdipGetImageDecodersSize _
                         Lib "gdiplus" (numDecoders As Long, _
                                        Size As Long) As GpStatus
Public Declare Function GdipSetImageAttributesCachedBackground _
                         Lib "gdiplus" (ByVal imageattr As Long, _
                                        ByVal enableFlag As Long) As GpStatus
Public Declare Function GdipTestControl _
                         Lib "gdiplus" (ByVal control As GpTestControlEnum, _
                                        param As Any) As GpStatus
#If GdipVersion >= 1.1 Then

    Public Declare Function GdipConvertToEmfPlus _
                         Lib "gdiplus" (ByVal refGraphics As Long, _
                                        conversionFailureFlag As Long, _
                                        ByVal emfType As emfType, _
                                        ByVal Description As Long, _
                                        ByVal out_metafile As Long) As GpStatus
    Public Declare Function GdipConvertToEmfPlusToFile _
                         Lib "gdiplus" (ByVal refGraphics As Long, _
                                        ByVal metafile As Long, _
                                        conversionFailureFlag As Long, _
                                        ByVal fileName As Long, _
                                        ByVal emfType As emfType, _
                                        ByVal Description As Long, _
                                        out_metafile As Long) As GpStatus
    Public Declare Function GdipConvertToEmfPlusToStream _
                         Lib "gdiplus" (ByVal refGraphics As Long, _
                                        ByVal metafile As Long, _
                                        conversionFailureFlag As Long, _
                                        stream As Any, _
                                        ByVal emfType As emfType, _
                                        ByVal Description As Long, _
                                        out_metafile As Long) As GpStatus
#End If

Public Declare Function GdipFlush _
                         Lib "gdiplus" (ByVal graphics As Long, _
                                        ByVal intention As FlushIntention) As GpStatus
Public Declare Function GdipAlloc Lib "gdiplus" (ByVal Size As Long) As Long
Public Declare Sub GdipFree Lib "gdiplus" (ByVal ptr As Long)

'===================================================================================
'  �������� / ��������
'===================================================================================

Public Declare Function GdiplusStartup _
                         Lib "gdiplus" (Token As Long, _
                                        inputbuf As GdiplusStartupInput, _
                                        Optional ByVal outputbuf As Long = 0) As GpStatus
Public Declare Function GdiplusShutdown Lib "gdiplus" (ByVal Token As Long) As GpStatus

Public Type GdiplusStartupInput
    GdiplusVersion As Long
    DebugEventCallback As Long
    SuppressBackgroundThread As Long
    SuppressExternalCodecs As Long
End Type

Public Enum GpStatus
    Ok = 0
    GenericError = 1
    InvalidParameter = 2
    OutOfMemory = 3
    ObjectBusy = 4
    InsufficientBuffer = 5
    NotImplemented = 6
    Win32Error = 7
    WrongState = 8
    Aborted = 9
    FileNotFound = 10
    ValueOverflow = 11
    AccessDenied = 12
    UnknownImageFormat = 13
    FontFamilyNotFound = 14
    FontStyleNotFound = 15
    NotTrueTypeFont = 16
    UnsupportedGdiplusVersion = 17
    GdiplusNotInitialized = 18
    PropertyNotFound = 19
    PropertyNotSupported = 20
    #If GdipVersion >= 1.1 Then
    ProfileNotFound = 21
    #End If
End Enum

Private Declare Function CLSIDFromString _
                          Lib "ole32.dll" (ByVal lpszProgID As Long, _
                                           pCLSID As CLSID) As Long

Public Enum GdipImageType
    Bmp
    EMF
    WMF
    Jpg
    Png
    Gif
    TIF
    ICO
End Enum

Public Const ImageEncoderSuffix As String = "-1A04-11D3-9A73-0000F81EF32E}"
Public Const ImageEncoderBMP As String = "{557CF400" & ImageEncoderSuffix
Public Const ImageEncoderJPG As String = "{557CF401" & ImageEncoderSuffix
Public Const ImageEncoderGIF As String = "{557CF402" & ImageEncoderSuffix
Public Const ImageEncoderEMF As String = "{557CF403" & ImageEncoderSuffix
Public Const ImageEncoderWMF As String = "{557CF404" & ImageEncoderSuffix
Public Const ImageEncoderTIF As String = "{557CF405" & ImageEncoderSuffix
Public Const ImageEncoderPNG As String = "{557CF406" & ImageEncoderSuffix
Public Const ImageEncoderICO As String = "{557CF407" & ImageEncoderSuffix
Public Const EncoderCompression As String = "{E09D739D-CCD4-44EE-8EBA-3FBF8BE4FC58}"
Public Const EncoderColorDepth As String = "{66087055-AD66-4C7C-9A18-38A2310B8337}"
Public Const EncoderScanMethod As String = "{3A4E2661-3109-4E56-8536-42C156E7DCFA}"
Public Const EncoderVersion As String = "{24D18C76-814A-41A4-BF53-1C219CCCF797}"
Public Const EncoderRenderMethod As String = "{6D42C53A-229A-4825-8BB7-5C99E2B9A8B8}"
Public Const EncoderQuality As String = "{1D5BE4B5-FA4A-452D-9CDD-5DB35105E7EB}"
Public Const EncoderTransformation As String = "{8D0EB2D1-A58E-4EA8-AA14-108074B7B6F9}"
Public Const EncoderLuminanceTable As String = "{EDB33BCE-0266-4A77-B904-27216099E717}"
Public Const EncoderChrominanceTable As String = "{F2E455DC-09B3-4316-8260-676ADA32481C}"
Public Const EncoderSaveFlag As String = "{292266FC-AC40-47BF-8CFC-A85B89A655DE}"
#If GdipVersion >= 1.1 Then
    Public Const EncoderColorSpace As String = "{AE7A62A0-EE2C-49D8-9D07-1BA8A927596E}"
    Public Const EncoderImageItems As String = "{63875E13-1F1D-45AB-9195-A29B6066A650}"
    Public Const EncoderSaveAsCMYK As String = "{A219BBC9-0A9D-4005-A3EE-3A421B8BB06C}"
#End If

Public ZeroPointF As POINTF, ZeroPointL As POINTL
Attribute ZeroPointL.VB_VarUserMemId = 1073741824
Public ZeroRectF As RECTF, ZeroRectL As RECTL
Attribute ZeroRectL.VB_VarUserMemId = 1073741826

Dim mToken As Long

Dim Pens() As Long, PenCount As Long
Attribute PenCount.VB_VarUserMemId = 1073741829
Dim Brushes() As Long, BrushCount As Long
Attribute BrushCount.VB_VarUserMemId = 1073741831
Dim StrFormats() As Long, StrFormatCount As Long
Attribute StrFormatCount.VB_VarUserMemId = 1073741833
Dim Matrixes() As Long, MatrixCount As Long
Attribute MatrixCount.VB_VarUserMemId = 1073741835

Public Function DeleteObjects()
    Dim i As Long

    For i = 1 To PenCount: GdipDeletePen Pens(i): Next
    For i = 1 To BrushCount: GdipDeleteBrush Brushes(i): Next
    For i = 1 To StrFormatCount: GdipDeleteStringFormat StrFormats(i): Next
    For i = 1 To MatrixCount: GdipDeleteMatrix Matrixes(i): Next
    PenCount = 0
    BrushCount = 0
    StrFormatCount = 0
    MatrixCount = 0
End Function

Public Function NewPen(ByVal Color As Long, ByVal Width As Single) As Long
    PenCount = PenCount + 1
    ReDim Preserve Pens(PenCount)

    GdipCreatePen1 Color, Width, UnitPixel, Pens(PenCount)
    NewPen = Pens(PenCount)
End Function

Public Function NewBrush(ByVal Color As Long) As Long
    BrushCount = BrushCount + 1
    ReDim Preserve Brushes(BrushCount)

    GdipCreateSolidFill Color, Brushes(BrushCount)
    NewBrush = Brushes(BrushCount)
End Function

Public Function NewStringFormat(ByVal Align As StringAlignment) As Long
    StrFormatCount = StrFormatCount + 1
    ReDim Preserve StrFormats(StrFormatCount)

    GdipCreateStringFormat 0, 0, StrFormats(StrFormatCount)
    GdipSetStringFormatAlign StrFormats(StrFormatCount), Align
    NewStringFormat = StrFormats(StrFormatCount)
End Function

Public Function NewMatrix(ByVal m11 As Single, _
                          ByVal m12 As Single, _
                          ByVal m21 As Single, _
                          ByVal m22 As Single, _
                          ByVal dx As Single, _
                          ByVal dy As Single) As Long

    MatrixCount = MatrixCount + 1
    ReDim Preserve Matrixes(MatrixCount)

    GdipCreateMatrix Matrixes(MatrixCount)
    GdipSetMatrixElements Matrixes(MatrixCount), m11, m12, m21, m22, dx, dy
    NewMatrix = Matrixes(MatrixCount)
End Function

Public Function NewRectF(ByVal Left As Single, _
                         ByVal Top As Single, _
                         ByVal Width As Single, _
                         ByVal Height As Single) As RECTF

    With NewRectF
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
End Function

Public Function NewRectL(ByVal Left As Single, _
                         ByVal Top As Long, _
                         ByVal Width As Long, _
                         ByVal Height As Long) As RECTL

    With NewRectL
        .Left = Left
        .Top = Top
        .Right = Width
        .Bottom = Height
    End With
End Function

Public Function NewPointF(ByVal X As Single, _
                          ByVal Y As Single) As POINTF

    NewPointF.X = X
    NewPointF.Y = Y
End Function

Public Function NewPointL(ByVal X As Single, _
                          ByVal Y As Single) As POINTL

    NewPointL.X = X
    NewPointL.Y = Y
End Function

Public Function NewPointsFPtr(ParamArray ptXY()) As Long
    If (UBound(ptXY) And 1) = 0 Then GoTo ErrHandle

    Dim ret() As POINTF, i As Long
    ReDim ret(0 To UBound(ptXY) \ 2)

    For i = 0 To UBound(ptXY) Step 2
        ret(i \ 2).X = ptXY(i)
        ret(i \ 2).Y = ptXY(i + 1)
    Next

    NewPointsFPtr = VarPtr(ret(0))

    Exit Function
ErrHandle:
    NewPointsFPtr = 0
End Function

Public Function NewPointsLPtr(ParamArray ptXY()) As Long
    If (UBound(ptXY) And 1) = 0 Then GoTo ErrHandle

    Dim ret() As POINTL, i As Long
    ReDim ret(0 To UBound(ptXY) \ 2)

    For i = 0 To UBound(ptXY) Step 2
        ret(i \ 2).X = ptXY(i)
        ret(i \ 2).Y = ptXY(i + 1)
    Next

    NewPointsLPtr = VarPtr(ret(0))

    Exit Function
ErrHandle:
    NewPointsLPtr = 0
End Function

Public Function NewColors(ParamArray colors()) As Long()
    Dim ret() As Long, i As Long

    ReDim ret(UBound(colors))
    For i = 0 To UBound(colors)
        ret(i) = colors(i)
    Next

    NewColors = ret
End Function

Public Function InitGDIPlus(Optional OnErrorMsgbox, _
                            Optional ByVal OnErrorEnd As Boolean = True) As GpStatus

    If mToken <> 0 Then
        Debug.Print "InitGDIPlus> GdiPlus�ѱ���ʼ��"
        Exit Function
    End If

    Dim uInput As GdiplusStartupInput
    Dim ret As GpStatus

    uInput.GdiplusVersion = 1
    ret = GdiplusStartup(mToken, uInput)

    If ret <> Ok Then
        If Not IsMissing(OnErrorMsgbox) Then MsgBox OnErrorMsgbox
        If OnErrorEnd Then End
    End If

    InitGDIPlus = ret
End Function

Public Sub TerminateGDIPlus()
    If mToken = 0 Then
        Debug.Print "TerminateGDIPlus> GdiPlus�ѱ�����"
        Exit Sub
    End If

    DeleteObjects
    GdiplusShutdown mToken

    mToken = 0
End Sub

Public Function InitGDIPlusTo(ByRef Token As Long, _
                              Optional OnErrorMsgbox, _
                              Optional ByVal OnErrorEnd As Boolean = True) As GpStatus

    If Token <> 0 Then
        Debug.Print "InitGDIPlusTo> GdiPlus�ѱ���ʼ��"
        Exit Function
    End If

    Dim uInput As GdiplusStartupInput
    Dim ret As GpStatus

    uInput.GdiplusVersion = 1
    ret = GdiplusStartup(Token, uInput)

    If ret <> Ok Then
        If Not IsMissing(OnErrorMsgbox) Then MsgBox OnErrorMsgbox
        If OnErrorEnd Then End
    End If

    InitGDIPlusTo = ret
End Function

Public Sub TerminateGDIPlusFrom(ByVal Token As Long)
    If Token = 0 Then
        Debug.Print "TerminateGDIPlusFrom> GdiPlus�ѱ�����"
        Exit Sub
    End If

    DeleteObjects
    GdiplusShutdown Token

    Token = 0
End Sub

#If GdipVersion >= 1.1 Then

Public Sub GdipCreateEffect2(ByVal EffectType As GdipEffectType, effect As Long)
    Select Case EffectType
    Case GdipEffectType.Blur: GdipCreateEffect &H633C80A4, &H482B1843, &H28BEF29E, &HD4FDC534, effect    'CLSIDFromString StrPtr(BlurEffectGuid), GetEffectClsid
    Case GdipEffectType.BrightnessContrast: GdipCreateEffect &HD3A1DBE1, &H4C178EC4, &H97EA4C9F, &H3D341CAD, effect    'CLSIDFromString StrPtr(BrightnessContrastEffectGuid), GetEffectClsid
    Case GdipEffectType.ColorBalance: GdipCreateEffect &H537E597D, &H48DA251E, &HCA296496, &HF8706B49, effect    'CLSIDFromString StrPtr(ColorBalanceEffectGuid), GetEffectClsid
    Case GdipEffectType.ColorCurve: GdipCreateEffect &HDD6A0022, &H4A6758E4, &H8ED49B9D, &H3DA581B8, effect    'CLSIDFromString StrPtr(ColorCurveEffectGuid), GetEffectClsid
    Case GdipEffectType.ColorLookupTable: GdipCreateEffect &HA7CE72A9, &H40D70F7F, &HC0D0CCB3, &H12325C2D, effect    'CLSIDFromString StrPtr(ColorLookupTableEffectGuid), GetEffectClsid
    Case GdipEffectType.ColorMatrix: GdipCreateEffect &H718F2615, &H40E37933, &H685F11A5, &H74DD14FE, effect    'CLSIDFromString StrPtr(ColorMatrixEffectGuid), GetEffectClsid
    Case GdipEffectType.HueSaturationLightness: GdipCreateEffect &H8B2DD6C3, &H4D87EB07, &H871F0A5, &H5F9C6AE2, effect    'CLSIDFromString StrPtr(HueSaturationLightnessEffectGuid), GetEffectClsid
    Case GdipEffectType.Levels: GdipCreateEffect &H99C354EC, &H4F3A2A31, &HA817348C, &H253AB303, effect    'CLSIDFromString StrPtr(LevelsEffectGuid), GetEffectClsid
    Case GdipEffectType.RedEyeCorrection: GdipCreateEffect &H74D29D05, &H426669A4, &HC53C4995, &H32B63628, effect    'CLSIDFromString StrPtr(RedEyeCorrectionEffectGuid), GetEffectClsid
    Case GdipEffectType.Sharpen: GdipCreateEffect &H63CBF3EE, &H402CC526, &HC562718F, &H4251BF40, effect    'CLSIDFromString StrPtr(SharpenEffectGuid), GetEffectClsid
    Case GdipEffectType.Tint: GdipCreateEffect &H1077AF00, &H44412848, &HAD448994, &H2C7A2D4C, effect    'CLSIDFromString StrPtr(TintEffectGuid), GetEffectClsid
    End Select
End Sub

Public Function GetAddress(ByVal lngAddr As Long) As Long
    GetAddress = lngAddr
End Function

#End If

Public Function GetImageEncoderClsid(ByVal ImageType As GdipImageType) As CLSID
    Select Case ImageType
    Case GdipImageType.Png: CLSIDFromString StrPtr(ImageEncoderPNG), GetImageEncoderClsid
    Case GdipImageType.Jpg: CLSIDFromString StrPtr(ImageEncoderJPG), GetImageEncoderClsid
    Case GdipImageType.Gif: CLSIDFromString StrPtr(ImageEncoderGIF), GetImageEncoderClsid
    Case GdipImageType.Bmp: CLSIDFromString StrPtr(ImageEncoderBMP), GetImageEncoderClsid
    Case GdipImageType.ICO: CLSIDFromString StrPtr(ImageEncoderICO), GetImageEncoderClsid
    Case GdipImageType.EMF: CLSIDFromString StrPtr(ImageEncoderEMF), GetImageEncoderClsid
    Case GdipImageType.WMF: CLSIDFromString StrPtr(ImageEncoderWMF), GetImageEncoderClsid
    Case GdipImageType.TIF: CLSIDFromString StrPtr(ImageEncoderTIF), GetImageEncoderClsid
    End Select
End Function

Public Function SaveImageToPNG(ByVal Image As Long, ByVal path As String) As GpStatus
    SaveImageToPNG = GdipSaveImageToFile(Image, StrPtr(path), GetImageEncoderClsid( _
                                                              Png), ByVal 0)
End Function

Public Function SaveImageToJPG(ByVal Image As Long, _
                               ByVal path As String, _
                               ByVal Quality As Long) As GpStatus

    Dim Params As EncoderParameters

    Params.Count = 1
    CLSIDFromString StrPtr(EncoderQuality), Params.Parameter.GUID
    Params.Parameter.NumberOfValues = 1
    Params.Parameter.Type = 4
    Params.Parameter.Value = VarPtr(Quality)

    SaveImageToJPG = GdipSaveImageToFile(Image, StrPtr(path), GetImageEncoderClsid( _
                                                              Jpg), Params)
End Function

Public Function SaveImageToGIF(ByVal Image As Long, ByVal path As String) As GpStatus
    SaveImageToGIF = GdipSaveImageToFile(Image, StrPtr(path), GetImageEncoderClsid( _
                                                              Gif), ByVal 0)
End Function

Public Function SaveImageToBMP(ByVal Image As Long, ByVal path As String) As GpStatus
    SaveImageToBMP = GdipSaveImageToFile(Image, StrPtr(path), GetImageEncoderClsid( _
                                                              Bmp), ByVal 0)
End Function

Public Function CreateBitmap(ByRef bitmap As Long, _
                             ByVal Width As Long, _
                             ByVal Height As Long, _
                             Optional ByVal PixelFormat As GpPixelFormat = PixelFormat32bppARGB) As GpStatus

    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat, ByVal 0, bitmap
End Function

Public Function CreateBitmapWithGraphics(ByRef bitmap As Long, _
                                         ByRef graphics As Long, _
                                         ByVal Width As Long, _
                                         ByVal Height As Long, _
                                         Optional ByVal PixelFormat As GpPixelFormat = PixelFormat32bppARGB) As GpStatus

    GdipCreateBitmapFromScan0 Width, Height, 0, PixelFormat, ByVal 0, bitmap
    GdipGetImageGraphicsContext bitmap, graphics
End Function


