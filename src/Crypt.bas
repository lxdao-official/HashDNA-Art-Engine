Attribute VB_Name = "Crypt"
Option Explicit

Public Const ALG_CLASS_HASH = 32768
Public Const ALG_SID_MD2 = 1
Public Const ALG_SID_MD4 = 2
Public Const ALG_SID_MD5 = 3
Public Const ALG_SID_SHA1 = 4
Public Const ALG_TYPE_ANY = 0

Public Const CRYPT_NEWKEYSET = 8

Public Const HP_HASHVAL = 2
Public Const HP_HASHSIZE = 4

Public Const PROV_RSA_FULL = 1

Public Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As Long, ByVal Algid As Long, ByVal hKey As Long, ByVal dwFlags As Long, ByRef phHash As Long) As Long
Public Declare Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As Long) As Long
Public Declare Function CryptGetHashParam Lib "advapi32.dll" (ByVal hHash As Long, ByVal dwParam As Long, pbData As Any, pdwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptHashData Lib "advapi32.dll" (ByVal hHash As Long, pbData As Any, ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Public Declare Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long

Public Function CryptStr(ByVal Str As String, Optional ByVal CryptMode As String = "md5") As String
    CryptStr = CryptByte(StrConv(Str, vbFromUnicode), CryptMode)
End Function

Public Function CryptByte(ByRef Buffer() As Byte, Optional ByVal CryptMode As String = "md5") As String
    Dim Algorithm As Long, hCtx As Long, lRes As Long, hHash As Long, lLen As Long, abHash() As Byte
    Select Case LCase(CryptMode)
    Case "md2"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    Case "md4"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    Case "md5"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    Case "sha1"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
    Case Else
        Exit Function
    End Select
    If CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0) = 0 Then
        CryptAcquireContext hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET
    End If
    CryptCreateHash hCtx, Algorithm, 0, 0, hHash
    If UBound(Buffer) >= 0 Then CryptHashData hHash, Buffer(0), UBound(Buffer) + 1, 0
    CryptGetHashParam hHash, HP_HASHSIZE, lLen, 4, 0
    ReDim abHash(lLen - 1)
    CryptGetHashParam hHash, HP_HASHVAL, abHash(0), lLen, 0
    CryptDestroyHash hHash
    CryptReleaseContext hCtx, 0
    For lRes = 0 To UBound(abHash)
        CryptByte = CryptByte & Right("0" & Hex(abHash(lRes)), 2)
    Next
    Erase abHash
End Function

Public Function CryptFile(ByVal fileName As String, Optional ByVal CryptMode As String = "sha1", Optional ByVal BlockSize As Long = 327680) As String
    Dim Algorithm As Long, hCtx As Long, lRes As Long, hHash As Long, lLen As Long, abHash() As Byte, Data() As Byte, fn As Integer, FileSize As Long
    Select Case LCase(CryptMode)
    Case "md2"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD2
    Case "md4"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD4
    Case "md5"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_MD5
    Case "sha1"
        Algorithm = ALG_CLASS_HASH Or ALG_TYPE_ANY Or ALG_SID_SHA1
    Case Else
        Exit Function
    End Select
    If CryptAcquireContext(hCtx, vbNullString, vbNullString, PROV_RSA_FULL, 0) = 0 Then
        CryptAcquireContext hCtx, vbNullString, vbNullString, PROV_RSA_FULL, CRYPT_NEWKEYSET
    End If
    CryptCreateHash hCtx, Algorithm, 0, 0, hHash
    fn = FreeFile
    ReDim Data(BlockSize - 1)
    Open fileName For Binary As fn
    FileSize = LOF(fn)
    Do While FileSize > 0
        Get fn, , Data
        CryptHashData hHash, Data(0), IIf(FileSize > BlockSize, BlockSize, FileSize), 0
        FileSize = FileSize - BlockSize
    Loop
    Erase Data
    Close fn
    CryptGetHashParam hHash, HP_HASHSIZE, lLen, 4, 0
    ReDim abHash(lLen - 1)
    CryptGetHashParam hHash, HP_HASHVAL, abHash(0), lLen, 0
    CryptDestroyHash hHash
    CryptReleaseContext hCtx, 0
    For lRes = 0 To UBound(abHash)
        CryptFile = CryptFile & Right("0" & Hex(abHash(lRes)), 2)
    Next
    Erase abHash
End Function
