VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'CryptAPIを利用したAES暗号化/復号クラス

Option Explicit

'CryptAPI用定数定義
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const ALG_TYPE_BLOCK As Long = 1536
Private Const ALG_CLASS_DATA_ENCRYPT As Long = 24576
Private Const ALG_SID_AES_128 As Long = 14
Private Const ALG_SID_AES_192 As Long = 15
Private Const ALG_SID_AES_256 As Long = 16
Private Const PROV_RSA_AES As Long = 24
Private Const CALG_AES_128 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_128
Private Const CALG_AES_192 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_192
Private Const CALG_AES_256 As Long = ALG_CLASS_DATA_ENCRYPT Or ALG_TYPE_BLOCK Or ALG_SID_AES_256
Private Const KP_IV As Long = 1
Private Const KP_PADDING As Long = 3
Private Const KP_MODE As Long = 4
Private Const PKCS5_PADDING As Long = 1
Private Const CRYPT_MODE_CBC As Long = 1
Private Const PLAINTEXTKEYBLOB As Long = 8
Private Const CUR_BLOB_VERSION As Long = 2

'WindowsAPI用定数定義
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER As Long = &H100
Private Const FORMAT_MESSAGE_IGNORE_INSERTS As Long = &H200
Private Const FORMAT_MESSAGE_FROM_STRING As Long = &H400
Private Const FORMAT_MESSAGE_FROM_HMODULE As Long = &H800
Private Const FORMAT_MESSAGE_FROM_SYSTEM As Long = &H1000
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY As Long = 8192
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK As Long = 255
Private Const LANG_NEUTRAL As Long = &H0
Private Const SUBLANG_DEFAULT As Long = &H1


'CryptoAPI定義
#If VBA7 And Win64 Then

Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As LongPtr, ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptImportKey Lib "advapi32.dll" ( _
    ByVal hProv As LongPtr, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As LongPtr) As Long
Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function CryptSetKeyParam Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal dwParam As Long, ByRef pbData As Any, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare PtrSafe Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As LongPtr, ByVal hHash As LongPtr, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

#Else

Private Declare Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" ( _
    ByRef phProv As Long, ByVal pszContainer As String, ByVal pszProvider As String, _
    ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptReleaseContext Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByVal dwFlags As Long) As Long
Private Declare Function CryptImportKey Lib "advapi32.dll" ( _
    ByVal hProv As Long, ByRef pbData As Any, ByVal dwDataLen As Long, _
    ByVal hPubKey As Long, ByVal dwFlags As Long, ByRef phKey As Long) As Long
Private Declare Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function CryptSetKeyParam Lib "advapi32.dll" ( _
    ByVal hKey As Long, ByVal dwParam As Long, ByRef pbData As Any, _
    ByVal dwFlags As Long) As Long
Private Declare Function CryptDecrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare Function CryptEncrypt Lib "advapi32.dll" ( _
    ByVal hKey As Long, ByVal hHash As Long, ByVal Final As Long, _
    ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long, _
    ByVal dwBufLen As Long) As Long

#End If


'WindowsAPI定義
#If VBA7 And Win64 Then

Private Declare PtrSafe Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare PtrSafe Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" ( _
    ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByRef lpBuffer As LongPtr, ByVal nSize As Long, _
    ByRef Arguments As Any) As Long
Private Declare PtrSafe Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" ( _
    ByVal lpString1 As LongPtr, ByVal lpString2 As LongPtr) As Long
Private Declare PtrSafe Function LocalFree Lib "kernel32.dll" (ByVal hMem As LongPtr) As Long

#Else

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" ( _
    ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function FormatMessage Lib "kernel32.dll" Alias "FormatMessageW" ( _
    ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, _
    ByVal dwLanguageId As Long, ByRef lpBuffer As Long, ByVal nSize As Long, _
    ByRef Arguments As Any) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyW" ( _
    ByVal lpString1 As Long, ByVal lpString2 As Long) As Long
Private Declare Function LocalFree Lib "kernel32.dll" (ByVal hMem As Long) As Long

#End If


'BLOBHEADERユーザ定義型
Private Type BLOBHEADER
    bType As Byte
    bVersion As Byte
    reserved As Integer
    aiKeyAlg As Long
End Type

'インポート用の鍵データのユーザ定義型
'
'PUBLICKEYSTRUC BLOBヘッダに続いて、鍵サイズ、鍵データが必要だが、
'鍵データについては鍵サイズによって配列サイズが変わるため、
'ロジック中で動的にメモリを確保するようにし、ここでは未定義とする
Private Type keyBlob
    hdr As BLOBHEADER
    keySize As Long
'    keyData() As Byte
End Type

'鍵長定数定義
Public Enum AESKeyBits
    AES_KEY128 = 128
    AES_KEY192 = 192
    AES_KEY256 = 256
End Enum

'エラーコード定義
Private Const ERR_CRYPT_API = vbObjectError + 513   'CryptAPIエラー
Private Const ERR_KEY_LENGTH = vbObjectError + 514  '鍵長エラー
Private Const ERR_IV_LENGTH = vbObjectError + 515   'IV長エラー

'AES/CBC/PKCS5Padding復号処理
'
'引数:
'   [in]         key: 鍵バイト列
'   [in]          iv: IVバイト列
'   [in,out]    data: [in]暗号化バイト列/[out]復号バイト列
'   [in]     keyBits: 鍵ビット長(デフォルト128bit)
'
'戻り値:
'   なし
Public Sub decrypt(ByRef key() As Byte, ByRef iv() As Byte, ByRef data() As Byte, Optional ByVal keyBits As AESKeyBits = AES_KEY128)
#If VBA7 And Win64 Then
    Dim hProv As LongPtr   'CSPハンドラ
    Dim hKey As LongPtr    '暗号鍵ハンドラ
#Else
    Dim hProv As Long   'CSPハンドラ
    Dim hKey As Long    '暗号鍵ハンドラ
#End If
    Dim algid As Long   '暗号化アルゴリズム

    On Error GoTo ErrorHandler

    'AESの鍵長から暗号化アルゴリズムIDを設定
    Select Case keyBits
        Case AES_KEY128
            algid = CALG_AES_128
        Case AES_KEY192
            algid = CALG_AES_192
        Case AES_KEY256
            algid = CALG_AES_256
    End Select

    Dim keyLength As Long   '鍵バイト長
    keyLength = keyBits / 8 'ビット->バイト変換

    '鍵長のチェック
    If UBound(key) + 1 <> keyLength Then
        Err.Raise ERR_KEY_LENGTH, "decrypt()", "鍵長が不正です: " & UBound(key) + 1 & "byte"
    End If

    'IV長のチェック
    If UBound(iv) + 1 <> 16 Then
        Err.Raise ERR_IV_LENGTH, "decrypt()", "IV長が不正です: " & UBound(iv) + 1 & "byte"
    End If

    'CSP(Cryptographic Service Provider)のハンドルを取得
    If Not CBool(CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptAcquireContext()", Err.LastDllError
    End If

    Dim blob As keyBlob '鍵データ(ユーザ定義型)
    Dim keyData() As Byte   '鍵データ(バイト列)

    '鍵データの作成
    'keyBlobユーザ定義型に鍵データを結合したバイト列を無理やり作成する
    blob.hdr.bType = PLAINTEXTKEYBLOB
    blob.hdr.bVersion = CUR_BLOB_VERSION
    blob.hdr.reserved = 0
    blob.hdr.aiKeyAlg = algid
    blob.keySize = keyLength
    ReDim keyData(LenB(blob) + blob.keySize - 1)
    Call CopyMemory(keyData(0), blob, LenB(blob))
    Call CopyMemory(keyData(LenB(blob)), key(0), keyLength)

    '鍵のインポート
    If Not CBool(CryptImportKey(hProv, keyData(0), UBound(keyData) + 1, 0, 0, hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptImportKey()", Err.LastDllError
    End If

    'パディング方式の設定(PKCS#5)
    If Not CBool(CryptSetKeyParam(hKey, KP_PADDING, PKCS5_PADDING, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_PADDING", Err.LastDllError
    End If

    'IV(Initialization Vector)の設定
    If Not CBool(CryptSetKeyParam(hKey, KP_IV, iv(0), 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_IV", Err.LastDllError
    End If

    '暗号モードの設定(暗号文ブロック連鎖モード)
    If Not CBool(CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_CBC, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_MODE", Err.LastDllError
    End If

    '暗号化バイト列長
    Dim dwDataLen As Long
    dwDataLen = UBound(data) + 1

    'CryptDecryptは、引数の暗号化バイト列に復号したバイト列を戻す仕様のため
    'メソッドの引数の暗号化バイト列をローカル変数にコピーして使用する
    Dim pbData() As Byte
    ReDim pbData(dwDataLen - 1)
    Call CopyMemory(pbData(0), data(0), UBound(data) + 1)

    '復号処理
    If Not CBool(CryptDecrypt(hKey, 0, True, 0, pbData(0), dwDataLen)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptDecrypt()", Err.LastDllError
    End If

    ReDim Preserve pbData(dwDataLen - 1)
    data = pbData

    '暗号鍵ハンドラの開放
    If Not CBool(CryptDestroyKey(hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptDestroyKey()", Err.LastDllError
    End If

    'CSPハンドラの開放
    If Not CBool(CryptReleaseContext(hProv, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptReleaseContext()", Err.LastDllError
    End If

    Exit Sub

ErrorHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    errMessage = ""

    If Err.Number <> 0 Then
        If Err.Number = ERR_CRYPT_API Then
            errNumber = Err.Description
            errSource = Err.Source
            errMessage = GetErrorText(Err.Description)
        Else
            errNumber = Err.Number
            errSource = Err.Source
            errMessage = Err.Description
        End If
    End If

    Err.Clear

    If Not hKey <> 0 Then
        '暗号鍵ハンドラの開放
        Call CryptDestroyKey(hKey)
    End If

    If Not hProv <> 0 Then
        'CSPハンドラの開放
        Call CryptReleaseContext(hProv, 0)
    End If

    On Error GoTo 0
    If errMessage <> "" Then
        Err.Raise Number:=errNumber, Source:=errSource, Description:=errMessage
    End If
End Sub

'AES/CBC/PKCS5Padding暗号化処理
'
'引数:
'   [in]         key: 鍵バイト列
'   [in]          iv: IVバイト列
'   [in,out]    data: [in]平文バイト列/[out]暗号化バイト列
'   [in]     keyBits: 鍵ビット長(デフォルト128bit)
'
'戻り値:
'   なし
Public Sub encrypt(ByRef key() As Byte, ByRef iv() As Byte, ByRef data() As Byte, Optional ByVal keyBits As AESKeyBits = AES_KEY128)
#If VBA7 And Win64 Then
    Dim hProv As LongPtr   'CSPハンドラ
    Dim hKey As LongPtr    '暗号鍵ハンドラ
#Else
    Dim hProv As Long   'CSPハンドラ
    Dim hKey As Long    '暗号鍵ハンドラ
#End If
    Dim algid As Long   '暗号化アルゴリズム

    On Error GoTo ErrorHandler

    'AESの鍵長から暗号化アルゴリズムIDを設定
    Select Case keyBits
        Case AES_KEY128
            algid = CALG_AES_128
        Case AES_KEY192
            algid = CALG_AES_192
        Case AES_KEY256
            algid = CALG_AES_256
    End Select

    Dim keyLength As Long   '鍵バイト長
    keyLength = keyBits / 8 'ビット->バイト変換

    '鍵長のチェック
    If UBound(key) + 1 <> keyLength Then
        Err.Raise ERR_KEY_LENGTH, "decrypt()", "鍵長が不正です: " & UBound(key) + 1 & "byte"
    End If

    'IV長のチェック
    If UBound(iv) + 1 <> 16 Then
        Err.Raise ERR_IV_LENGTH, "decrypt()", "IV長が不正です: " & UBound(iv) + 1 & "byte"
    End If

    'CSP(Cryptographic Service Provider)のハンドルを取得
    If Not CBool(CryptAcquireContext(hProv, vbNullString, vbNullString, PROV_RSA_AES, CRYPT_VERIFYCONTEXT)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptAcquireContext()", Err.LastDllError
    End If

    Dim blob As keyBlob '鍵データ(ユーザ定義型)
    Dim keyData() As Byte   '鍵データ(バイト列)

    '鍵データの作成
    'keyBlobユーザ定義型に鍵データを結合したバイト列を無理やり作成する
    blob.hdr.bType = PLAINTEXTKEYBLOB
    blob.hdr.bVersion = CUR_BLOB_VERSION
    blob.hdr.reserved = 0
    blob.hdr.aiKeyAlg = algid
    blob.keySize = keyLength
    ReDim keyData(LenB(blob) + blob.keySize - 1)
    Call CopyMemory(keyData(0), blob, LenB(blob))
    Call CopyMemory(keyData(LenB(blob)), key(0), keyLength)

    '鍵のインポート
    If Not CBool(CryptImportKey(hProv, keyData(0), UBound(keyData) + 1, 0, 0, hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptImportKey()", Err.LastDllError
    End If

    'パディング方式の設定(PKCS#5)
    If Not CBool(CryptSetKeyParam(hKey, KP_PADDING, PKCS5_PADDING, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_PADDING", Err.LastDllError
    End If

    'IV(Initialization Vector)の設定
    If Not CBool(CryptSetKeyParam(hKey, KP_IV, iv(0), 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_IV", Err.LastDllError
    End If

    '暗号モードの設定(暗号文ブロック連鎖モード)
    If Not CBool(CryptSetKeyParam(hKey, KP_MODE, CRYPT_MODE_CBC, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptSetKeyParam():KP_MODE", Err.LastDllError
    End If

    '平文バイト列長
    Dim dwPlainDataLen As Long
    dwPlainDataLen = UBound(data) + 1

    '暗号化バイト列長
    Dim dwCryptDataLen As Long
    dwCryptDataLen = dwPlainDataLen

    'CryptEncryptは、引数の平文バイト列に暗号化したバイト列を戻す仕様のため
    'メソッドの引数の平文バイト列をローカル変数にコピーして使用する
    Dim pbData() As Byte
    ReDim pbData(dwPlainDataLen - 1)
    Call CopyMemory(pbData(0), data(0), dwPlainDataLen)

    '暗号化処理
    '暗号化後のバイト列長をあらかじめ問い合わせてバッファを拡張しておく
    If Not CBool(CryptEncrypt(hKey, 0, True, 0, ByVal 0&, dwCryptDataLen, dwPlainDataLen)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptEncrypt()", Err.LastDllError
    End If
    If dwCryptDataLen > dwPlainDataLen Then
        ReDim Preserve pbData(dwCryptDataLen - 1)
    End If
    If Not CBool(CryptEncrypt(hKey, 0, True, 0, pbData(0), dwPlainDataLen, dwCryptDataLen)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptEncrypt()", Err.LastDllError
    End If

    data = LeftB(pbData, dwCryptDataLen)

    '暗号鍵ハンドラの開放
    If Not CBool(CryptDestroyKey(hKey)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptDestroyKey()", Err.LastDllError
    End If

    'CSPハンドラの開放
    If Not CBool(CryptReleaseContext(hProv, 0)) Then
        Err.Raise ERR_CRYPT_API, "decrypt()->CryptReleaseContext()", Err.LastDllError
    End If

    Exit Sub

ErrorHandler:
    Dim errNumber As Long
    Dim errSource As String
    Dim errMessage As String

    errMessage = ""

    If Err.Number <> 0 Then
        If Err.Number = ERR_CRYPT_API Then
            errNumber = Err.Description
            errSource = Err.Source
            errMessage = GetErrorText(Err.Description)
        Else
            errNumber = Err.Number
            errSource = Err.Source
            errMessage = Err.Description
        End If
    End If

    Err.Clear

    If Not hKey <> 0 Then
        '暗号鍵ハンドラの開放
        Call CryptDestroyKey(hKey)
    End If

    If Not hProv <> 0 Then
        'CSPハンドラの開放
        Call CryptReleaseContext(hProv, 0)
    End If

    On Error GoTo 0
    If errMessage <> "" Then
        Err.Raise Number:=errNumber, Source:=errSource, Description:=errMessage
    End If
End Sub

'MAKELANGIDマクロのメソッド化による実装
Private Function MAKELANGID(ByVal p As Long, ByVal s As Long) As Long
    MAKELANGID = (CLng(CInt(s)) * 1024) Or CLng(CInt(p))
End Function

'エラーコードからエラーメッセージを取得する
Private Function GetErrorText(ByVal ErrorCode As Long) As String
#If VBA7 And Win64 Then
    Dim lpBuffer As LongPtr
#Else
    Dim lpBuffer As Long
#End If
    Dim messageLength As Long

    messageLength = FormatMessage( _
        FORMAT_MESSAGE_ALLOCATE_BUFFER Or FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, _
        0, ErrorCode, MAKELANGID(LANG_NEUTRAL, SUBLANG_DEFAULT), _
        lpBuffer, 0, 0)

    If messageLength = 0 Then
        GetErrorText = ""
    Else
        GetErrorText = Space$(messageLength)
        Call lstrcpy(ByVal StrPtr(GetErrorText), ByVal lpBuffer)
        Call LocalFree(lpBuffer)
    End If
End Function

