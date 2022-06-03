
# 概要

VbscriptでRSA暗号化、復号を行うClassです。

# プロパティ

## XmlPrivateKey

`string`

RSAのPrivateKey（XML形式の文字列）を指定します。

## XmlPublicKey

`string`

RSAのPublicKey（XML形式の文字列）を指定します。

## KeyExchangeAlgorithm

`string`

RSA のこの実装で使用可能なキー交換アルゴリズムの名前を取得します。


## KeySize

`int`

カーソル キーのサイズを取得します。

## PersistKeyInCsp

`boolean`

暗号化サービス プロバイダー (CSP) にキーを保存するかどうかを示す値を取得します。(真偽値)

## SignatureAlgorithm

# メソッド

## Encrypt


### 引数

DataToEncrypt `string` 

暗号化したい平文 

DoOAEPPadding `boolean` 

OAEP パディング (Windows XP 以降を実行しているコンピューターでのみ使用可能) を使用して、直接 RSA を暗号化する場合は true。それ以外の場合で、PKCS#1 v1.5 パディングを使用するには false。

### 戻り値
Base64エンコードされた文字列 `string`

## Decrypt

### 引数

EncryptValue `string` 

Base64エンコードされた暗号化された文字列 

DoOAEPPadding `boolean` 

OAEP パディング (Windows XP 以降を実行しているコンピューターでのみ使用可能) を使用して、直接 RSA を暗号化する場合は true。それ以外の場合で、PKCS#1 v1.5 パディングを使用するには false。

### 戻り値
復号文字列 `string`

# その他

参考にしたソースコード

```
'http://www.rainylain.jp/vc/windows_script.htm
Dim strSrcTextValue ' 暗号化したい文字列
Dim binEncryptValue ' 暗号化された値（バイナリ値：SafeArry）
Dim binDecryptValue ' 復号化された値（バイナリ値：SafeArry）
Dim strDecryptValue ' 復号化された文字列
    
strSrcTextValue = "abc"

' 定義
' ADODB.Stream
Const adTypeBinary = 1
Const adTypeText = 2
Const adSaveCreateNotExist = 1
Const adSaveCreateOverWrite = 2
Const adReadAll = -1
Const adReadLine = -2

' Original
Const csUnicode = "unicode" ' Script Default
Const csUTF8 = "utf-8"
Const csJIS = "iso-2022-jp"
Const csShiftJIS = "Shift_JIS"
Const csEUCJP = "euc-jp"
   

' バイナリを扱うために ADODB.Stream を使用する
Set objStream = CreateObject("ADODB.Stream")
  

' レジストリの「HKEY_CLASSES_ROOT\System.Security.Cryptography.～Provider」群を
' 見れば、使用できる暗号の種類が分かる。ここではRSA暗号化を指定している
Set objCrypt = CreateObject("System.Security.Cryptography.VbsRSACryptoServiceProvider")


' Encrypt
' ストリームをオープンして暗号化したい文字列を書き込む
objStream.Open
objStream.Type = adTypeText
objStream.Charset = csUnicode
objStream.WriteText strSrcTextValue

' ストリームをバイナリ指定にしてvbSafeArray（配列）をEncrypt に渡して暗号化する
' （Encrypt/Decryptには 8209型 = vbArray(8192) + vbByte(17)が要求される）
' これで binEncryptValue に暗号化された値がバイナリ値で格納された状態になるので
' 文字列化するなりそのままバイナリ値として、ファイルに書き込んだりレジストリに格納
' すれば良いだろう
objStream.Position = 0
objStream.Type = adTypeBinary

binEncryptValue = objCrypt.Encrypt(objStream.Read(adReadAll), False)

objStream.Close

WScript.Echo Mid(binEncryptValue, 2)
' XMLでカギ情報を出力(True：秘密カギと公開カギを出力、False：公開カギのみ出力)
'   このプロセス内でのみ暗号化/復号化したい場合は必要ないが、後で使用したいような場合は
'   カギ情報を残しておかないとならない。
'   ※間違っても平文で他人も見えるような場所に保存してはならない！
strXML = objCrypt.ToXmlString(True)


WScript.Echo strXML

' 暗号化カギを格納しているXMLを入力する（XMLは保存先から別途何らかの方法で読み込むこと）
' これにより、プロセスを終わらせても暗号化/復号化の続きができる。
objCrypt.FromXmlString(strXML)

' Decrypt
' ストリームに暗号化されたバイナリ値を格納する
' 暗号化されたバイナリ値をファイルから読み込みたい場合はWriteではなくLoadFromFileする
objStream.Open
objStream.Type = adTypeBinary
objStream.Write binEncryptValue
    
' ストリームをバイナリ指定にしてvbSafeArray（配列）をDecrypt に渡して復号化する
objStream.Position = 0

binDecryptValue = objCrypt.Decrypt(objStream.Read(adReadAll), False)

objStream.Close

' Unicode Textのバイトオーダーを示す BOM コードを削除する
' これによりスクリプト上で通常の文字列として比較などが行えるようになる
strDecryptValue = Mid(binDecryptValue, 2)
WScript.Echo "Decrypt: " & strDecryptValue

' 終了処理
Set objCrypt = Nothing
Set objStream = Nothing
```