Class VbsRSACrypto
    ' 定数
    ' ADODB.Stream
    Private adTypeBinary
    Private adTypeText
    Private adSaveCreateNotExist
    Private adSaveCreateOverWrite
    Private adReadAll
    Private adReadLine
    ' Original
    Private csUnicode
    Private csUTF8
    Private csJIS
    Private csShiftJIS
    Private csEUCJP
    ' Objects
    Private objRsaCSP
    Private objStream
    Private objUtf8

    Private pXMLPrivateKey
    Private pXMLPublicKey

    'Constructor
    Private Sub Class_Initialize
        ' ADODB.Stream
        adTypeBinary = 1
        adTypeText = 2
        adSaveCreateNotExist = 1
        adSaveCreateOverWrite = 2
        adReadAll = -1
        adReadLine = -2

        ' Original
        csUnicode = "unicode" ' Script Default
        csUTF8 = "utf-8"
        csJIS = "iso-2022-jp"
        csShiftJIS = "Shift_JIS"
        csEUCJP = "euc-jp"

        '暗号化のためのライブラリ
        Set objRsaCSP = CreateObject("System.Security.Cryptography.VbsRSACryptoServiceProvider")

        'バイナリを扱うためのライブラリ
        Set objStream = CreateObject("ADODB.Stream")

        'Utf8データを扱うためのライブラリ
        Set objUtf8 = CreateObject("System.Text.UTF8Encoding")
    End Sub

    'Destructor
    Private Sub Class_Terminate
        Set objRsaCSP = Nothing
        Set objStream = Nothing
        Set objUtf8 = Nothing
    End Sub

    'Property
    Property Let XmlPrivateKey(value)
        pXMLPrivateKey = value
    End Property
    Property Get XmlPrivateKey
        XmlPrivateKey = pXMLPrivateKey
    End Property
    Property Let XmlPublicKey(value)
        pXMLPublicKey = value
    End Property
    Property Get XmlPublicKey
        XmlPublicKey = pXMLPublicKey
    End Property

    Property Get KeyExchangeAlgorithm
        KeyExchangeAlgorithm = objRsaCSP.KeyExchangeAlgorithm
    End Property
    Property Get KeySize
        KeySize = objRsaCSP.KeySize
    End Property
    Property Get PersistKeyInCsp
        PersistKeyInCsp = objRsaCSP.PersistKeyInCsp
    End Property
    Property Get SignatureAlgorithm
        SignatureAlgorithm = objRsaCSP.SignatureAlgorithm
    End Property

    Sub EchoProperty()

    End Sub
    'Method
    '[引数]
    'DataToEncrypt:文字列
    'DoOAEPPadding:真偽値
    '[戻り値]
    'Base64エンコードされた文字列
    Function Encrypt(DataToEncrypt, DoOAEPPadding)
        Dim bin

        objStream.Open
        objStream.Type = adTypeText
        objStream.Charset = csUTF8
        objStream.WriteText DataToEncrypt
        objStream.Position = 0
        objStream.Type = adTypeBinary

        objRsaCSP.FromXmlString(XMLPublicKey)
        bin = objRsaCSP.Encrypt(objStream.Read(adReadAll), DoOAEPPadding)
        Encrypt = Base64Encode(bin)
        objStream.Close
    End Function
    '[引数]
    'EncryptValue:Base64エンコードされた文字列
    'DoOAEPPadding:真偽値
    '[戻り値]
    '文字列
    Function Decrypt(EncryptValue, DoOAEPPadding)
        Dim bin

        binEncryptValue = Base64Decode(EncryptValue)
        objStream.Open
        objStream.Type = adTypeBinary
        objStream.Write binEncryptValue
        objStream.Position = 0

        objRsaCSP.FromXmlString(XmlPrivateKey)
        bin = objRsaCSP.Decrypt(objStream.Read(adReadAll), DoOAEPPadding)
        Decrypt = GetUtf8String(bin)
        objStream.Close
    End Function
    '[引数]
    'binData:バイト配列
    '[戻り値]
    'Base64エンコードされた文字列
    '[参考URL]
    'https://phpvbs.verygoodtown.com/vbscript-base64_encode-function-jp/
    Private Function Base64Encode(binData)
        Dim dom, element
        Set dom = CreateObject("Microsoft.XMLDOM")
        Set element = dom.CreateElement("tmp")
        element.DataType = "bin.base64"
        element.NodeTypedValue = binData
        Base64Encode = element.Text
    End Function
    '[引数]
    'Base64String:Base64エンコードされた文字列
    '[戻り値]
    'バイト配列
    '[参考URL]
    'https://phpvbs.verygoodtown.com/vbscript-base64_decode-function-jp/
    Private Function Base64Decode(Base64String)
        Dim dom, element

        Set dom = CreateObject("Microsoft.XMLDOM")
        Set element = dom.createElement("tmp")
        element.DataType = "bin.base64"
        element.Text = Base64String
        Base64Decode = element.NodeTypedValue
    End Function
    Private Function GetUtf8String(binData)
        GetUtf8String = objUtf8.GetString((binData))
    End Function
End Class