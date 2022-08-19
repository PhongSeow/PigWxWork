'**********************************
'* Name: PigRsa
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Handle text conversion, MD5, Base64, etc.|处理RSA的加解密
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 17/10/2019
'1.0.2  2019-10-18 Decrypt和Encrypt重载函数
'1.0.3  2019-11-17 增加 密钥 bytes 数据处理
'1.1  16/10/2021 Modify LoadPubKey,mDecrypt,mEncrypt
'**********************************

Imports System.Security.Cryptography
Public Class PigRsa
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.6"

    ''' <summary>密钥组成部分结构长度</summary>
    Private Structure struEncKeyPartLen
        Dim Modulus As Short
        Dim Exponent As Short
        ''' <remarks>以下的可能为0，只有选择了私钥才有</remarks>
        Dim P As Short
        Dim Q As Short
        Dim DP As Short
        Dim DQ As Short
        Dim InverseQ As Short
        Dim D As Short
    End Structure

    Private mrsaMain As RSACryptoServiceProvider = New RSACryptoServiceProvider
    Private mbolIsLoadEncKey As Boolean

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    ''' <remarks>是否已导入密钥</remarks>
    Public ReadOnly Property IsLoadEncKey As Boolean
        Get
            IsLoadEncKey = mbolIsLoadEncKey
        End Get
    End Property


    ''' <remarks>解密数据</remarks>
    Public Overloads Function Decrypt(EncBase64Str As String, ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oGEText As PigText
            oGEText = New PigText(EncBase64Str, TextType, PigText.enmNewFmt.FromBase64)
            Dim abUnEncBytes(0) As Byte
            strStepName = "mEncrypt"
            strRet = Me.mDecrypt(oGEText.TextBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oGEText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oGEText.Text
            oGEText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", ex, True)
        End Try
    End Function

    ''' <remarks>解密数据</remarks>
    Public Overloads Function Decrypt(EncBytes As Byte(), ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oGEText As PigText
            Dim abUnEncBytes(0) As Byte
            strStepName = "mEncrypt"
            strRet = Me.mDecrypt(EncBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oGEText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oGEText.Text
            oGEText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", ex, True)
        End Try
    End Function

    ''' <remarks>解密数据</remarks>
    Public Overloads Function Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Decrypt = Me.mDecrypt(EncBytes, UnEncBytes)
    End Function


    ''' <remarks>解密数据（内部）</remarks>
    Private Function mDecrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Dim strStepName As String = ""
        Try
            If mbolIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
            UnEncBytes = mrsaMain.Decrypt(EncBytes, True)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mDecrypt", ex, False)
        End Try
    End Function

    ''' <remarks>加密数据</remarks>
    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBase64Str As String) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim abEncBytes As Byte()
            ReDim abEncBytes(0)
            strStepName = "mEncrypt"
            strRet = Me.mEncrypt(SrcBytes, abEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            EncBase64Str = Convert.ToBase64String(abEncBytes)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Encrypt", ex, False)
        End Try
    End Function


    ''' <remarks>加密数据</remarks>
    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Return Me.mEncrypt(SrcBytes, EncBytes)
    End Function


    ''' <remarks>加密数据（内部）</remarks>
    Private Function mEncrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Dim strStepName As String = ""
        Try
            If mbolIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
            strStepName = "mrsaMain.Encrypt"
            EncBytes = mrsaMain.Encrypt(SrcBytes, True)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mEncrypt", ex, False)
        End Try
    End Function

    ''' <summary>生成密钥</summary>
    ''' <param name="IsIncPriKey">是否生成私钥</param>
    ''' <param name="OutXmlKey">输出为XML</param>
    Public Function MkPubKey(IsIncPriKey As Boolean, ByRef OutXmlKey As String) As String
        Try
            OutXmlKey = mrsaMain.ToXmlString(IsIncPriKey)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkPubKey", ex, False)
        End Try
    End Function

    ''' <summary>生成密钥</summary>
    ''' <param name="IsIncPriKey">是否生成私钥</param>
    ''' <param name="OutBytes">输出为字节数组</param>
    Public Function MkPubKey(IsIncPriKey As Boolean, ByRef OutBytes As Byte()) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
            Dim strXmlKey As String = mrsaMain.ToXmlString(IsIncPriKey)
            Dim ekplAny As struEncKeyPartLen
            Dim oGEXml As New PigXml(False), oGEText As PigText, strItem As String
            strXmlKey = oGEXml.XmlGetStr(strXmlKey, "RSAKeyValue")
            Dim oGEBytes As New PigBytes
            oGEBytes.RestPos()
            With ekplAny
                strItem = oGEXml.XmlGetStr(strXmlKey, "Modulus")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .Modulus = oGEText.TextBytes.Length
                strStepName = ".SetValue(Modulus)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "Exponent")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8)
                .Exponent = oGEText.TextBytes.Length
                strStepName = ".SetValue(Exponent)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "P")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .P = oGEText.TextBytes.Length
                strStepName = ".SetValue(P)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "Q")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .Q = oGEText.TextBytes.Length
                strStepName = ".SetValue(Q)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "DP")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .DP = oGEText.TextBytes.Length
                strStepName = ".SetValue(DP)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "DQ")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .DQ = oGEText.TextBytes.Length
                strStepName = ".SetValue(DQ)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "InverseQ")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .InverseQ = oGEText.TextBytes.Length
                strStepName = ".SetValue(InverseQ)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
                strItem = oGEXml.XmlGetStr(strXmlKey, "D")
                oGEText = New PigText(strItem, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
                .D = oGEText.TextBytes.Length
                strStepName = ".SetValue(D)"
                strRet = oGEBytes.SetValue(oGEText.TextBytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
                oGEText = Nothing
            End With
            Dim oStru2Bytes As New Stru2Bytes
            strStepName = ".Stru2Bytes(ekplAny, OutBytes)"
            strRet = oStru2Bytes.Stru2Bytes(ekplAny, OutBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Dim gbOutBytes As New PigBytes
            strStepName = ".SetValue(CInt(OutBytes.Length))"
            strRet = gbOutBytes.SetValue(CInt(OutBytes.Length))
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = ".SetValue(OutBytes)"
            strRet = gbOutBytes.SetValue(OutBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = ".SetValue(oGEBytes.Main)"
            strRet = gbOutBytes.SetValue(oGEBytes.Main)
            If strRet <> "OK" Then Throw New Exception(strRet)
            OutBytes = gbOutBytes.Main
            oStru2Bytes = Nothing
            oGEXml = Nothing
            oGEBytes = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkPubKey", ex, False)
        End Try
    End Function

    ''' <summary>导入密钥</summary>
    ''' <param name="KeyBytes">密钥数组</param>
    Public Overloads Function LoadPubKey(KeyBytes As Byte()) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            strStepName = "New PigBytes"
            Dim oGEBytes As New PigBytes(KeyBytes)
            If oGEBytes.LastErr <> "" Then Throw New Exception(oGEBytes.LastErr)
            Dim oStru2Bytes As New Stru2Bytes
            Dim ekplAny As struEncKeyPartLen
            Dim intStruLen As Integer = oGEBytes.GetInt32Value
            strStepName = "Bytes2Stru"
            strRet = oStru2Bytes.Bytes2Stru(oGEBytes.GetBytesValue(intStruLen), ekplAny.GetType, ekplAny)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Dim oGEXml As New PigXml(False), oGEText As PigText
            strStepName = "Mk Xml"
            oGEXml.AddEleLeftSign("RSAKeyValue")
            strStepName = "New PigText(ekplAny.Modulus)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.Modulus))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("Modulus", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.Exponent)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.Exponent), PigText.enmTextType.UTF8)
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("Exponent", oGEText.Text)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.P)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.P))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("P", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.Q)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.Q))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("Q", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.DP)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.DP))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("DP", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.DQ)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.DQ))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("DQ", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.InverseQ)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.InverseQ))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("InverseQ", oGEText.Base64)
            oGEText = Nothing
            strStepName = "New PigText(ekplAny.D)"
            oGEText = New PigText(oGEBytes.GetBytesValue(ekplAny.D))
            If oGEText.LastErr <> "" Then Throw New Exception(oGEText.LastErr)
            oGEXml.AddEle("D", oGEText.Base64)
            oGEText = Nothing
            oGEXml.AddEleRightSign("RSAKeyValue")
            strStepName = "FromXmlString"
            mrsaMain.FromXmlString(oGEXml.MainXmlStr)
            mbolIsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadPubKey", ex)
        End Try
    End Function


    ''' <summary>导入密钥</summary>
    ''' <param name="XmlKey">XML密钥</param>
    Public Overloads Function LoadPubKey(XmlKey As String) As String
        Dim strStepName As String = ""
        Try
            strStepName = "FromXmlString"
            mrsaMain.FromXmlString(XmlKey)
            mbolIsLoadEncKey = True
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadPubKey", strStepName, ex)
        End Try
    End Function

    Protected Overrides Sub Finalize()
        mrsaMain = Nothing
        MyBase.Finalize()
    End Sub
End Class
