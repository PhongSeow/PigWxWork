'**********************************
'* Name: PigAes
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Handle operations related to byte array division 【处理除字节数组相关的操作】
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 2019-10-27
'1.0.2  2019-10-29
'1.0.3  2019-10-31  稳定版本，去掉 EncriptStr 和 DecriptStr
'1.0.5  2019-12-8   修改LoadEncKey
'1.0.6  24/8/2021   Modify mDecrypt,mEncrypt,mMkEncKey
'1.1  16/10/2021   Modify mDecrypt,mEncrypt,LoadEncKey,LoadEncKey,Decrypt
'************************************
Imports System.Security.Cryptography
Imports System.Text
Public Class PigAes
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.8"
    Private mabEncKey As Byte()
    Private mbolIsLoadEncKey As Boolean

#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
    Private maesMain As System.Security.Cryptography.Aes = System.Security.Cryptography.Aes.Create("AES")
#End If


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
            Dim oPigText As PigText
            oPigText = New PigText(EncBase64Str, TextType, PigText.enmNewFmt.FromBase64)
            Dim abUnEncBytes(0) As Byte
            strStepName = "mDecrypt"
            strRet = Me.mDecrypt(oPigText.TextBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oPigText.Text
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", strStepName, ex)
        End Try
    End Function

    ''' <remarks>解密数据</remarks>
    Public Overloads Function Decrypt(EncBytes As Byte(), ByRef UnEncStr As String, TextType As PigText.enmTextType) As String
        Dim strStepName As String = ""
        Dim strRet As String
        Try
            Dim oPigText As PigText
            Dim abUnEncBytes As Byte()
            ReDim abUnEncBytes(0)
            strStepName = "mDecrypt"
            strRet = Me.mDecrypt(EncBytes, abUnEncBytes)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oPigText = New PigText(abUnEncBytes, TextType)
            UnEncStr = oPigText.Text
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Decrypt", strStepName, ex)
        End Try
    End Function

    ''' <remarks>解密数据</remarks>
    Public Overloads Function Decrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Return Me.mDecrypt(EncBytes, UnEncBytes)
    End Function


    ''' <remarks>解密数据（内部）</remarks>
    Private Function mDecrypt(EncBytes As Byte(), ByRef UnEncBytes As Byte()) As String
        Dim strStepName As String = ""
        Try
            If mbolIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            strStepName = "CreateDecryptor"
            Dim ictAny As ICryptoTransform = maesMain.CreateDecryptor()
            strStepName = "TransformFinalBlock"
            UnEncBytes = ictAny.TransformFinalBlock(EncBytes, 0, EncBytes.Length)
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mDecrypt", ex)
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
            Return Me.GetSubErrInf("Encrypt", ex)
        End Try
    End Function


    ''' <remarks>加密数据</remarks>
    Public Function Encrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Return Me.mEncrypt(SrcBytes, EncBytes)
    End Function


    ''' <remarks>加密数据（内部）</remarks>
    Private Function mEncrypt(SrcBytes As Byte(), ByRef EncBytes As Byte()) As String
        Try
            If mbolIsLoadEncKey = False Then
                Throw New Exception("Key not imported")
            End If
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            Dim ictAny As ICryptoTransform = maesMain.CreateEncryptor()
            EncBytes = ictAny.TransformFinalBlock(SrcBytes, 0, SrcBytes.Length)
            ictAny = Nothing
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mEncrypt", ex, False)
        End Try
    End Function

    ''' <summary>获取随机数</summary>
    ''' <param name="BeginNum">起始数</param>
    ''' <param name="EndNum">结束数</param>
    Private Function mGetRandNum(BeginNum As Integer, EndNum As Integer) As Integer
        Dim i As Long
        Try
            If BeginNum > EndNum Then
                i = BeginNum
                BeginNum = EndNum
                EndNum = i
            End If
            Randomize()   '初始化随机数生成器
            mGetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            mGetRandNum = 0
        End Try
    End Function

    ''' <remarks>生成密钥</remarks>
    Public Function MkEncKey(ByRef Base64EncKey As String) As String
        Try
            Return Me.mMkEncKey(256, Base64EncKey)
        Catch ex As Exception
            Return Me.GetSubErrInf("MkEncKey", ex, False)
        End Try
    End Function

    Private Function mMkEncKey(EncKeyLen As Integer, ByRef Base64EncKey As String) As String
        Try
            Select Case EncKeyLen
                Case 128, 192, 256
                Case Else
                    Throw New Exception("Key length cannot be greater than " & EncKeyLen.ToString)
            End Select
            Dim i As Integer
            ReDim mabEncKey(EncKeyLen - 1)
            For i = 0 To EncKeyLen - 1
                mabEncKey(i) = Me.mGetRandNum(0, 255)
            Next
            Dim oPigText As New PigText(mabEncKey)
            Base64EncKey = oPigText.Base64
            oPigText = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("MkEncKey", ex)
        End Try
    End Function

    ''' <remarks>导入密钥</remarks>
    Public Overloads Function LoadEncKey(Base64EncKey As String) As String
        Try
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            Dim oPigText As New PigText(Base64EncKey, PigText.enmTextType.UTF8, PigText.enmNewFmt.FromBase64)
            mabEncKey = (New MD5CryptoServiceProvider).ComputeHash(oPigText.TextBytes)
            With maesMain
                .BlockSize = mabEncKey.Length * 8
                .Key = mabEncKey
                .IV = mabEncKey
                .Mode = CipherMode.CBC
                .Padding = PaddingMode.PKCS7
            End With
            mbolIsLoadEncKey = True
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadEncKey", ex, False)
        End Try
    End Function

    ''' <remarks>导入密钥</remarks>
    Public Overloads Function LoadEncKey(EncKey As Byte()) As String
        Try
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            mabEncKey = (New MD5CryptoServiceProvider).ComputeHash(EncKey)
            With maesMain
                .BlockSize = mabEncKey.Length * 8
                .Key = mabEncKey
                .IV = mabEncKey
                .Mode = CipherMode.CBC
                .Padding = PaddingMode.PKCS7
            End With
            mbolIsLoadEncKey = True
#Else
            Throw New Exception("Need to run in .net 4.0 or higher framework")
#End If
            Return "OK"
        Catch ex As Exception
            mbolIsLoadEncKey = False
            Return Me.GetSubErrInf("LoadEncKey", ex, False)
        End Try
    End Function

    Protected Overrides Sub Finalize()
        MyBase.Finalize()
    End Sub

End Class
