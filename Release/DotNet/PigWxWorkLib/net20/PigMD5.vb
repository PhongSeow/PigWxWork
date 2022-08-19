'**********************************
'* Name: PigMD5
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Standard MD5 treatment and personalized MD5 treatment.
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.4
'* Create Time: 1/12/2018
'1.0.2  增加小猪MD5，即取源串来的16位MD5加上源串的前128个字节（不足取源串长度）的16位MD5
'1.0.3  PigMD5增加数组
'1.0.4  Add PigBaseMini
'**********************************
Public Class PigMD5
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.4"

    ''' <summary>源数据</summary>
    Private mabSrcData As Byte()
    ''' <summary>MD数组</summary>
    Private mabHashValue As Byte()
    ''' <summary>MD字符串</summary>
    Private mstrMD5 As String
    ''' <summary>小猪MD字符串</summary>
    Private mabPigMD5 As Byte()
    Private mstrPigMD5 As String

    Public Enum enmTextType '文本类型
        Unknow = 0
        Unicode = 1
        UTF8 = 2
        Ascii = 3
    End Enum

    Sub New(SrcData As Byte())
        MyBase.New(CLS_VERSION)
        Try
            mabSrcData = SrcData
            mabHashValue = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(mabSrcData)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub

    Sub New(SrcStr As String, TextType As enmTextType)
        MyBase.New(CLS_VERSION)
        Try
            Select Case TextType
                Case enmTextType.Ascii
                    mabSrcData = System.Text.Encoding.ASCII.GetBytes(SrcStr)
                Case enmTextType.Unicode
                    mabSrcData = System.Text.Encoding.Unicode.GetBytes(SrcStr)
                Case enmTextType.UTF8
                    mabSrcData = System.Text.Encoding.UTF8.GetBytes(SrcStr)
                Case Else
                    mabSrcData = System.Text.Encoding.UTF8.GetBytes(SrcStr)
            End Select
            mabHashValue = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(mabSrcData)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("New", ex)
        End Try
    End Sub


    Public ReadOnly Property MD5Bytes As Byte()
        Get
            MD5Bytes = mabHashValue
        End Get
    End Property

    Public ReadOnly Property PigMD5Bytes As Byte()
        Get
            If Len(mstrPigMD5) <= 0 Then Me.mMkPigMD5()
            PigMD5Bytes = mabPigMD5
        End Get
    End Property


    Private Function mMD5() As String
        Try
            If Len(mstrMD5) <= 0 Then
                Dim i As Integer
                mstrMD5 = ""
                For i = 0 To 15 '选择32位字符的加密结果
                    mstrMD5 &= Right("00" & Hex(mabHashValue(i)).ToLower, 2)
                Next
            End If
            Return mstrMD5
        Catch ex As Exception
            Return ""
        End Try
    End Function


    Public Overloads ReadOnly Property MD5 As String
        Get
            MD5 = mMD5()
        End Get
    End Property

    Public Overloads ReadOnly Property MD5(Is16Bit As Boolean) As String
        Get
            Dim strMD5 As String = mMD5()
            If Is16Bit = True Then
                Return Mid(strMD5, 9, 16)
            Else
                Return strMD5
            End If
        End Get
    End Property

    Private Sub mMkPigMD5()
        Try
            Dim i As Short, intLen As Long = mabSrcData.Length - 1
            If intLen > 127 Then intLen = 127
            Dim abPig(intLen) As Byte
            For i = 0 To intLen
                abPig(i) = mabSrcData(intLen - i)
            Next
            Dim oPig As New PigMD5(abPig)
            ReDim mabPigMD5(15)
            For i = 0 To 7
                mabPigMD5(i) = mabHashValue(i + 4)
                mabPigMD5(i + 8) = oPig.MD5Bytes(i + 4)
            Next
            mstrPigMD5 = ""
            For i = 0 To 15 '选择32位字符的加密结果
                mstrPigMD5 &= Right("00" & Hex(mabPigMD5(i)).ToLower, 2)
            Next
            oPig = Nothing
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mMkPigMD5", ex)
        End Try
    End Sub

    Public ReadOnly Property PigMD5 As String
        Get
            Try
                If Len(mstrPigMD5) <= 0 Then
                    Me.mMkPigMD5()
                End If
                Return mstrPigMD5
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property

End Class
