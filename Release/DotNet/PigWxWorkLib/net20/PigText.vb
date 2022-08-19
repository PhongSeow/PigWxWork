'**********************************
'* Name: PigText
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Handle text conversion, MD5, Base64, etc.|处理文本转换，MD5，Base64等。
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.16
'* Create Time: 1/9/2019
'1.0.2  2019-9-2    增加 mNew(IsFromBase64)
'1.0.5  2019-10-16  增加 mstrText
'1.0.6  2019-10-17  增加 UrlEncode
'1.0.6  2019-10-18  增加 SHA1
'1.0.8  2019-10-19  增加 HexStr
'1.0.9  2019-11-17  New Bug
'1.0.10 2019-12-23  去fGEToolsCs 和 增加 GEBaseMini
'1.0.11 2019-12-26  增加 PigMD5 即加强MD5，因为两个字符串的MD5有可能相同，加强版是将源字节数组的前128字节的Base64的16位MD5加在后面
'1.0.12 2019-12-28  修订 PigMD5 BUG
'1.0.13 2020-1-4    加入GEMD5
'1.0.15 2020-2-6    修改enmTextType
'1.0.16 2020-2-12   修改 mNew
'**********************************
Public Class PigText
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.16"

    Public Enum enmTextType '文本类型
        UnknowOrBin = 0
        Unicode = 1
        UTF8 = 2
        Ascii = 3
    End Enum

    ''' <summary>通过初始化数据的格式</summary>
    Public Enum enmNewFmt
        ''' <summary>文本字符串</summary>
        FromText = 0
        ''' <summary>Base64字符串</summary>
        FromBase64 = 2
        ''' <summary>16进制格式</summary>
        FromHexStr = 3
    End Enum

    Private mabSrcText As Byte() '源文本的Byte数组，未压缩
    Private mabCompressText As Byte() '压缩过的源文本数组
    Private mstrText As String '文本的字符串
    Private mintTextType As enmTextType
    Private moMD5 As PigMD5
    Private mstrHexStr As String
    '--------------------------
    ''' <remarks>16进制格式字符串</remarks>
    Public ReadOnly Property HexStr() As String
        Get
            Try
                If mstrHexStr = "" Then
                    Dim intLen As Integer = mabSrcText.Length, i As Integer
                    Dim sbAny As New System.Text.StringBuilder("")
                    For i = 0 To intLen - 1
                        sbAny.Append(mabSrcText(i).ToString("x2"))
                    Next
                    mstrHexStr = sbAny.ToString
                    sbAny = Nothing
                End If
                Return mstrHexStr
            Catch ex As Exception
                Me.SetSubErrInf("HexStr", ex)
                Return ""
            End Try
        End Get
    End Property

    ''' <remarks>小猪MD5</remarks>
    Public ReadOnly Property PigMD5() As String

        Get
            Dim strStepName As String = ""
            Try
                If moMD5 Is Nothing Then
                    strStepName = "New PigMD5"
                    moMD5 = New PigMD5(mabSrcText)

                End If
                Return moMD5.PigMD5
            Catch ex As Exception
                Me.SetSubErrInf("PigMD5", strStepName, ex)
                Return ""
            End Try
        End Get
    End Property

    ''' <remarks>SHA1</remarks>
    Public ReadOnly Property SHA1() As String
        Get
            Try
                Dim shaAny As New System.Security.Cryptography.SHA1CryptoServiceProvider
                Dim abToHash() As Byte
                abToHash = shaAny.ComputeHash(mabSrcText)
                SHA1 = ""
                For Each bAny As Byte In abToHash
                    SHA1 += bAny.ToString("x2")
                Next
                shaAny.Clear()
                shaAny = Nothing
            Catch ex As Exception
                Return ""
            End Try
        End Get
    End Property

    ''' <remarks>获取Base64编码文本</remarks>
    Public ReadOnly Property Base64() As String
        Get
            Base64 = Convert.ToBase64String(mabSrcText)
        End Get
    End Property

    ''' <remarks>MD5数组</remarks>
    Public ReadOnly Property MD5Bytes As Byte()
        Get
            If moMD5 Is Nothing Then moMD5 = New PigMD5(mabSrcText)
            MD5Bytes = moMD5.MD5Bytes
        End Get
    End Property

    ''' <remarks>32位MD5文本</remarks>
    Public ReadOnly Property MD5() As String
        Get
            If moMD5 Is Nothing Then moMD5 = New PigMD5(mabSrcText)
            MD5 = moMD5.MD5
        End Get
    End Property


    ''' <remarks>32/16位MD5文本</remarks>
    Public ReadOnly Property MD5(Is16Bit As Boolean) As String
        Get
            If moMD5 Is Nothing Then moMD5 = New PigMD5(mabSrcText)
            MD5 = moMD5.MD5(Is16Bit)
        End Get
    End Property


    ''' <remarks>压缩过的文本数组</remarks>
    Public ReadOnly Property CompressTextBytes As Byte()
        Get
            Dim strStepName As String = ""
            Try
                If mabCompressText Is Nothing Then
                    Dim oCompressor As New PigCompressor
                    strStepName = "Compress(mabSrcText)"
                    mabCompressText = oCompressor.Compress(mabSrcText)
                    If oCompressor.LastErr <> "" Then Err.Raise(-1, , oCompressor.LastErr)
                    oCompressor = Nothing
                End If
                CompressTextBytes = mabCompressText
                Me.ClearErr()
            Catch ex As Exception
                Me.SetSubErrInf("CompressTextBytes", strStepName, ex)
                Return Nothing
            End Try
        End Get
    End Property

    ''' <remarks>文本数组</remarks>
    Public ReadOnly Property TextBytes As Byte()
        Get
            TextBytes = mabSrcText
        End Get
    End Property

    ''' <remarks>文本的字符串</remarks>
    Public ReadOnly Property Text As String
        Get
            If mstrText = "" Then
                Select Case mintTextType
                    Case enmTextType.Ascii
                        mstrText = System.Text.Encoding.ASCII.GetString(mabSrcText)
                    Case enmTextType.Unicode
                        mstrText = System.Text.Encoding.Unicode.GetString(mabSrcText)
                    Case enmTextType.UTF8
                        mstrText = System.Text.Encoding.UTF8.GetString(mabSrcText)
                End Select
            End If
            Text = mstrText
        End Get
    End Property

    ''' <remarks>文本类型</remarks>
    Public ReadOnly Property TextType As enmTextType
        Get
            TextType = mintTextType
        End Get
    End Property

    Sub New(BinData As Byte())
        MyBase.New(CLS_VERSION)
        mabSrcText = BinData
        mintTextType = enmTextType.UTF8
    End Sub

    Sub New(BinData As Byte(), TextType As enmTextType)
        MyBase.New(CLS_VERSION)
        mabSrcText = BinData
        mintTextType = TextType
    End Sub

    Sub New(StrData As String, TextType As enmTextType)
        MyBase.New(CLS_VERSION)
        Me.mNew(StrData, TextType, enmNewFmt.FromText)
    End Sub

    Sub New(StrData As String, TextType As enmTextType, NewFmt As enmNewFmt)
        MyBase.New(CLS_VERSION)
        Me.mNew(StrData, TextType, NewFmt)
    End Sub

    Private Sub mNew(StrData As String, TextType As enmTextType, NewFmt As enmNewFmt)
        mintTextType = TextType
        Select Case NewFmt
            Case enmNewFmt.FromText
                Select Case mintTextType
                    Case enmTextType.Ascii
                        mabSrcText = System.Text.Encoding.ASCII.GetBytes(StrData)
                    Case enmTextType.Unicode
                        mabSrcText = System.Text.Encoding.Unicode.GetBytes(StrData)
                    Case enmTextType.UTF8
                        mabSrcText = System.Text.Encoding.UTF8.GetBytes(StrData)
                End Select
            Case enmNewFmt.FromBase64
                mabSrcText = Convert.FromBase64String(StrData)
            Case enmNewFmt.FromHexStr
                Dim intLen As Integer = StrData.Length / 2, i As Integer
                ReDim mabSrcText(intLen - 1)
                For i = 0 To intLen - 1
                    mabSrcText(i) = CInt(StrData.Substring(i, 2))
                Next
        End Select
    End Sub
End Class
