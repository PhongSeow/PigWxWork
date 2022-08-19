'**********************************
'* Name: PigBytes
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Handle operations related to byte array division 【处理除字节数组相关的操作】
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.23
'* Create Time: 2019-10-22
'*1.0.2  2019-10-24  
'*1.0.3  2019-10-25  优化 mSetValue 和 mGetValue 等，去掉一些没有属性
'*1.0.5  2019-11-2  优化接口 GetBytesValue 等
'*1.0.6  2019-11-7  增加 ToVarDef
'*1.0.7  2019-11-11  ExtByMD5
'*1.0.8  2019-11-16  CopyTo CopyFrom
'*1.0.9  2019-11-16  GetBytes，去掉 TxRes
'*1.0.10  2019-11-19 增加压缩和解压
'*1.0.11 2019-11-27  mToVarDef
'*1.0.12 2020-2-13  修改 GetStru
'*1.0.13 2020-3-4  增加 IsMatchBytes
'*1.0.15 2020-3-5  增加 PigMD5Bytes
'*1.0.16 2020-4-13  New 通过对象序列化实例化 类名前面加上 <Serializable> 就可以序列化
'*1.0.17 2020-4-14  增加 GetDeSerializeObj
'*1.0.18 2020-6-8  增加 PigMD5
'*1.0.19 2020-6-8  优化 GetDeSerializeObj
'*1.0.20 2020-6-16 优化 IsPigMD5Mate
'*1.0.21  1/2/2021 Err.Raise change to Throw New Exception|Err.Raise改为Throw New Exception
'*1.0.22  2/2/2021 remove Formatters.Binary.BinaryFormatter 
'*1.0.23  6/5/2021 Add New(InitBase64Str)
'************************************

Imports System.Runtime.Serialization
Public Class PigBytes
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.23"

    Private mabMain As Byte()
    ''' <summary>是否转换出错则为零，如果是则调用接口不会出错</summary>
    Private mbolIsErrToZero As Boolean

    Private mlngCurrPos As Long = 0 '当前位置
    Private mbolIsAutoExt As Boolean = True  '是否自动扩展数组
    Private mabPigMD5 As Byte()
    Private mstrPigMD5 As String
    '    Private mbfMain As Formatters.Binary.BinaryFormatter
    '    Private mmsMain As System.IO.MemoryStream

    Public Sub RestPos()
        mlngCurrPos = 0
    End Sub

    Public Sub PosToEnd()
        mlngCurrPos = Me.mabMain.Length - 1
    End Sub

    Public Sub PosToAddNew()
        mlngCurrPos = Me.mabMain.Length
    End Sub

    ''' <summary>压缩数组</summary>
    Public Overloads Function UnCompress() As String
        Try
            ReDim mabPigMD5(0)
            Dim oCompressor As New PigCompressor
            Dim abAny(0) As Byte
            abAny = oCompressor.Depress(Me.mabMain)
            If oCompressor.LastErr <> "" Then Throw New Exception(oCompressor.LastErr)
            Me.mabMain = abAny
            oCompressor = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("UnCompress", ex)
        End Try
    End Function

    ''' <summary>比较数组是否一样</summary>
    Public Overloads Function IsMatchBytes(ByRef MatchBytes As Byte()) As Boolean
        Try
            If Me.mabMain.Length <> MatchBytes.Length Then Err.Raise(-1)
            Dim i As Integer, bolIsMatch As Boolean = True
            For i = 0 To MatchBytes.Length - 1
                If Me.mabMain(i) <> MatchBytes(i) Then
                    bolIsMatch = False
                    Exit For
                End If
            Next
            If bolIsMatch = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function

    ''' <summary>比较数组是否一样</summary>
    Public Overloads Function IsMatchBytes(BeginPos As Integer, ByRef MatchBytes As Byte()) As String
        Try
            If (Me.mabMain.Length - BeginPos + 1) <> MatchBytes.Length Then Err.Raise(-1)
            Dim i As Integer, bolIsMatch As Boolean = True
            For i = 0 To MatchBytes.Length - 1
                If Me.mabMain(i + BeginPos) <> MatchBytes(i) Then
                    bolIsMatch = False
                    Exit For
                End If
            Next
            If bolIsMatch = True Then
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function


    ''' <summary>压缩数组</summary>
    Public Overloads Function UnCompress(ByRef UnCompressBytes As Byte()) As String
        Try
            ReDim mabPigMD5(0)
            Dim oCompressor As New PigCompressor
            UnCompressBytes = oCompressor.Depress(Me.mabMain)
            If oCompressor.LastErr <> "" Then Throw New Exception(oCompressor.LastErr)
            oCompressor = Nothing
            Return "OK"
        Catch ex As Exception
            UnCompressBytes = Nothing
            Return Me.GetSubErrInf("UnCompress", ex)
        End Try
    End Function

    ''' <summary>压缩数组</summary>
    Public Overloads Function Compress(ByRef CompressBytes As Byte()) As String
        Try
            Dim oCompressor As New PigCompressor
            CompressBytes = oCompressor.Compress(Me.mabMain)
            If oCompressor.LastErr <> "" Then Throw New Exception(oCompressor.LastErr)
            oCompressor = Nothing
            Return "OK"
        Catch ex As Exception
            CompressBytes = Nothing
            Return Me.GetSubErrInf("Compress", ex)
        End Try
    End Function

    ''' <summary>压缩数组</summary>
    Public Overloads Function Compress() As String
        Try
            Dim oCompressor As New PigCompressor
            Dim abAny(0) As Byte
            abAny = oCompressor.Compress(Me.mabMain)
            If oCompressor.LastErr <> "" Then Throw New Exception(oCompressor.LastErr)
            Me.mabMain = abAny
            oCompressor = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Compress", ex)
        End Try
    End Function

    ''' <summary>将字节数组转换为变量定义</summary>
    ''' <param name="ItemsPerRow">一行的项目数</param>
    ''' <returns>定义串</returns>
    ''' <remarks></remarks>
    Public Overloads Function ToVarDef(ItemsPerRow As Integer) As String
        ToVarDef = Me.mToVarDef(ItemsPerRow)
    End Function

    ''' <summary>将字节数组转换为变量定义</summary>
    ''' <returns>定义串</returns>
    ''' <remarks></remarks>
    Public Overloads Function ToVarDef() As String
        ToVarDef = Me.mToVarDef(128)
    End Function

    ''' <summary>获取结构</summary>
    ''' <param name="SrcStru">源结构变量</param>
    ''' <remarks></remarks>
    Public Overloads Function SetStru(SrcStru As Object) As String
        Dim strStepName As String = ""
        Try
            Dim oStru2Bytes As New Stru2Bytes, strRet As String
            Dim abStru(0) As Byte
            strStepName = "Stru2Bytes"
            strRet = oStru2Bytes.Stru2Bytes(SrcStru, abStru)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Dim intStruLen As Integer = abStru.Length
            ReDim Me.mabMain(0)
            Me.RestPos()
            strStepName = "SetValue(intStruLen)"
            strRet = Me.SetValue(CInt(intStruLen))
            If strRet <> "OK" Then Throw New Exception(strRet)
            strStepName = "SetValue(abStru)"
            strRet = Me.SetValue(abStru)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oStru2Bytes = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("SetStru", ex)
        End Try
    End Function

    ''' <summary>获取结构</summary>
    ''' <param name="TarType">结构类型，例：结构变量.GetType</param>
    ''' <param name="TarStru">结构变量</param>
    ''' <param name="IsInit">结构是否已初始化</param>
    ''' <remarks></remarks>
    Public Overloads Function GetStru(TarType As Type, ByRef TarStru As Object, ByRef IsInit As Boolean) As String
        Dim strStepName As String = ""
        Try
            Dim oStru2Bytes As New Stru2Bytes, strRet As String
            IsInit = True
            Me.RestPos()
            strStepName = "GetInt32Value"
            Dim intStruLen As Integer = Me.GetInt32Value()
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            If intStruLen = 0 Then
                IsInit = False
                Throw New Exception("Structure not init")
            End If
            strStepName = "Bytes2Stru"
            Dim abAny(0) As Byte
            strStepName = "GetBytesValue"
            abAny = Me.GetBytesValue(intStruLen)
            If Me.LastErr <> "" Then Throw New Exception(Me.LastErr)
            strStepName = "Bytes2Stru"
            strRet = oStru2Bytes.Bytes2Stru(abAny, TarType, TarStru)
            If strRet <> "OK" Then Throw New Exception(strRet)
            oStru2Bytes = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("GetStru", ex)
        End Try
    End Function


    ''' <summary>将字节数组转换为变量定义</summary>
    ''' <param name="ItemsPerRow">一行的项目数</param>
    ''' <returns>定义串</returns>
    ''' <remarks></remarks>
    Private Function mToVarDef(ItemsPerRow As Integer) As String
        Try
            Dim i As Long
            If ItemsPerRow <= 0 Then ItemsPerRow = 128
            mToVarDef = "Dim abAny() As Byte ="
            For i = 0 To Me.mabMain.Length - 1
                If i = 0 Then
                    mToVarDef &= "{" & Me.mabMain(i).ToString
                ElseIf i Mod ItemsPerRow = 0 Then
                    mToVarDef &= " _" & vbCrLf
                Else
                    mToVarDef &= "," & Me.mabMain(i).ToString
                End If
            Next
            mToVarDef &= "}"
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mToVarDef", ex)
            mToVarDef = ""
        End Try
    End Function

    Public Overloads Function SetValue(ByRef ByteValue As Byte(), CopyLen As Integer) As String
        SetValue = Me.mCopyFrom(ByteValue, , Me.CurrPos, CopyLen)
    End Function

    Public Overloads Function SetValue(ByRef ByteValue As Byte()) As String
        SetValue = Me.mCopyFrom(ByteValue, , Me.CurrPos, ByteValue.Length)
    End Function
    Public Overloads Function SetValue(LongValue As Long) As String
        Return Me.mSetValue(mlngCurrPos, LongValue)
    End Function

    Public Overloads Function SetValue(IntegerValue As Integer) As String
        Return Me.mSetValue(mlngCurrPos, IntegerValue)
    End Function

    Public Overloads Function SetValue(DateTimeValue As DateTime) As String
        Return Me.mSetValue(mlngCurrPos, DateTimeValue)
    End Function

    Public Overloads Function SetValue(DoubleValue As Double) As String
        Return Me.mSetValue(mlngCurrPos, DoubleValue)
    End Function

    Public Overloads Function SetValue(ByteValue As Byte) As String
        Return Me.mSetValue(mlngCurrPos, ByteValue)
    End Function

    Public Overloads Function SetValue(BooleanValue As Boolean) As String
        Return Me.mSetValue(mlngCurrPos, BooleanValue)
    End Function

    Public Overloads Function GetByteValue() As Byte
        Try
            Dim strRet As String = Me.mGetValue(TypeCode.Byte, mlngCurrPos, GetByteValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetByteValue", ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function GetBytesValue(CopyLen As Long) As Byte()
        Try
            ReDim GetBytesValue(CopyLen - 1)
            Dim strRet As String = Me.mCopyTo(GetBytesValue, Me.CurrPos, , CopyLen)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetBytesValue", ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function GetBooleanValue() As Boolean
        Try
            Dim strRet As String = Me.mGetValue(TypeCode.Boolean, mlngCurrPos, GetBooleanValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetBooleanValue", ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function GetDoubleValue() As Double
        Try
            Dim strRet As String = Me.mGetValue(TypeCode.Double, mlngCurrPos, GetDoubleValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetDoubleValue", ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function GetInt32Value() As Integer
        Try
            Dim strRet As String = Me.mGetValue(TypeCode.Int32, mlngCurrPos, GetInt32Value)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetInt32Value", ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function GetInt64Value() As Long
        Try
            Dim strRet As String = Me.mGetValue(TypeCode.Int64, mlngCurrPos, GetInt64Value)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetInt64Value", ex)
            Return Nothing
        End Try
    End Function

    Public Overloads Function GetDateTimeValue() As DateTime
        Try
            Dim strRet As String = Me.mGetValue(TypeCode.DateTime, mlngCurrPos, GetDateTimeValue)
            If strRet <> "OK" Then Throw New Exception(strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetDateTimeValue", ex)
            Return Nothing
        End Try
    End Function

    Private Function mIsOverBytes(StartPos As Long, BytesLen As Long) As Boolean
        If StartPos + BytesLen > Me.mabMain.Length Then
            Return True
        Else
            Return False
        End If
    End Function

    Private Function mGetValue(TypeCode As System.TypeCode, StartPos As Long, ByRef GetValue As Object) As String
        Try
            Const IS_OVER_ARRAY As String = "Array range exceeded"
            Dim lngLen As Long = 0
            If StartPos < 0 Then StartPos = 0
            If StartPos > Me.mabMain.Length Then Throw New Exception(IS_OVER_ARRAY)
            Select Case TypeCode
                Case System.TypeCode.Boolean
                    lngLen = 1
                    If Me.mIsOverBytes(StartPos, lngLen) = True Then Throw New Exception(IS_OVER_ARRAY)
                    GetValue = BitConverter.ToBoolean(Me.mabMain, StartPos)
                Case System.TypeCode.Byte
                    lngLen = 1
                    If Me.mIsOverBytes(StartPos, lngLen) = True Then Throw New Exception(IS_OVER_ARRAY)
                    GetValue = CByte(Me.mabMain(StartPos))
                Case System.TypeCode.DateTime
                    lngLen = 8
                    If Me.mIsOverBytes(StartPos, lngLen) = True Then Throw New Exception(IS_OVER_ARRAY)
                    GetValue = DateTime.FromOADate(BitConverter.ToDouble(Me.mabMain, StartPos))
                Case System.TypeCode.Double
                    lngLen = 8
                    If Me.mIsOverBytes(StartPos, lngLen) = True Then Throw New Exception(IS_OVER_ARRAY)
                    GetValue = BitConverter.ToDouble(Me.mabMain, StartPos)
                Case System.TypeCode.Int32
                    lngLen = 4
                    If Me.mIsOverBytes(StartPos, lngLen) = True Then Throw New Exception(IS_OVER_ARRAY)
                    GetValue = BitConverter.ToInt32(Me.mabMain, StartPos)
                Case System.TypeCode.Int64
                    lngLen = 8
                    If Me.mIsOverBytes(StartPos, lngLen) = True Then Throw New Exception(IS_OVER_ARRAY)
                    GetValue = BitConverter.ToInt32(Me.mabMain, StartPos)
                Case Else
                    Throw New Exception("No support at the moment " & TypeCode.ToString)
            End Select
            mlngCurrPos += lngLen
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mGetValue", ex)
        End Try
    End Function

    Private Function mSetValue(StartPos As Long, SetValue As Object) As String
        Try
            ReDim mabPigMD5(0)
            Dim abAny As Byte()
            ReDim abAny(0)
            If StartPos < 0 Then StartPos = 0
            Dim tcAny As System.TypeCode = SetValue.GetTypeCode
            Select Case tcAny
                Case System.TypeCode.Boolean
                    abAny = BitConverter.GetBytes(SetValue)
                Case System.TypeCode.Byte
                    ReDim abAny(0)
                    abAny(0) = SetValue
                Case System.TypeCode.DateTime
                    abAny = BitConverter.GetBytes(SetValue.ToOADate)
                Case System.TypeCode.Double
                    abAny = BitConverter.GetBytes(SetValue)
                Case System.TypeCode.Int32
                    abAny = BitConverter.GetBytes(SetValue)
                Case System.TypeCode.Int64
                    abAny = BitConverter.GetBytes(SetValue)
                Case Else
                    Throw New Exception("No support at the moment " & tcAny.ToString)
            End Select
            If abAny.Length > (Me.mabMain.Length - StartPos) Then
                If Me.IsAutoExt = True Then
                    '自动扩展，新的：(位置-原长)+新数据长+原长-1=位置+新数据长-1
                    Dim lngReLen As Long = StartPos + abAny.Length - 1
                    ReDim Preserve Me.mabMain(lngReLen)
                Else
                    Throw New Exception("Set data will export this byte array overflow")
                End If
            End If
            abAny.CopyTo(Me.mabMain, StartPos)
            mlngCurrPos += abAny.Length
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mSetValue", ex)
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
            '初始化随机数生成器
            Randomize()
            mGetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            mGetRandNum = 0
        End Try
    End Function

    ''' <summary>用MD5扩充数组</summary>
    ''' <param name="ExtCnt">扩展个数，每个有16个字节</param>
    Public Function ExtByMD5(ExtCnt As Integer) As String
        Dim strStepName As String = ""
        Try
            Dim strRet As String = "", oMD5 As PigMD5
            strStepName = "abSrc=Me.mabMain"
            Dim abSrc As Byte() = Me.mabMain
            Dim i As Integer, intLen As Integer = CInt(abSrc.Length) / ExtCnt
            Dim oGEBytes As New PigBytes(Me.mabMain)
            Dim abAny(0) As Byte
            oGEBytes.RestPos()
            Me.PosToAddNew()
            For i = 0 To ExtCnt - 1
                strStepName = "GetBytesValue(" & i.ToString & ")"
                abAny = oGEBytes.GetBytesValue(intLen)
                If oGEBytes.LastErr <> "" Then Throw New Exception(oGEBytes.LastErr)
                oMD5 = New PigMD5(abAny)
                strStepName = "SetValue(" & i.ToString & ")"
                strRet = Me.SetValue(oMD5.MD5Bytes)
                If strRet <> "OK" Then Throw New Exception(strRet)
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("ExtByMD5", ex)
        End Try
    End Function

    ''' <summary>填充随机数</summary>
    Public Function FillRandNo() As String
        Try
            ReDim Me.mabMain(mabMain.Length - 1)
            For i = 0 To Me.mabMain.Length - 1
                Me.mabMain(i) = mGetRandNum(0, 255)
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("FillRandNo", ex)
        End Try
    End Function

    'Public Sub New(ByRef InitObj As Object)
    '    MyBase.New(CLS_VERSION)
    '    Dim strStepName As String = ""
    '    Try
    '        strStepName = "New BinaryFormatter And MemoryStream"
    '        mbfMain = New Formatters.Binary.BinaryFormatter
    '        mmsMain = New System.IO.MemoryStream
    '        strStepName = "Serialize"
    '        mbfMain.Serialize(mmsMain, InitObj)
    '        ReDim mabMain(mmsMain.Length - 1)
    '        strStepName = "ReadAsync"
    '        mmsMain.Position = 0
    '        'mmsMain.ReadAsync(mabMain, 0, mmsMain.Length - 1)
    '        mmsMain.Read(mabMain, 0, mmsMain.Length - 1)
    '        mmsMain.Close()
    '        Me.ClearErr()
    '    Catch ex As Exception
    '        mbfMain = Nothing
    '        mmsMain = Nothing
    '        Me.SetSubErrInf("New", strStepName, ex)
    '    End Try
    'End Sub

    Public Sub New(InitLen As Long)
        MyBase.New(CLS_VERSION)
        If InitLen < 0 Then InitLen = 0
        ReDim Me.mabMain(InitLen - 1)
        Me.RestPos()
    End Sub

    Public Sub New(InitBytes As Byte())
        MyBase.New(CLS_VERSION)
        Me.mCopyFrom(InitBytes)
        Me.RestPos()
    End Sub

    Public Sub New(InitBase64Str As String)
        MyBase.New(CLS_VERSION)
        Me.mabMain = Convert.FromBase64String(InitBase64Str)
        ReDim Me.mabPigMD5(0)
    End Sub

    Public Sub New(InitBytes As Byte(), SrcStartPos As Long, CopyLen As Long)
        MyBase.New(CLS_VERSION)
        Me.mCopyFrom(InitBytes, 0, SrcStartPos, CopyLen)
        Me.RestPos()
    End Sub

    Public Sub New(InitBytes As Byte(), CopyLen As Long)
        MyBase.New(CLS_VERSION)
        Me.mCopyFrom(InitBytes, 0, 0, CopyLen)
        Me.RestPos()
    End Sub

    '''' <summary>获取反序列化对象，引用格式：CType(GetDeSerializeObj,对象类型)</summary>
    'Public Function GetDeSerializeObj() As Object
    '    Dim strStepName As String = ""
    '    Try
    '        strStepName = "Determine serialization preparation"
    '        mbfMain = New Formatters.Binary.BinaryFormatter
    '        mmsMain = New System.IO.MemoryStream
    '        '            If mbfMain Is Nothing Or mmsMain Is Nothing Then Err.Raise(-1,, "未初始化")
    '        If mabMain Is Nothing Then Throw New Exception("Array not initialized")
    '        strStepName = "WriteAsync"
    '        mmsMain.Position = 0
    '        mmsMain.Flush()
    '        mmsMain.Write(mabMain, 0, mabMain.Length - 1)
    '        strStepName = "Deserialize"
    '        mmsMain.Position = 0
    '        'mmsMain.Seek(0, IO.SeekOrigin.Begin)
    '        GetDeSerializeObj = mbfMain.Deserialize(mmsMain)
    '        Me.ClearErr()
    '    Catch ex As Exception
    '        Me.SetSubErrInf("GetDeSerializeObj", strStepName, ex)
    '        Return Nothing
    '    End Try
    'End Function

    ''' <summary>复制数组到本类</summary>
    ''' <param name="SrcBytes">源数组</param>
    ''' <param name="CopyLen">源数组</param>
    ''' <remarks>SrcBytes：源数组，SrcStartPos：源数组开始位置，TarStartPos：目标数组开始位置，CopyLen：复制长度</remarks>
    Public Function CopyFrom(ByRef SrcBytes As Byte(), CopyLen As Long) As String
        CopyFrom = Me.mCopyFrom(SrcBytes, 0, 0, CopyLen)
    End Function

    ''' <summary>复制数组到本类</summary>
    ''' <param name="SrcBytes">源数组</param>
    ''' <param name="SrcStartPos">源数组</param>
    ''' <param name="CopyLen">源数组</param>
    ''' <remarks>SrcBytes：源数组，SrcStartPos：源数组开始位置，TarStartPos：目标数组开始位置，CopyLen：复制长度</remarks>
    Public Function CopyFrom(ByRef SrcBytes As Byte(), SrcStartPos As Long, CopyLen As Long) As String
        CopyFrom = Me.mCopyFrom(SrcBytes, SrcStartPos, 0, CopyLen)
    End Function

    ''' <summary>复制数组到本类</summary>
    ''' <param name="SrcBytes">源数组</param>
    ''' <param name="SrcStartPos">源数组起始位置</param>
    ''' <param name="TarStartPos">目标数组起始位置</param>
    ''' <param name="CopyLen">复制长度</param>
    Public Function CopyFrom(ByRef SrcBytes As Byte(), Optional SrcStartPos As Long = 0, Optional TarStartPos As Long = 0, Optional CopyLen As Long = 0) As String
        CopyFrom = Me.mCopyFrom(SrcBytes, SrcStartPos, TarStartPos, CopyLen)
    End Function

    ''' <summary>复制数组到本类</summary>
    ''' <param name="SrcBytes">源数组</param>
    ''' <param name="SrcStartPos">源数组起始位置</param>
    ''' <param name="TarStartPos">目标数组起始位置</param>
    ''' <param name="CopyLen">复制长度</param>
    ''' <remarks>SrcBytes：源数组，SrcStartPos：源数组开始位置，TarStartPos：目标数组开始位置，CopyLen：复制长度</remarks>
    Private Function mCopyFrom(ByRef SrcBytes As Byte(), Optional SrcStartPos As Long = 0, Optional TarStartPos As Long = 0, Optional CopyLen As Long = 0) As String
        Try
            ReDim mabPigMD5(0)
            If SrcStartPos < 0 Then SrcStartPos = 0
            If TarStartPos < 0 Then TarStartPos = 0
            If CopyLen <= 0 Then CopyLen = SrcBytes.Length
            If Me.mabMain Is Nothing Then ReDim Me.mabMain(0)
            If (TarStartPos + CopyLen - 1) > Me.mabMain.Length Then
                ReDim Preserve Me.mabMain(TarStartPos + CopyLen - 1)
            End If
            If SrcStartPos = 0 Then
                SrcBytes.CopyTo(Me.mabMain, TarStartPos)
            Else
                Dim i As Long
                For i = 0 To CopyLen - 1
                    Me.mabMain(i) = SrcBytes(i + SrcStartPos)
                Next
            End If
            mlngCurrPos += CopyLen
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mCopyFrom", ex, False)
        End Try
    End Function

    Public Sub New()
        MyBase.New(CLS_VERSION)
        ReDim Me.mabMain(0)
        ReDim Me.mabPigMD5(0)
    End Sub

    Private Function mSetPigMD5() As String
        Try
            Dim oMD5 As New PigMD5(Me.mabMain)
            mabPigMD5 = oMD5.PigMD5Bytes
            mstrPigMD5 = oMD5.PigMD5
            oMD5 = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mSetPigMD5", ex)
        End Try
    End Function

    ''' <summary>豚豚MD5字节数组</summary>
    Public Overloads ReadOnly Property PigMD5Bytes As Byte()
        Get
            Try
                If mabPigMD5.Length <> 16 Then
                    Me.mSetPigMD5()
                End If
                Return mabPigMD5
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>豚豚MD5字符串</summary>
    Public Overloads ReadOnly Property PigMD5 As String
        Get
            Try
                If Len(mstrPigMD5) <> 16 Then
                    Me.mSetPigMD5()
                End If
                Return mstrPigMD5
            Catch ex As Exception
                Return Space(16)
            End Try
        End Get
    End Property

    ''' <summary>Base64字符串</summary>
    Public Overloads ReadOnly Property Base64Str As String
        '8字节
        Get
            Base64Str = Convert.ToBase64String(Me.mabMain)
        End Get
    End Property

    ''' <remarks>16进制字符串</remarks>
    Public ReadOnly Property HexStr(Optional SplitChar As String = "|") As String
        '8字节
        Get
            Dim intLen As Integer = Me.mabMain.Length, i As Integer
            Dim sbAny As New System.Text.StringBuilder("")
            For i = 0 To intLen - 1
                If i > 0 Then sbAny.Append(SplitChar)
                sbAny.Append(Me.mabMain(i).ToString("x2"))
            Next
            HexStr = sbAny.ToString
            sbAny = Nothing
        End Get
    End Property

    ''' <remarks>当前数组位置</remarks>
    Public Property CurrPos As Long
        Get
            CurrPos = mlngCurrPos
        End Get
        Set(value As Long)
            Select Case value
                Case Is < 0
                    mlngCurrPos = 0
                Case Is > Me.mabMain.Length - 1
                    mlngCurrPos = Me.mabMain.Length - 1
            End Select
        End Set
    End Property

    ''' <remarks>是否自动扩展数组</remarks>
    Public Property IsAutoExt As String
        Get
            IsAutoExt = mbolIsAutoExt
        End Get
        Set(value As String)
            mbolIsAutoExt = value
        End Set
    End Property


    ''' <remarks>是否转换出错则为零</remarks>
    Public Property IsErrToZero As Boolean
        Get
            IsErrToZero = mbolIsErrToZero
        End Get
        Set(value As Boolean)
            mbolIsErrToZero = value
        End Set
    End Property

    ''' <remarks>主数组</remarks>
    Public Property Main() As Byte()
        Get
            Try
                Return mabMain
            Catch ex As Exception
                Return Nothing
            End Try
        End Get
        Set(value As Byte())
            mabMain = value
            ReDim mabPigMD5(0)
        End Set
    End Property


    ''' <remarks>复制数组到其他数组</remarks>
    Public Overloads Function CopyTo(ByRef TarBytes As Byte(), Optional SrcStartPos As Long = 0, Optional TarStartPos As Long = 0, Optional CopyLen As Long = 0) As String
        CopyTo = Me.mCopyTo(TarBytes, SrcStartPos, TarStartPos, CopyLen)
    End Function

    ''' <remarks>复制数组到其他数组</remarks>
    Public Overloads Function CopyTo(ByRef TarBytes As Byte(), Optional SrcStartPos As Long = 0, Optional CopyLen As Long = 0) As String
        CopyTo = Me.mCopyTo(TarBytes, SrcStartPos, 0, CopyLen)
    End Function

    ''' <remarks>复制数组到其他数组</remarks>
    Public Overloads Function CopyTo(ByRef TarBytes As Byte(), CopyLen As Long) As String
        CopyTo = Me.mCopyTo(TarBytes, 0, 0, CopyLen)
    End Function

    ''' <summary>获取主数组的一部分</summary>
    ''' <param name="StartPos">开始位置</param>
    ''' <param name="GetLen">获取长度</param>
    Public Overloads Function GetBytes(StartPos As Long, GetLen As Long) As Byte()
        Try
            GetBytes = Nothing
            If StartPos < 0 Then Throw New Exception("Start position cannot be less than 0")
            Dim strRet As String = Me.mCopyTo(GetBytes, StartPos, 0, GetLen)
            Me.ClearErr()
        Catch ex As Exception
            GetBytes = Nothing
            Me.SetSubErrInf("GetBytes", ex)
        End Try
    End Function

    ''' <summary>获取主数组从开始位置到最后</summary>
    ''' <param name="StartPos">开始位置</param>
    Public Overloads Function GetBytes(StartPos As Long) As Byte()
        Try
            GetBytes = Nothing
            If StartPos < 0 Then Throw New Exception("Start position cannot be less than 0")
            Dim strRet As String = Me.mCopyTo(GetBytes, StartPos, 0, Me.mabMain.Length - StartPos + 1)
            Me.ClearErr()
        Catch ex As Exception
            GetBytes = Nothing
            Me.SetSubErrInf("GetBytes", ex)
        End Try
    End Function


    ''' <remarks>复制数组到其他数组,TarBytes：目标数组，SrcStartPos：源数组开始位置，TarStartPos：目标数组开始位置，CopyLen：复制长度</remarks>
    Private Function mCopyTo(ByRef TarBytes As Byte(), Optional SrcStartPos As Long = 0, Optional TarStartPos As Long = 0, Optional CopyLen As Long = 0) As String
        Try
            If SrcStartPos < 0 Then SrcStartPos = 0
            If TarStartPos < 0 Then TarStartPos = 0
            If (TarStartPos + CopyLen) > TarBytes.Length Then
                '如果复制后目标数组装不下，则要扩大
                ReDim Preserve TarBytes(TarStartPos + CopyLen - 1)
            End If

            If SrcStartPos = 0 And CopyLen = 0 Then
                Me.mabMain.CopyTo(TarBytes, TarStartPos)
            Else
                Dim i As Long, lngmabMainMax As Integer = Me.mabMain.Length - 1
                For i = 0 To CopyLen - 1
                    If (i + SrcStartPos) <= lngmabMainMax Then
                        TarBytes(i) = Me.mabMain(i + SrcStartPos)
                    End If
                Next
            End If
            mlngCurrPos += CopyLen
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mCopyTo", ex, False)
        End Try
    End Function

    Public Function IsPigMD5Mate(ByRef ToMateMD5Bytes As Byte()) As Boolean
        Try
            IsPigMD5Mate = True
            If ToMateMD5Bytes.Length = 16 Then
                For i = 0 To 15
                    If Me.PigMD5Bytes(i) <> ToMateMD5Bytes(i) Then
                        IsPigMD5Mate = False
                        Exit For
                    End If
                Next
            Else
                IsPigMD5Mate = False
            End If
        Catch ex As Exception
            Return False
        End Try
    End Function
    'Protected Overrides Sub Finalize()
    '    If Not mbfMain Is Nothing Then mbfMain = Nothing
    '    If Not mmsMain Is Nothing Then mmsMain = Nothing
    '    MyBase.Finalize()
    'End Sub
End Class
