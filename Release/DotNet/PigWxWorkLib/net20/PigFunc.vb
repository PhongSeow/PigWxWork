'**********************************
'* Name: PigFunc
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Some common functions|一些常用的功能函数
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.18
'* Create Time: 2/2/2021
'*1.0.2  1/3/2021   Add UrlEncode,UrlDecode
'*1.0.3  20/7/2021   Add GECBool,GECLng
'*1.0.4  26/7/2021   Modify UrlEncode
'*1.0.5  26/7/2021   Modify UrlEncode
'*1.0.6  24/8/2021   Modify GetIpList
'*1.1    4/9/2021    Add Date2Lng,Lng2Date,Src2CtlStr,CtlStr2Src,AddMultiLineText
'*1.2    2/1/2022    Modify IsFileExists
'*1.3    12/1/2022   Add GetHostName,GetHostIp,mGetHostIp,GetEnvVar,GetUserName,GetComputerName,mGetHostIpList
'*1.4    23/1/2022   Add IsOsWindows,MyOsCrLf,MyOsPathSep
'*1.5    3/2/2022   Add GetFileText,SaveTextToFile, modify GetFilePart
'*1.6    3/2/2022   Add GetFmtDateTime, modify GENow
'*1.7    13/2/2022   Add DeleteFolder,DeleteFile
'*1.8    23/2/2022   Add PLSqlCsv2Bcp
'*1.9    20/3/2022   Modify GetProcThreadID
'*1.10   2/4/2022   Add GetHumanSize
'*1.11   9/4/2022   Add DigitalToChnName,ConvertHtmlStr,GetAlignStr,GetRepeatStr
'*1.12   14/5/2022  Rename and modify OptLogInfMain to mOptLogInfMain,SaveTextToFile to mSaveTextToFile, add ASyncOptLogInf,ASyncSaveTextToFile
'*1.13   15/5/2022  Add ASyncRet_SaveTextToFile, modify mASyncSaveTextToFile,ASyncSaveTextToFile
'*1.14   16/5/2022  Modify mASyncSaveTextToFile,ASyncSaveTextToFile
'*1.15   31/5/2022  Add GEInt, modify GECBool
'*1.16   5/7/2022   Modify GetHostIp
'*1.17   6/7/2022   Add GetFileVersion
'*1.18   19/7/2022  Add GetFileUpdateTime,GetFileCreateTime,GetFileMD5
'**********************************
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Environment
Imports System.Threading


Public Class PigFunc
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.18.2"

    Public Event ASyncRet_SaveTextToFile(SyncRet As StruASyncRet)

    ''' <summary>对齐方式|Alignment</summary>
    Public Enum EnmAlignment
        Left = 1
        Right = 2
        Center = 3
    End Enum


    ''' <summary>文件的部分</summary>
    Public Enum EnmFilePart
        Path = 1         '路径
        FileTitle = 2    '文件名
        ExtName = 3      '扩展名
        DriveNo = 4      '驱动器名
    End Enum

    ''' <summary>获取随机字符串的方式</summary>
    Public Enum EnmGetRandString
        NumberOnly = 1      '只有数字
        NumberAndLetter = 2 '只有数字和字母(包括大小写)
        DisplayChar = 3     '全部可显示字符(ASCII 33-126)
        AllAsciiChar = 4    '全部ASCII码(返回结果以16进制方式显示)
    End Enum

    ''' <summary>
    ''' 转换HTML标记的方式|How to convert HTML tags
    ''' </summary>
    Public Enum EmnHowToConvHtml
        DisableHTML = 1            '使HTML标记失效(空格、制表符及回车符都会被转换)
        EnableHTML = 2             '使HTML标记生效(空格会被还原，但制表符及回车符不会被还原)
        DisableHTMLOnlySymbol = 3  '使HTML标记失效(只有标记的符号本身受影响，空格、制表符及回车符不会被转换)
        EnableHTMLOnlySymbol = 4   '使HTML标记生效(只有标记的符号本身被还原)
    End Enum


    Sub New()
        MyBase.New(CLS_VERSION)
    End Sub


    ''' <remarks>获取比率的描述</remarks>
    Public Overloads Function GetRateDesc(Rate As Decimal) As String
        GetRateDesc = Math.Round(Rate * 100, 2) & " %"
    End Function

    ''' <remarks>获取比率的描述</remarks>
    Public Overloads Function GetRateDesc(Rate As Decimal, Optional RoundNo As Integer = 2) As String
        GetRateDesc = Math.Round(Rate * 100, RoundNo) & " %"
    End Function

    ''' <remarks>同步等待</remarks>
    Public Sub Delay(DelayMs As Single)
        System.Threading.Thread.Sleep(DelayMs)
    End Sub

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath)
    End Function


    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = True
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    ''' <remarks>写日志</remarks>
    Private Function mASyncOptLogInfMain(OptStr As String, LogFilePath As String, Optional IsShowLocalInf As Boolean = True, Optional LineBufSize As Integer = 10240) As String
        Try
            Dim struMain As mStruOptLogInfMain
            With struMain
                .OptStr = OptStr
                .LogFilePath = LogFilePath
                .IsShowLocalInf = IsShowLocalInf
                .LineBufSize = LineBufSize
            End With
            Dim oThread As New Thread(AddressOf mOptLogInfMain)
            oThread.Start(struMain)
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mASyncOptLogInfMain", ex)
        End Try
    End Function

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath, IsShowLocalInf)
    End Function

    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = IsShowLocalInf
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String, LineBufSize As Integer) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath, , LineBufSize)
    End Function


    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, LineBufSize As Integer) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = True
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    ''' <remarks>异步写日志</remarks>
    Public Overloads Function ASyncOptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean, LineBufSize As Integer) As String
        Return Me.mASyncOptLogInfMain(OptStr, LogFilePath, IsShowLocalInf, LineBufSize)
    End Function
    ''' <remarks>写日志</remarks>
    Public Overloads Function OptLogInf(OptStr As String, LogFilePath As String, IsShowLocalInf As Boolean, LineBufSize As Integer) As String
        Dim struMain As mStruOptLogInfMain
        With struMain
            .OptStr = OptStr
            .LogFilePath = LogFilePath
            .IsShowLocalInf = IsShowLocalInf
            .LineBufSize = LineBufSize
        End With
        OptLogInf = Me.mOptLogInfMain(struMain)
    End Function

    Private Structure mStruOptLogInfMain
        Public OptStr As String
        Public LogFilePath As String
        Public IsShowLocalInf As Boolean
        Public LineBufSize As Integer
    End Structure

    Private Function mOptLogInfMain(StruMain As mStruOptLogInfMain) As String
        Dim LOG As New PigStepLog("mOptLogInfMain")
        Try
            LOG.StepName = "Check StruOptLogInfMain"
            With StruMain
                If .LineBufSize <= 0 Then .LineBufSize = 10240
                If .LogFilePath = "" Then Throw New Exception("LogFilePath invalid")
            End With
            LOG.StepName = "New FileStream"
            Dim sfAny As New FileStream(StruMain.LogFilePath, System.IO.FileMode.Append, System.IO.FileAccess.Write, System.IO.FileShare.Write, StruMain.LineBufSize, False)
            LOG.StepName = "New StreamWriter"
            Dim swAny = New StreamWriter(sfAny)
            If StruMain.IsShowLocalInf = True Then StruMain.OptStr = "[" & GENow() & "][" & GetProcThreadID() & "]" & StruMain.OptStr
            LOG.StepName = "WriteLine"
            swAny.WriteLine(StruMain.OptStr)
            LOG.StepName = "Close"
            swAny.Close()
            sfAny.Close()
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(StruMain.LogFilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <remarks>获取进程及线程标识</remarks>
    Public Function GetProcThreadID() As String
        Return System.Diagnostics.Process.GetCurrentProcess.Id.ToString & "." & System.Threading.Thread.CurrentThread.ManagedThreadId.ToString
    End Function

    ''' <remarks>获取进程号</remarks>
    Public Function GetProcID() As Long
        Return Me.fMyPID
    End Function


    ''' <remarks>获取32位MD5字符串</remarks>
    Public Function GEMD5(SrcStr As String) As String
        Dim bytSrc2Hash As Byte() = (New System.Text.ASCIIEncoding).GetBytes(SrcStr)
        Dim bytHashValue As Byte() = CType(System.Security.Cryptography.CryptoConfig.CreateFromName("MD5"), System.Security.Cryptography.HashAlgorithm).ComputeHash(bytSrc2Hash)
        Dim i As Integer
        GEMD5 = ""
        For i = 0 To 15 '选择32位字符的加密结果
            GEMD5 += Right("00" & Hex(bytHashValue(i)).ToLower, 2)
        Next
    End Function

    ''' <remarks>获取16位MD5字符串</remarks>
    Public Function Get16BitMD5(SrcStr As String) As String
        Get16BitMD5 = Mid(GEMD5(SrcStr), 9, 16)
    End Function

    ''' <remarks>显示精确到毫秒的时间</remarks>
    Public Function GENow() As String
        Return Format(Now, "yyyy-MM-dd HH:mm:ss.fff")
    End Function

    Public Function GetFmtDateTime(SrcTime As DateTime, Optional TimeFmt As String = "yyyy-MM-dd HH:mm:ss.fff") As String
        Return Format(SrcTime, TimeFmt)
    End Function

    Public Function GetHostName() As String
        Return Dns.GetHostName()
    End Function

    Public Function GetHostIp() As String
        Return mGetHostIp(False)
    End Function

    Public Function GetHostIpList() As String
        Return mGetHostIpList(False)
    End Function

    Public Function GetHostIpList(IsIPv6 As Boolean) As String
        Return mGetHostIpList(IsIPv6)
    End Function

    ''' <summary>
    ''' 获取主机的IP|Get the IP of the host
    ''' </summary>
    ''' <param name="PriorityIpHead">优先匹配的IP地址开头，如果匹配不到则取第一个IP|The first IP address to be matched first. If it cannot be matched, the first IP will be taken</param>
    ''' <returns></returns>
    Public Function GetHostIp(PriorityIpHead As String) As String
        Try
            Dim strIpList As String = Me.mGetHostIpList(False)
            Dim strFirstIp As String = ""
            GetHostIp = ""
            Do While True
                Dim strIp As String = Me.GetStr(strIpList, "", ";")
                If strIp = "" Then Exit Do
                If strIp <> "127.0.0.1" Then
                    If strFirstIp = "" Then strFirstIp = strIp
                    If Left(strIp, Len(PriorityIpHead)) = PriorityIpHead Then
                        GetHostIp = strIp
                        Exit Do
                    End If
                End If
            Loop
            If GetHostIp = "" Then GetHostIp = strFirstIp
        Catch ex As Exception
            Me.SetSubErrInf("GetHostIp", ex)
            Return ""
        End Try
    End Function


    Public Function GetHostIp(IsIPv6 As Boolean) As String
        Return mGetHostIp(IsIPv6)
    End Function

    Public Function GetHostIp(IsIPv6 As Boolean, IpHead As String) As String
        Return mGetHostIp(IsIPv6, IpHead)
    End Function

    Private Function mGetHostIpList(IsIPv6 As Boolean) As String
        mGetHostIpList = ""
        For Each oIPAddress As IPAddress In Dns.GetHostAddresses(Dns.GetHostName)
            With oIPAddress
                Dim strHostIp = .ToString()
                If IsIPv6 = True Then
                    If InStr(strHostIp, ":") > 0 Then mGetHostIpList &= strHostIp & ";"
                ElseIf InStr(strHostIp, ".") > 0 Then
                    mGetHostIpList &= strHostIp & ";"
                End If
            End With
        Next
    End Function

    Private Function mGetHostIp(IsIPv6 As Boolean, Optional IpHead As String = "") As String
        mGetHostIp = ""
        For Each oIPAddress As IPAddress In Dns.GetHostAddresses(Dns.GetHostName)
            With oIPAddress
                mGetHostIp = .ToString()
                If IsIPv6 = True Then
                    If InStr(mGetHostIp, ":") > 0 Then
                        If IpHead = "" Then
                            Exit For
                        ElseIf UCase(Left(mGetHostIp, Len(IpHead))) = UCase(IpHead) Then
                            Exit For
                        End If
                    End If
                ElseIf InStr(mGetHostIp, ".") > 0 Then
                    If IpHead = "" Then
                        Exit For
                    ElseIf Left(mGetHostIp, Len(IpHead)) = IpHead Then
                        Exit For
                    End If
                End If
            End With
            mGetHostIp = ""
        Next
    End Function

    ''' <remarks>获取随机数</remarks>
    Public Function GetRandNum(BeginNum As Integer, EndNum As Integer) As Integer
        Dim i As Long
        Try
            If BeginNum > EndNum Then
                i = BeginNum
                BeginNum = EndNum
                EndNum = i
            End If
            Randomize()   '初始化随机数生成器
            GetRandNum = Int((EndNum - BeginNum + 1) * Rnd() + BeginNum)
        Catch ex As Exception
            GetRandNum = 0
        End Try
    End Function

    ''' <remarks>产生随机字符串</remarks>
    Public Function GetRandString(StrLen As Integer, Optional MemberType As enmGetRandString = enmGetRandString.DisplayChar) As String
        Dim i As Integer
        Dim intChar As Integer
        Try
            GetRandString = ""
            For i = 1 To StrLen
                Select Case MemberType
                    Case enmGetRandString.AllAsciiChar
                        intChar = GetRandNum(0, 255)
                    Case enmGetRandString.DisplayChar    '!-~
                        intChar = GetRandNum(33, 126)
                    Case enmGetRandString.NumberAndLetter
                        intChar = GetRandNum(1, 3)
                        Select Case intChar
                            Case 1  '0-9
                                intChar = GetRandNum(48, 57)
                            Case 2  'A-Z
                                intChar = GetRandNum(65, 90)
                            Case 3  'a-z
                                intChar = GetRandNum(97, 122)
                        End Select
                    Case enmGetRandString.NumberOnly
                        intChar = GetRandNum(48, 57)
                End Select
                If MemberType = enmGetRandString.AllAsciiChar Then
                    GetRandString = GetRandString & Right("0" & Hex(intChar), 2)
                Else
                    GetRandString = GetRandString & Chr(intChar)
                End If
            Next
        Catch ex As Exception
            GetRandString = ""
        End Try

    End Function

    ''' <remarks>生成主键值，唯一</remarks>
    Public Function GetPKeyValue(SrcKey As String, Is32Bit As Boolean) As String
        Dim strRndStr As String, strRndKey As String, i As Long, strTmp As String, strChar As String, strChar2 As String
        Dim intCharAscAdd As Integer, intStrLen As Integer
        GetPKeyValue = ""
        Try
            If Is32Bit = True Then
                intStrLen = 32
            Else
                intStrLen = 16
            End If
            strRndStr = GetRandString(64, enmGetRandString.DisplayChar)
            strRndKey = GetRandString(intStrLen, enmGetRandString.NumberOnly)
            strTmp = GENow() & "-" & SrcKey & "-" & strRndStr
            If Is32Bit = True Then
                strTmp = GEMD5(strTmp)
            Else
                strTmp = Get16BitMD5(strTmp)
            End If
            strTmp = LCase(strTmp)
            For i = 1 To intStrLen
                strChar = Mid(strTmp, i, 1)
                strChar2 = Mid(strRndKey, i, 1)
                intCharAscAdd = 0
                Select Case strChar2
                    Case "0"    '不变
                    Case "1"    '数字不变，字母变大写
                        Select Case strChar
                            Case "0" To "9"
                            Case Else
                                intCharAscAdd = -32
                        End Select
                    Case "2" To "5"    '位移到 g-j
                        Select Case strChar
                            Case "0" To "9"
                                intCharAscAdd = (Asc("g") + Int(strChar2) - 1) - Asc("0")
                            Case Else
                                intCharAscAdd = (Asc("g") + Int(strChar2) - 1) - Asc("a")
                        End Select
                    Case "6" To "9"    '位移到 G-J
                        Select Case strChar
                            Case "0" To "9"
                                intCharAscAdd = (Asc("G") + Int(strChar2) - 1) - Asc("0")
                            Case Else
                                intCharAscAdd = (Asc("G") + Int(strChar2) - 1) - Asc("a")
                        End Select
                    Case Else   '不变
                End Select
                strChar2 = Chr(Asc(strChar) + intCharAscAdd)
                Select Case strChar2
                    Case "0" To "9"
                        strChar = strChar2
                    Case "a" To "z"
                        strChar = strChar2
                    Case "A" To "Z"
                        strChar = strChar2
                    Case Else
                End Select
                GetPKeyValue &= strChar
            Next
        Catch ex As Exception
            GetPKeyValue = ""
        End Try
    End Function

    ''' <remarks>获取IP列表，主IP</remarks>
    Public Sub GetIpList(ByRef IpList As String, ByRef MainIp As String, MainIpRole As String)
        'IpList:本机IP列表
        'MainIp:主IP
        'MainIpRole:主IP规则，如：56.0.，如果IP列表中有第一个匹配的就是主IP
        Try
            Dim aipaAny() As System.Net.IPAddress, strIp As String = "", intFindCnt As Integer = 0, intLen As Integer = 0
            aipaAny = System.Net.Dns.GetHostAddresses(System.Environment.MachineName.ToString)
            IpList = ""
            MainIp = ""
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
            Dim i As Integer 
            For i = 0 To aipaAny.Count - 1
                strIp = aipaAny(i).ToString
                If InStr(strIp, "::") = 0 Then
                    '这种IP没有有效的IP4地址，如：fe80::95fa:649f:cbab:5912%21，不要

                    IpList &= strIp & ";"
                    intFindCnt += 1
                    If intFindCnt = 1 Then
                        '第一个先假定为主IP
                        MainIp = strIp
                    Else
                        intLen = Len(MainIpRole)
                        If intLen > 0 Then
                            If Microsoft.VisualBasic.Left(strIp, Len(MainIpRole)) = MainIpRole Then
                                'IP的前N位与主IP规则匹配
                                MainIp = strIp
                            End If
                        End If
                    End If
                End If
            Next
#End If
        Catch ex As Exception
            IpList = ""
            MainIp = ""
        End Try
    End Sub

    ''' <remarks>获取文件路径的组成部分</remarks>
    Public Function GetFilePart(ByVal FilePath As String, Optional FilePart As enmFilePart = enmFilePart.FileTitle) As String
        Dim strTemp As String, i As Long, lngLen As Long
        Dim strPath As String = "", strFileTitle As String = ""
        Dim strOsPathSep As String = Me.OsPathSep
        Try
            GetFilePart = ""
            Select Case FilePart
                Case enmFilePart.DriveNo
                    GetFilePart = GetStr(FilePath, "", ":", False)
                    If GetFilePart = "" Then
                        GetFilePart = GetStr(FilePath, "", "$", False)
                        If GetFilePart <> "" Then GetFilePart = GetFilePart & "$"
                    End If
                Case enmFilePart.ExtName
                    lngLen = Len(FilePath)
                    For i = lngLen To 1 Step -1
                        Select Case Mid(FilePath, i, 1)
                            Case "/", ":", "$", "\"
                                Exit For
                            Case "."
                                GetFilePart = Mid(FilePath, i + 1)
                                Exit For
                        End Select

                    Next
                Case enmFilePart.FileTitle, enmFilePart.Path
                    Do While True
                        strTemp = GetStr(FilePath, "", strOsPathSep, True)
                        If Len(strTemp) = 0 Then
                            If Right(strPath, 1) = strOsPathSep Then
                                If Right(strPath, 2) <> ":" & strOsPathSep Then
                                    strPath = Left(strPath, Len(strPath) - 1)
                                End If
                            ElseIf Left(FilePath, 1) = strOsPathSep Then
                                strPath = strOsPathSep
                                FilePath = Mid(FilePath, 2)
                            End If
                            If FilePath <> "" Then
                                strFileTitle = FilePath
                            Else
                                strFileTitle = "."
                            End If
                            Exit Do
                        End If
                        strPath = strPath & strTemp & strOsPathSep
                    Loop
                    If FilePart = enmFilePart.FileTitle Then
                        GetFilePart = strFileTitle
                    Else
                        GetFilePart = strPath
                    End If
                Case Else
                    GetFilePart = ""
            End Select
        Catch ex As Exception
            GetFilePart = ""
        End Try
    End Function

    ''' <remarks>截取字符串</remarks>
    Public Function GetStr(ByRef SourceStr As String, strBegin As String, strEnd As String, Optional IsCut As Boolean = True) As String
        Dim lngBegin As Long
        Dim lngEnd As Long
        Dim lngBeginLen As Long
        Dim lngEndLen As Long
        Try
            lngBeginLen = Len(strBegin)
            lngBegin = InStr(SourceStr, strBegin, CompareMethod.Text)
            lngEndLen = Len(strEnd)
            If lngEndLen = 0 Then
                lngEnd = Len(SourceStr) + 1
            Else
                lngEnd = InStr(lngBegin + lngBeginLen + 1, SourceStr, strEnd, CompareMethod.Text)
                If lngBegin = 0 Then Err.Raise(-1, , "lngBegin=0")
            End If
            If lngEnd <= lngBegin Then Err.Raise(-1, , "lngEnd <= lngBegin")
            If lngBegin = 0 Then Err.Raise(-1, , "lngBegin=0[2]")

            GetStr = Mid(SourceStr, lngBegin + lngBeginLen, (lngEnd - lngBegin - lngBeginLen))
            If IsCut = True Then
                SourceStr = Left(SourceStr, lngBegin - 1) & Mid(SourceStr, lngEnd + lngEndLen)
            End If
            Me.ClearErr()
        Catch ex As Exception
            GetStr = ""
            Me.SetSubErrInf("GetStr", ex)
        End Try
    End Function



    Public Function UrlEncode(SrcUrl As String) As String
        Try
#If NETCOREAPP Or NET5_0_OR_GREATER Then
            UrlEncode = System.Web.HttpUtility.UrlEncode(SrcUrl)
#Else
            UrlEncode = System.Uri.EscapeDataString(SrcUrl)
#End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("UrlEncode", ex)
            Return Nothing
        End Try
    End Function

    Public Function UrlDecode(DecodeUrl As String) As String
        Try
#If NETCOREAPP Or NET5_0_OR_GREATER Then
            UrlDecode = System.Web.HttpUtility.UrlDecode(DecodeUrl)
#Else
            UrlDecode = System.Uri.UnescapeDataString(DecodeUrl)
#End If

            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("UrlDecode", ex)
            Return Nothing
        End Try
    End Function

    Public Function IsRegexMatch(SrcStr As String, RegularExp As String) As Boolean
        Return System.Text.RegularExpressions.Regex.IsMatch(SrcStr, RegularExp)
    End Function

    Public Function GECBool(vData As String) As Boolean
        Try
            Return CBool(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECBool", ex)
            Return False
        End Try
    End Function

    Public Function GECBool(vData As String, IsEmptyTrue As Boolean) As Boolean
        Try
            If IsEmptyTrue = True And vData = "" Then
                Return True
            Else
                Return CBool(vData)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("GECBool", ex)
            Return False
        End Try
    End Function

    Public Function GEInt(vData As String) As Integer
        Try
            Return CInt(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECLng", ex)
            Return 0
        End Try
    End Function

    Public Function GECLng(vData As String) As Long
        Try
            Return CLng(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECLng", ex)
            Return 0
        End Try
    End Function

    Public Function GEDec(vData As String) As Decimal
        Try
            Return CDec(vData)
        Catch ex As Exception
            Me.SetSubErrInf("CDec", ex)
            Return 0
        End Try
    End Function

    Public Function GECDate(vData As String) As DateTime
        Try
            Return Date.Parse(vData)
        Catch ex As Exception
            Me.SetSubErrInf("GECDate", ex)
            Return DateTime.MinValue
        End Try
    End Function

    ''' <summary>
    ''' Gets the number of microseconds of Greenwich mean time for the current time|获取当前时间的格林威治时间微秒数
    ''' </summary>
    ''' <param name="DateValue"></param>
    ''' <returns></returns>
    Public Function Date2Lng(DateValue As DateTime) As Long
        Dim dteStart As New DateTime(1970, 1, 1)
        Dim mtsTimeDiff As TimeSpan = DateValue - dteStart
        Try
            Return mtsTimeDiff.TotalMilliseconds
        Catch ex As Exception
            Me.SetSubErrInf("Date2Lng", ex)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="LngValue">The number of milliseconds since 1970-1-1</param>
    ''' <param name="IsLocalTime">Convert to local time</param>
    ''' <returns></returns>
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
    Public  Function Lng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            Dim intHourAdd As Integer = 0
            If IsLocalTime = False Then
                Dim oTimeZoneInfo As System.TimeZoneInfo
                oTimeZoneInfo = System.TimeZoneInfo.Local
                intHourAdd = oTimeZoneInfo.GetUtcOffset(Now).Hours
            End If

            Return dteStart.AddSeconds(LngValue + intHourAdd * 3600)
            Me.ClearErr()
        Catch ex As Exception
            Return dteStart
            Me.SetSubErrInf("Lng2Date", ex)
        End Try
    End Function
#Else
    Public Function Lng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            If IsLocalTime = False Then
                Lng2Date = dteStart.AddMilliseconds(LngValue - System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours * 3600000)
            Else
                Lng2Date = dteStart.AddMilliseconds(LngValue)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("Lng2Date", ex)
            Return DateTime.MinValue
        End Try
    End Function
#End If

    Public Function Src2CtlStr(ByRef SrcStr As String) As String
        Try
            If SrcStr.IndexOf(vbCrLf) > 0 Then SrcStr = Replace(SrcStr, vbCrLf, "\r\n")
            If SrcStr.IndexOf(vbCr) > 0 Then SrcStr = Replace(SrcStr, vbCr, "\r")
            If SrcStr.IndexOf(vbTab) > 0 Then SrcStr = Replace(SrcStr, vbTab, "\t")
            If SrcStr.IndexOf(vbBack) > 0 Then SrcStr = Replace(SrcStr, vbBack, "\b")
            If SrcStr.IndexOf(vbFormFeed) > 0 Then SrcStr = Replace(SrcStr, vbFormFeed, "\f")
            If SrcStr.IndexOf(vbVerticalTab) > 0 Then SrcStr = Replace(SrcStr, vbVerticalTab, "\v")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Src2CtlStr", ex)
        End Try
    End Function

    Public Function CtlStr2Src(ByRef CtlStr As String) As String
        Try
            If CtlStr.IndexOf("\r\n") > 0 Then CtlStr = Replace(CtlStr, "\r\n", vbCrLf)
            If CtlStr.IndexOf("\r") > 0 Then CtlStr = Replace(CtlStr, "\r", vbCr)
            If CtlStr.IndexOf("\t") > 0 Then CtlStr = Replace(CtlStr, "\t", vbTab)
            If CtlStr.IndexOf("\b") > 0 Then CtlStr = Replace(CtlStr, "\b", vbBack)
            If CtlStr.IndexOf(vbFormFeed) > 0 Then CtlStr = Replace(CtlStr, "\f", vbFormFeed)
            If CtlStr.IndexOf(vbVerticalTab) > 0 Then CtlStr = Replace(CtlStr, "\v", vbVerticalTab)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CtlStr2Src", ex)
        End Try
    End Function

    Public Function AddMultiLineText(ByRef MainText As String, NewLine As String, Optional LeftTabs As Integer = 0) As String
        Try
            Dim strTabs As String = ""
            If LeftTabs > 0 Then
                For i = 1 To LeftTabs
                    strTabs &= vbTab
                Next
            End If
            MainText &= strTabs & NewLine & vbCrLf
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddMultiLineText", ex)
        End Try
    End Function

    Public Function IsFileExists(FilePath As String) As Boolean
        Try
            Return IO.File.Exists(FilePath)
        Catch ex As Exception
            Me.SetSubErrInf("IsFileExists", ex)
            Return Nothing
        End Try
    End Function

    Public Function IsFolderExists(FolderPath As String) As Boolean
        Try
            Return Directory.Exists(FolderPath)
        Catch ex As Exception
            Me.SetSubErrInf("IsFolderExists", ex)
            Return Nothing
        End Try
    End Function

    Public Function GetEnvVar(EnvVarName As String) As String
        Return GetEnvironmentVariable(EnvVarName)
    End Function

    Public Function GetUserName() As String
        Return Environment.UserName
    End Function
    Public Function CreateFolder(FolderPath As String) As String
        Try
            Directory.CreateDirectory(FolderPath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("CreateFolder", ex)
        End Try
    End Function

    Public Function MoveFile(SrcFilePath As String, DestFilePath As String, Optional IsForceOverride As Boolean = False) As String
        Dim LOG As New PigStepLog("MoveFile")
        Try
            If IsForceOverride = True Then
                If File.Exists(DestFilePath) = True Then
                    LOG.StepName = "Delete(DestFilePath)"
                    File.Delete(DestFilePath)
                End If
            End If
            LOG.StepName = "Move(SrcFilePath, DestFilePath)"
            File.Move(SrcFilePath, DestFilePath)
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(SrcFilePath)
            LOG.AddStepNameInf(DestFilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function


    Public Function DeleteFile(FilePath As String) As String
        Try
            File.Delete(FilePath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFile", ex)
        End Try
    End Function

    Public Function DeleteFolder(FolderPath As String) As String
        Try
            Directory.Delete(FolderPath)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFolder", ex)
        End Try
    End Function

    Public Function DeleteFolder(FolderPath As String, IsSubDir As Boolean) As String
        Try
            Directory.Delete(FolderPath, IsSubDir)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("DeleteFolder", ex)
        End Try
    End Function

    Public Function GetComputerName() As String
        Return Dns.GetHostName()
    End Function

    Public Function IsOsWindows() As Boolean
        Return Me.IsWindows
    End Function

    Public Function MyOsCrLf() As String
        Return Me.OsCrLf
    End Function
    Public Function MyOsPathSep() As String
        Return Me.OsPathSep
    End Function

    Public Function GetFileVersion(FilePath As String, ByRef FileVersion As String) As String
        Dim LOG As New PigStepLog("GetFileVersion")
        Try
            LOG.StepName = "GetVersionInfo"
            Dim oGetVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(FilePath)
            FileVersion = oGetVersionInfo.FileVersion
            oGetVersionInfo = Nothing
            Return "OK"
        Catch ex As Exception
            FileVersion = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileMD5(FilePath As String, ByRef FileMD5 As Date) As String
        Dim LOG As New PigStepLog("GetFileMD5")
        Try
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(FilePath)
            LOG.StepName = "LoadFile"
            LOG.Ret = oPigFile.LoadFile()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            FileMD5 = oPigFile.MD5
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            GetFileMD5 = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFilePigMD5(FilePath As String, ByRef FilePigMD5 As Date) As String
        Dim LOG As New PigStepLog("GetFileCreateTime")
        Try
            LOG.StepName = "New PigFile"
            Dim oPigFile As New PigFile(FilePath)
            LOG.StepName = "LoadFile"
            LOG.Ret = oPigFile.LoadFile()
            If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
            FilePigMD5 = oPigFile.PigMD5
            oPigFile = Nothing
            Return "OK"
        Catch ex As Exception
            FilePigMD5 = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileCreateTime(FilePath As String, ByRef FileCreateTime As Date) As String
        Dim LOG As New PigStepLog("GetFileCreateTime")
        Try
            LOG.StepName = "New FileInfo"
            Dim oFileInfo As New FileInfo(FilePath)
            LOG.StepName = "CreationTime"
            FileCreateTime = oFileInfo.CreationTime
            oFileInfo = Nothing
            Return "OK"
        Catch ex As Exception
            FileCreateTime = #1/1/1900#
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileUpdateTime(FilePath As String, ByRef FileUpdateTime As Date) As String
        Dim LOG As New PigStepLog("GetFileUpdateTime")
        Try
            LOG.StepName = "New FileInfo"
            Dim oFileInfo As New FileInfo(FilePath)
            LOG.StepName = "LastWriteTime"
            FileUpdateTime = oFileInfo.LastWriteTime
            oFileInfo = Nothing
            Return "OK"
        Catch ex As Exception
            FileUpdateTime = #1/1/1900#
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function GetFileText(FilePath As String, ByRef FileText As String) As String
        Dim LOG As New PigStepLog("GetFileText")
        Try
            LOG.StepName = "New StreamReader"
            Dim srMain As New StreamReader(FilePath)
            LOG.StepName = "ReadToEnd"
            FileText = srMain.ReadToEnd()
            LOG.StepName = "Close"
            srMain.Close()
            Return "OK"
        Catch ex As Exception
            FileText = ""
            LOG.AddStepNameInf(FilePath)
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function ASyncSaveTextToFile(FilePath As String, SaveText As String, ByRef OutThreadID As Integer) As String
        Return Me.mASyncSaveTextToFile(FilePath, SaveText, OutThreadID)
    End Function


    Private Function mASyncSaveTextToFile(FilePath As String, SaveText As String, ByRef OutThreadID As Integer) As String
        Try
            Dim struMain As mStruSaveTextToFile
            With struMain
                .FilePath = FilePath
                .SaveText = SaveText
                .IsAsync = True
            End With
            Dim oThread As New Thread(AddressOf mSaveTextToFile)
            oThread.Start(struMain)
            OutThreadID = oThread.ManagedThreadId
            oThread = Nothing
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mASyncSaveTextToFile", ex)
        End Try
    End Function

    Public Function SaveTextToFile(FilePath As String, SaveText As String) As String
        Dim struMain As mStruSaveTextToFile
        With struMain
            .FilePath = FilePath
            .SaveText = SaveText
        End With
        Return Me.mSaveTextToFile(struMain)
    End Function


    Private Structure mStruSaveTextToFile
        Public FilePath As String
        Public SaveText As String
        Public IsAsync As Boolean
    End Structure

    Private Function mSaveTextToFile(StruMain As mStruSaveTextToFile) As String
        Dim LOG As New PigStepLog("mSaveTextToFile")
        Dim sarMain As StruASyncRet
        Try
            If StruMain.IsAsync = True Then
                sarMain.BeginTime = Now
            End If
            LOG.StepName = "New StreamWriter"
            Dim swMain As New StreamWriter(StruMain.FilePath, False)
            LOG.StepName = "Write"
            swMain.Write(StruMain.SaveText)
            LOG.StepName = "Close"
            swMain.Close()
            If StruMain.IsAsync = True Then
                sarMain.EndTime = Now
                sarMain.ThreadID = Thread.CurrentThread.ManagedThreadId
                sarMain.Ret = "OK"
                RaiseEvent ASyncRet_SaveTextToFile(sarMain)
            End If
            Return "OK"
        Catch ex As Exception
            LOG.AddStepNameInf(StruMain.FilePath)
            If StruMain.IsAsync = True Then
                sarMain.EndTime = Now
                sarMain.ThreadID = Thread.CurrentThread.ManagedThreadId
                sarMain.Ret = Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
                RaiseEvent ASyncRet_SaveTextToFile(sarMain)
            End If
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    ''' <summary>
    ''' 将一行PLSQL Developer 导出的CSV文本行转换成标准的SQL Server BCP格式行|Convert a line of CSV text exported by PL/SQL Developer into a standard SQL Server BCP format line
    ''' </summary>
    ''' <param name="CsvLine"></param>
    ''' <param name="BcpLine"></param>
    ''' <returns></returns>
    Public Function PLSqlCsv2Bcp(CsvLine As String, ByRef BcpLine As String) As String
        Try
            CsvLine &= ","
            Do While True
                If InStr(CsvLine, """"",") = 0 Then Exit Do
                CsvLine = Replace(CsvLine, """"",", """ "",")
            Loop
            BcpLine = ""
            Dim bolIsBegin As Boolean = False
            Do While True
                Dim strCol As String = Me.GetStr(CsvLine, """", """,")
                If strCol = "" Then Exit Do
                If strCol = " " Then strCol = ""
                If bolIsBegin = True Then
                    BcpLine &= vbTab
                Else
                    bolIsBegin = True
                End If
                BcpLine &= strCol
            Loop
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("PLSqlCsv2Bcp", ex)
        End Try
    End Function

    Public Function GetHumanSize(SrcSize As Decimal) As String
        Try
            Dim strSize As String = SrcSize.ToString
            If SrcSize < 0 Then Throw New Exception("Invalid size")
            If SrcSize < 1024 Then
                strSize = SrcSize.ToString & " bytes"
            Else
                SrcSize = SrcSize / 1024
                If SrcSize < 1024 Then
                    strSize = Math.Round(SrcSize, 2) & " KB"
                Else
                    SrcSize = SrcSize / 1024
                    If SrcSize < 1024 Then
                        strSize = Math.Round(SrcSize, 2) & " MB"
                    Else
                        SrcSize = SrcSize / 1024
                        If SrcSize < 1024 Then
                            strSize = Math.Round(SrcSize, 2) & " GB"
                        Else
                            SrcSize = SrcSize / 1024
                            If SrcSize < 1024 Then
                                strSize = Math.Round(SrcSize, 2) & " TB"
                            Else
                                SrcSize = SrcSize / 1024
                                If SrcSize < 1024 Then
                                    strSize = Math.Round(SrcSize, 2) & " PB"
                                Else
                                    SrcSize = SrcSize / 1024
                                    strSize = Math.Round(SrcSize, 2) & " EB"
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            Return strSize
        Catch ex As Exception
            Return ex.Message.ToString
        End Try
    End Function

    ''' <summary>
    ''' 将数字转换成中文显示|Convert numbers to Chinese display
    ''' </summary>
    ''' <param name="Digital">数字</param>
    ''' <param name="IsDecimal">是否小数，不是小数部分只支持千百十个位</param>
    ''' <param name="IsUCase">是否大写数字</param>
    ''' <returns></returns>
    Public Function DigitalToChnName(Digital As String, IsDecimal As Boolean, Optional IsUCase As Boolean = False) As String
        Try
            Dim i As Integer
            Dim lngDigital As Integer
            Dim lngDigitalLen As Integer
            Dim strDigital(0 To 9) As String
            Dim strUnit(0 To 3) As String
            If IsDecimal = False Then
                Select Case Len(Digital)
                    Case Is < 4
                        Digital = Right("0000" & Digital, 4)
                    Case Is > 4
                        Digital = Right(Digital, 4)
                End Select
            End If
            lngDigitalLen = Len(Digital)

            strDigital(0) = "零"
            If IsUCase = True Then
                strDigital(1) = "壹"
                strDigital(2) = "贰"
                strDigital(3) = "叁"
                strDigital(4) = "肆"
                strDigital(5) = "伍"
                strDigital(6) = "陆"
                strDigital(7) = "柒"
                strDigital(8) = "捌"
                strDigital(9) = "玖"

                strUnit(0) = "仟"
                strUnit(1) = "佰"
                strUnit(2) = "拾"
                strUnit(3) = ""
            Else
                strDigital(1) = "一"
                strDigital(2) = "二"
                strDigital(3) = "三"
                strDigital(4) = "四"
                strDigital(5) = "五"
                strDigital(6) = "六"
                strDigital(7) = "七"
                strDigital(8) = "八"
                strDigital(9) = "九"

                strUnit(0) = "千"
                strUnit(1) = "百"
                strUnit(2) = "十"
                strUnit(3) = ""
            End If
            DigitalToChnName = ""
            For i = 0 To lngDigitalLen - 1
                lngDigital = CLng(Mid(Digital, i + 1, 1))
                If IsDecimal = False Then
                    '这种情况lngDigitalLen为4
                    If lngDigital = 0 Then
                        DigitalToChnName &= strDigital(lngDigital)
                    Else
                        DigitalToChnName &= strDigital(lngDigital) & strUnit(i)
                    End If
                Else
                    DigitalToChnName &= strDigital(lngDigital)
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("DigitalToChnName", ex)
            Return ""
        End Try

    End Function

    ''' <summary>
    ''' 转换HTML标记
    ''' </summary>
    ''' <param name="HtmlStr">源HTML字符串|Source HTML string</param>
    ''' <param name="HowToConvHtml">如何转换|How to convert</param>
    ''' <returns></returns>
    Public Function ConvertHtmlStr(SrcHtml As String, HowToConvHtml As EmnHowToConvHtml) As String

        Try
            ConvertHtmlStr = SrcHtml
            Select Case HowToConvHtml
                Case EmnHowToConvHtml.DisableHTML
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "<", "&lt;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, ">", "&gt;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, " ", "&nbsp;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, vbTab, "&nbsp;&nbsp;&nbsp;&nbsp;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, vbCrLf, "<br>")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;br&gt;", "<br>")
                Case EmnHowToConvHtml.EnableHTML
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;", "<")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&gt;", ">")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&nbsp;&nbsp;&nbsp;&nbsp;", vbTab) '必须在空格转换前面
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&nbsp;", " ")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;br&gt;", "<br>")
                Case EmnHowToConvHtml.DisableHTMLOnlySymbol
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "<", "&lt;")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, ">", "&gt;")
                Case EmnHowToConvHtml.EnableHTMLOnlySymbol
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&lt;", "<")
                    ConvertHtmlStr = Replace(ConvertHtmlStr, "&gt;", ">")
                Case Else
                    Throw New Exception("Invalid HowToConvHtml " & HowToConvHtml.ToString)
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("ConvertHtmlStr", ex)
            Return ""
        End Try
    End Function


    Public Function GetUUID() As String
        Dim LOG As New PigStepLog("GetUUID")
        Try
            If Me.IsWindows = True Then
                GetUUID = ""
            Else
                Dim strFilePath As String = "/proc/sys/kernel/random/uuid"
                LOG.StepName = "New PigFile"
                Dim oPigFile As New PigFile(strFilePath)
                If oPigFile.LastErr <> "" Then
                    LOG.AddStepNameInf(strFilePath)
                    Throw New Exception(oPigFile.LastErr)
                End If
                LOG.StepName = "oPigFile.LoadFile"
                LOG.Ret = oPigFile.LoadFile()
                If LOG.Ret <> "OK" Then
                    LOG.AddStepNameInf(strFilePath)
                    Throw New Exception(LOG.Ret)
                End If
                LOG.StepName = "New PigText"
                Dim oPigText As New PigText(oPigFile.GbMain.Main, PigText.enmTextType.Unicode)
                If oPigText.LastErr <> "" Then
                    If oPigFile.GbMain Is Nothing Then
                        LOG.AddStepNameInf("oPigFile.GbMain Is Nothing")
                    ElseIf oPigFile.GbMain.Main Is Nothing Then
                        LOG.AddStepNameInf("oPigFile.GbMain.Main Is Nothing")
                    Else
                        LOG.AddStepNameInf("Main.Length=" & oPigFile.GbMain.Main.Length)
                    End If
                    Throw New Exception(oPigText.LastErr)
                End If
                GetUUID = oPigText.Text
                GetUUID = oPigFile.GbMain.Main.Length
                oPigFile = Nothing
                oPigText = Nothing
            End If
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 获取对齐的字符串|Gets the aligned string
    ''' </summary>
    ''' <param name="SrcStr">源串|Source string</param>
    ''' <param name="Alignment">对齐方式|Alignment</param>
    ''' <param name="RowLen">行长度|Row length</param>
    ''' <returns></returns>
    Public Function GetAlignStr(SrcStr As String, Alignment As EnmAlignment, RowLen As Integer) As String
        Try
            Dim intSrcLen As Integer = Len(SrcStr)
            GetAlignStr = ""
            Select Case Alignment
                Case EnmAlignment.Left
                    If intSrcLen >= RowLen Then
                        GetAlignStr = Left(SrcStr, RowLen)
                    Else
                        GetAlignStr = SrcStr & Me.GetRepeatStr(RowLen - intSrcLen， " ")
                    End If
                Case EnmAlignment.Right
                    If intSrcLen >= RowLen Then
                        GetAlignStr = Right(SrcStr, RowLen)
                    Else
                        GetAlignStr = Me.GetRepeatStr(RowLen - intSrcLen， " ") & SrcStr
                    End If
                Case EnmAlignment.Center
                    Dim intBegin As Integer = (RowLen - intSrcLen) / 2
                    Dim intMidLen As Integer = intSrcLen
                    If intBegin < 1 Then intBegin = 1
                    Select Case intSrcLen
                        Case < RowLen
                            GetAlignStr = Me.GetRepeatStr(intBegin， " ") & SrcStr & Me.GetRepeatStr(RowLen - intBegin - intSrcLen, " ")
                        Case = RowLen
                            GetAlignStr = SrcStr
                        Case > RowLen
                            intBegin = (intSrcLen - RowLen) / 2
                            If intBegin < 1 Then intBegin = 1
                            intMidLen = RowLen
                            GetAlignStr = Mid(SrcStr, intBegin, intMidLen)
                    End Select
                Case Else
                    Throw New Exception("Invalid Alignment " & Alignment.ToString)
            End Select
        Catch ex As Exception
            Me.SetSubErrInf("GetAlignStr", ex)
            Return ""
        End Try
    End Function

    Public Function GetRepeatStr(Number As Integer, SrcStr As String) As String
        Try
            GetRepeatStr = ""
            For i = 1 To Number
                GetRepeatStr &= SrcStr
            Next
        Catch ex As Exception
            Me.SetSubErrInf("GetRepeatStr", ex)
            Return ""
        End Try
    End Function

End Class
