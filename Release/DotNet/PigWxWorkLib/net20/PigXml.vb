'**********************************
'* Name: PigXml
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Processing XML string splicing and parsing. 处理XML字符串拼接及解析
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.10
'* Create Time: 8/11/2019
'1.0.2  2019-11-10  修改bug
'1.0.3  2020-5-26  修改bug
'1.0.5  2020-6-5  增加 CData处理和流处理
'1.0.6  2020-6-7  增加 CData处理和流处理优化
'1.0.8  2020-6-19  修改 FillByXmlReader
'1.0.9  2020-6-20  修改 Bug
'1.0.10 2020-7-6    增加 XmlGetInt
'1.0.11 2/2/2021   Modify mLng2Date
'1.0.12 24/8/2021   Modify mLng2Date
'1.0.13 24/8/2021   Modify mLng2Date for NETCOREAPP3_1_OR_GREATER
'1.1 24/8/2021   Modify mGetStr
'1.2 22/12/2021   Modify mXMLAddStr
'1.3 3/1/2022   Modify Err.Raise to Throw New Exception
'1.4 29/2/2022   Remove FillByXmlReader, add XmlDocument,GetXmlDocText
'1.5 30/5/2022   Add mInitXmlDocument,InitXmlDocument,mGetXmlDoc, modify GetXmlDocText,mGetXmlDoc
'1.6 31/5/2022   Add XmlDocGetInt
'1.7 3/6/2022   Modify xpXMLAddWhere,AddEleLeftSign,mXMLAddStr, add AddEleLeftAttribute
'1.8 6/6/2022   Add XmlDocGetStr
'1.9 15/6/2022  Modify XmlDocGetStr
'1.10 11/7/2022  Add AddXmlFragment,mSrc2CtlStr,mCtlStr2Src,mUnEscapeXmlValue,mEscapeXmlValue, modify mXMLAddStr,mGetXmlDoc
'*******************************************************

Imports System.Xml
Public Class PigXml
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.10.8"
    Private mstrMainXml As String
    Public Property XmlDocument As XmlDocument
    Private ReadOnly Property mPigFunc As New PigFunc

    'Private mslMain As SortedList

    ''' <summary>增加一个XML标记的部分</summary>
    Private Enum xpXMLAddWhere
        ''' <summary>左标记</summary>
        Left = 0
        ''' <summary>右标记</summary>
        Right = 1
        ''' <summary>左右标记</summary>
        Both = 2
        ''' <summary>只有值</summary>
        ValueOnly = 3
        ''' <summary>左标记开始</summary>
        LeftBegin = 4
        ''' <summary>左标记属性</summary>
        LeftAttribute = 5
        ''' <summary>左标记结束</summary>
        LeftEnd = 5
    End Enum

    ''' <summary>增加一个XML标记的部分</summary>
    Private Enum EmnGetXmlDoc
        Text = 0
        Attribute = 1
        Node = 2
    End Enum

    Private msbMain As System.Text.StringBuilder    '主体的XML
    ''' <summary>
    ''' 生成的XML是否加入回车符|Whether carriage return is added to the generated XML
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property IsCrLf As Boolean
    ''' <summary>
    ''' 生成的XML的值是否转义控制字符|Whether the value of the generated XML escapes the control character
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property IsChgCtrlChar As Boolean


    Public Sub New(IsCrLf As Boolean, Optional IsChgCtrlChar As Boolean = True)
        MyBase.New(CLS_VERSION)
        Me.IsCrLf = IsCrLf
        Me.IsChgCtrlChar = IsChgCtrlChar
        Me.Clear()
    End Sub

    Public Sub SetMainXml(InXml As String)
        mstrMainXml = InXml
    End Sub

    '''' <summary>通过XmlReader填充元素</summary>
    '''' <param name="oXmlReader">XmlReader对象</param>
    '''' <param name="DateKeys">日期型的键值，有则转为日期型</param>
    'Public Function FillByXmlReader(ByRef oXmlReader As XmlReader, Optional DateKeys As String = "") As String
    '    Dim strStepName As String = ""
    '    Try
    '        strStepName = "mslMain.Clear"
    '        mslMain.Clear()
    '        strStepName = "读取XML流"
    '        While oXmlReader.Read
    '            Select Case oXmlReader.NodeType
    '                Case XmlNodeType.Element
    '                    Dim strEleName As String = oXmlReader.Name
    '                    Dim intEleIdx As Integer = mslMain.IndexOfKey(strEleName)
    '                    strStepName = strEleName & ".Read"
    '                    oXmlReader.Read() '读取内容
    '                    Dim strValue As String = oXmlReader.Value
    '                    Select Case oXmlReader.NodeType
    '                        Case XmlNodeType.CDATA
    '                            If intEleIdx = -1 Then
    '                                mslMain.Add(strEleName, strValue)
    '                            End If
    '                        Case Else
    '                            If InStr(strValue, ".") > 0 Then
    '                                Dim decValue As Decimal = CDec(strValue)
    '                                If intEleIdx = -1 Then
    '                                    mslMain.Add(strEleName, decValue)
    '                                End If
    '                            Else
    '                                Dim lngValue As Long = CLng(strValue)
    '                                If InStr(DateKeys, "<" & strEleName & ">") > 0 Then
    '                                    Dim dteValue As DateTime = Me.mLng2Date(lngValue)
    '                                    If intEleIdx = -1 Then
    '                                        mslMain.Add(strEleName, dteValue)
    '                                    End If
    '                                Else
    '                                    If intEleIdx = -1 Then
    '                                        mslMain.Add(strEleName, lngValue)
    '                                    End If
    '                                End If
    '                            End If
    '                    End Select
    '            End Select
    '        End While
    '        Return "OK"
    '    Catch ex As Exception
    '        Return Me.GetSubErrInf("FillByXmlReader", strStepName, ex)
    '    End Try
    'End Function


    ''' <summary>获取一个元素字符串值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetStr(XMLSign As String) As String
        Return Me.mXmlGetStr(mstrMainXml, XMLSign, False)
    End Function

    ''' <summary>获取一个元素长整值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetLong(XMLSign As String) As Long
        Try
            XmlGetLong = CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetLong", ex)
            Return 0
        End Try
    End Function

    ''' <summary>获取一个元素整值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetInt(XMLSign As String) As Integer
        Try
            XmlGetInt = CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetLong", ex)
            Return 0
        End Try
    End Function

    ''' <summary>获取一个元素浮点值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetDec(XMLSign As String) As Decimal
        Try
            XmlGetDec = CDec(Me.mXmlGetStr(mstrMainXml, XMLSign, False))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetDec", ex)
            Return 0
        End Try
    End Function

    ''' <summary>整型转日期型</summary>
    ''' <param name="LngValue">整型值，即1970-1-1以来的秒数</param>
    ''' <param name="IsLocalTime">是否本地时间，否则为格林威治时间</param>
#If NET40_OR_GREATER Or NETCOREAPP3_1_OR_GREATER Then
    Private Function mLng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
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
            Me.SetSubErrInf("mLng2Date", ex)
        End Try
    End Function
#Else
    Private Function mLng2Date(LngValue As Long, Optional IsLocalTime As Boolean = True) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            If IsLocalTime = False Then
                mLng2Date = dteStart.AddMilliseconds(LngValue - System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours * 3600000)
            Else
                mLng2Date = dteStart.AddMilliseconds(LngValue)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("mLng2Date", ex)
            Return DateTime.MinValue
        End Try
    End Function
#End If



    ''' <summary>获取一个元素日期值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetDate(XMLSign As String) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            XmlGetDate = Me.mLng2Date(CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False)))
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetDate", ex)
            Return dteStart
        End Try
    End Function

    ''' <summary>获取一个元素日期值</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetDate(XMLSign As String, IsLocalTime As Boolean) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            XmlGetDate = Me.mLng2Date(CLng(Me.mXmlGetStr(mstrMainXml, XMLSign, False)), IsLocalTime)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("XmlGetDate", ex)
            Return dteStart
        End Try
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="SrcXmlStr">源XML串，读取后会截去</param>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="IsCData">是否CDATA</param>
    Private Overloads Function mXmlGetStr(ByRef SrcXmlStr As String, XMLSign As String, IsCData As Boolean) As String
        Try
            Dim strBegin As String, strEnd As String
            strBegin = "<" & XMLSign & ">"
            strEnd = "</" & XMLSign & ">"
            If IsCData = True Then
                strBegin &= "<![CDATA["
                strEnd = "]]>" & strEnd
            End If
            mXmlGetStr = Me.mGetStr(SrcXmlStr, strBegin, strEnd, True)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mXmlGetStr", ex)
            Return ""
        End Try
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="IsCData">是否CDATA</param>
    Public Overloads Function XmlGetStr(XMLSign As String, IsCData As Boolean) As String
        Return Me.mXmlGetStr(mstrMainXml, XMLSign, IsCData)
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="SrcXmlStr">XML源串</param>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function XmlGetStr(ByRef SrcXmlStr As String, XMLSign As String) As String
        Return Me.mXmlGetStr(SrcXmlStr, XMLSign, False)
    End Function

    ''' <summary>获取一个元素的字符串值</summary>
    ''' <param name="SrcXmlStr">XML源串</param>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="IsCData">是否CDATA</param>
    Public Overloads Function XmlGetStr(ByRef SrcXmlStr As String, XMLSign As String, IsCData As Boolean) As String
        Return Me.mXmlGetStr(SrcXmlStr, XMLSign, IsCData)
    End Function



    ''' <remarks>截取字符串</remarks>
    Private Function mGetStr(ByRef SrcStr As String, BeginKey As String, EndKey As String, Optional IsCut As Boolean = True) As String
        Dim intBegin As Integer
        Dim intEnd As Integer
        Dim intBeginLen As Integer
        Dim intEndLen As Integer
        Try
            intBeginLen = Len(BeginKey)
            intBegin = InStr(SrcStr, BeginKey)
            intEndLen = Len(EndKey)
            If intEndLen = 0 Then
                intEnd = Len(SrcStr) + 1
            Else
                intEnd = InStr(intBegin + intBeginLen + 1, SrcStr, EndKey)
                If intEnd = 0 Then Throw New Exception("intEnd is 0")
            End If
            If intEnd <= intBegin Then Throw New Exception("intEnd <= intBegin")
            If intBegin = 0 Then Throw New Exception("intBegin is 0")
            mGetStr = Mid(SrcStr, intBegin + intBeginLen, (intEnd - intBegin - intBeginLen))
            If IsCut = True Then
                SrcStr = Left(SrcStr, intBegin - 1) & Mid(SrcStr, intEnd + intEndLen)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mGetStr", ex)
            Return ""
        End Try
    End Function


    ''' <summary>增加一个完整元素</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, Me.IsCrLf)
            If AddEle <> "OK" Then Throw New Exception(AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个完整元素</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    ''' <param name="IsCData">是否CDATA</param>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String, IsCData As Boolean) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, Me.IsCrLf, , IsCData)
            If AddEle <> "OK" Then Throw New Exception(AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个完整元素带缩进</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    ''' <remarks></remarks>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String, LeftTab As Integer) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, Me.IsCrLf, LeftTab)
            If AddEle <> "OK" Then Throw New Exception(AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function

    ''' <summary>增加一个完整元素带缩进</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="XMLValue">元素值</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    ''' <param name="IsCData">是否CDATA</param>
    ''' <remarks></remarks>
    Public Overloads Function AddEle(XMLSign As String, XMLValue As String, LeftTab As Integer, IsCData As Boolean) As String
        Try
            AddEle = Me.mXMLAddStr(msbMain, XMLSign, XMLValue, xpXMLAddWhere.Both, Me.IsCrLf, LeftTab, IsCData)
            If AddEle <> "OK" Then Throw New Exception(AddEle)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEle", ex)
        End Try
    End Function


    ''' <summary>
    ''' 增加一个XML值|Add an XML value
    ''' </summary>
    ''' <param name="XMLValue"></param>
    ''' <param name="IsCData">是否使用<![CDATA[]]>是则不对内容转义|Whether to use <![CDATA[]]> if yes, the content will not be escaped</param>
    ''' <returns></returns>
    Public Function AddEleValue(XMLValue As String, Optional IsCData As Boolean = False) As String
        Try
            If Me.IsChgCtrlChar = True Then Me.mSrc2CtlStr(XMLValue)
            If IsCData = True Then
                Me.msbMain.Append("<![CDATA[" & XMLValue & "]]>")
            Else
                Dim strRet As String = Me.mEscapeXmlValue(XMLValue)
                If strRet <> "OK" Then Throw New Exception(strRet)
                Me.msbMain.Append(XMLValue)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleValue", ex)
        End Try
    End Function


    ''' <summary>
    ''' 增加一个XML片段，对片段的内容不作转义和控制字符转换处理。|Add an XML fragment, and do not escape the contents of the fragment and control character conversion.
    ''' </summary>
    ''' <param name="XmlFragment"></param>
    ''' <returns></returns>
    Public Function AddXmlFragment(XmlFragment As String) As String
        Try
            Me.msbMain.Append(XmlFragment)
            If Me.IsCrLf = True Then Me.msbMain.Append(Me.OsCrLf)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddXmlFragment", ex)
        End Try
    End Function

    Private Sub mSrc2CtlStr(ByRef SrcStr As String)
        Try
            If SrcStr.IndexOf(vbCrLf) > 0 Then SrcStr = Replace(SrcStr, vbCrLf, "\r\n")
            If SrcStr.IndexOf(vbCr) > 0 Then SrcStr = Replace(SrcStr, vbCr, "\r")
            If SrcStr.IndexOf(vbTab) > 0 Then SrcStr = Replace(SrcStr, vbTab, "\t")
            If SrcStr.IndexOf(vbBack) > 0 Then SrcStr = Replace(SrcStr, vbBack, "\b")
            If SrcStr.IndexOf(vbFormFeed) > 0 Then SrcStr = Replace(SrcStr, vbFormFeed, "\f")
            If SrcStr.IndexOf(vbVerticalTab) > 0 Then SrcStr = Replace(SrcStr, vbVerticalTab, "\v")
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mSrc2CtlStr", ex)
        End Try
    End Sub


    ''' <summary>
    ''' 转义XML值|Escape XML value
    ''' </summary>
    ''' <param name="InXmlStr"></param>
    ''' <returns></returns>
    Private Function mEscapeXmlValue(ByRef InXmlStr As String) As String
        Try
            If InStr(InXmlStr, "<") > 0 Then InXmlStr = Replace(InXmlStr, "<", "&lt;")
            If InStr(InXmlStr, ">") > 0 Then InXmlStr = Replace(InXmlStr, ">", "&gt;")
            If InStr(InXmlStr, "&") > 0 Then InXmlStr = Replace(InXmlStr, "&", "&apos;")
            '            If InStr(InXmlStr, "'") > 0 Then InXmlStr = Replace(InXmlStr, "'", "&lt;")
            If InStr(InXmlStr, """") > 0 Then InXmlStr = Replace(InXmlStr, """", "&quot")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mEscapeXmlValue", ex)
        End Try
    End Function

    Private Function mUnEscapeXmlValue(ByRef InXmlStr As String) As String
        Try
            If InStr(InXmlStr, "&lt;") > 0 Then InXmlStr = Replace(InXmlStr, "&lt;", "<")
            If InStr(InXmlStr, "&gt;") > 0 Then InXmlStr = Replace(InXmlStr, "&gt;", ">")
            If InStr(InXmlStr, "&apos;") > 0 Then InXmlStr = Replace(InXmlStr, "&apos;", "&")
            'If InStr(InXmlStr, "&lt;") > 0 Then InXmlStr = Replace(InXmlStr, "&lt;", "'")
            'If InStr(InXmlStr, "&quot") > 0 Then InXmlStr = Replace(InXmlStr, "&quot", """")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mUnEscapeXmlValue", ex)
        End Try
    End Function

    ''' <summary>增加一个XML左标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Function AddEleLeftSign(XMLSign As String, Optional IsBegin As Boolean = False) As String
        Try
            If IsBegin = True Then
                AddEleLeftSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.LeftBegin, False)
            Else
                AddEleLeftSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Left, Me.IsCrLf)
            End If
            If AddEleLeftSign <> "OK" Then Throw New Exception(AddEleLeftSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleLeftSign", ex)
        End Try
    End Function

    Public Function AddEleLeftAttribute(AttributeName As String, AttributeValue As String) As String
        Try
            AddEleLeftAttribute = Me.mXMLAddStr(msbMain, "", "", xpXMLAddWhere.LeftAttribute, False,,, AttributeName, AttributeValue)
            If AddEleLeftAttribute <> "OK" Then Throw New Exception(AddEleLeftAttribute)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleLeftAttribute", ex)
        End Try
    End Function

    Public Function AddEleLeftSignEnd() As String
        Try
            msbMain.Append(">")
            If Me.IsCrLf = True Then msbMain.Append(Me.OsCrLf)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleLeftSignEnd", ex)
        End Try
    End Function


    ''' <summary>增加一个XML左标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    Public Function AddEleLeftSign(XMLSign As String, LeftTab As Integer, Optional IsBegin As Boolean = False) As String
        Try
            If IsBegin = True Then
                AddEleLeftSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.LeftBegin, False, LeftTab)
            Else
                AddEleLeftSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Left, Me.IsCrLf, LeftTab)
            End If
            If AddEleLeftSign <> "OK" Then Throw New Exception(AddEleLeftSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleLeftSign", ex)
        End Try
    End Function

    ''' <summary>增加一个XML右标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    Public Overloads Function AddEleRightSign(XMLSign As String) As String
        Try
            AddEleRightSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Right, Me.IsCrLf)
            If AddEleRightSign <> "OK" Then Throw New Exception(AddEleRightSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleRightSign", ex)
        End Try
    End Function

    ''' <summary>增加一个XML右标记</summary>
    ''' <param name="XMLSign">元素标记</param>
    ''' <param name="LeftTab">前面缩进的TAB数</param>
    Public Overloads Function AddEleRightSign(XMLSign As String, LeftTab As Integer) As String
        Try
            AddEleRightSign = Me.mXMLAddStr(msbMain, XMLSign, "", xpXMLAddWhere.Right, Me.IsCrLf, LeftTab)
            If AddEleRightSign <> "OK" Then Throw New Exception(AddEleRightSign)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("AddEleRightSign", ex)
        End Try
    End Function

    Private Function mXMLAddStr(ByRef sbAny As System.Text.StringBuilder, ByRef XMLSign As String, ByRef XMLValue As String,
                         Optional ByVal AddWhere As xpXMLAddWhere = xpXMLAddWhere.Both,
                         Optional ByVal IsCrlf As Boolean = False,
                         Optional ByVal LeftTab As Integer = 0,
                         Optional ByVal IsCData As Boolean = False,
                         Optional ByVal AttributeName As String = "",
                         Optional ByVal AttributeValue As String = ""
                                ) As String
        Try
            Select Case AddWhere
                Case xpXMLAddWhere.Left, xpXMLAddWhere.Both, xpXMLAddWhere.Right
10:                 If LeftTab > 0 Then
                        For i = 1 To LeftTab
                            sbAny.Append(vbTab)
                        Next
                    End If
            End Select
            Select Case AddWhere
                Case xpXMLAddWhere.LeftAttribute
                    sbAny.Append(" ")
                    sbAny.Append(AttributeName)
                    sbAny.Append("=""" & AttributeValue & """")
                Case xpXMLAddWhere.LeftEnd
                    sbAny.Append(">")
                Case xpXMLAddWhere.LeftBegin
                    sbAny.Append("<")
                    sbAny.Append(XMLSign)
                Case xpXMLAddWhere.Left
                    sbAny.Append("<")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                    If IsCData = True Then
                        sbAny.Append("<![CDATA[")
                    End If
                Case xpXMLAddWhere.Right
                    If IsCData = True Then
                        sbAny.Append("]]>")
                    End If
30:                 sbAny.Append("</")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                Case xpXMLAddWhere.Both
40:                 sbAny.Append("<")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                    Me.AddEleValue(XMLValue, IsCData)
                    sbAny.Append("</")
                    sbAny.Append(XMLSign)
                    sbAny.Append(">")
                Case xpXMLAddWhere.ValueOnly
                    sbAny.Append(XMLValue)
            End Select
50:         If IsCrlf = True Then
                If Me.IsWindows = True Then
                    sbAny.Append(vbCrLf)
                Else
                    sbAny.Append(vbLf)
                End If
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mXMLAddStr", ex)
        End Try
    End Function

    Public Function Clear() As String
        '初始化内部XML
        Try
            msbMain = Nothing
            msbMain = New System.Text.StringBuilder("")
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function


    Public ReadOnly Property MainXmlStr As String
        Get
            If msbMain Is Nothing Then
                Return mstrMainXml
            Else
                Return msbMain.ToString
            End If
        End Get
    End Property

    ''' <summary>
    ''' 快速获取一个元素的XML结点|Quickly get the XmlNode of an element
    ''' </summary>
    ''' <param name="XmlKey">键值，例如：元素1.元素2.[...元素N]|Key value, for example: element 1 Element 2 [... Element n]</param>
    ''' <returns></returns>
    Public Function GetXmlDocNode(XmlKey As String) As XmlNode
        Try
            GetXmlDocNode = Nothing
            Dim strRet As String = Me.mGetXmlDoc(EmnGetXmlDoc.Node, XmlKey,, GetXmlDocNode)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("GetXmlDocNode", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 快速获取一个元素的XML结点|Quickly get the XmlNode of an element
    ''' </summary>
    ''' <param name="XmlKey">键值，例如：元素1.元素2.[...元素N]|Key value, for example: element 1 Element 2 [... Element n]</param>
    ''' <param name="SkipTimes">最后一个元素，如有名称重复，此为跳过的次数|If the name of the last element is repeated, this is the number of times to skip</param>
    ''' <returns></returns>
    Public Function GetXmlDocNode(XmlKey As String, SkipTimes As Integer) As XmlNode
        Try
            GetXmlDocNode = Nothing
            Dim strRet As String = Me.mGetXmlDoc(EmnGetXmlDoc.Node, XmlKey, , GetXmlDocNode,, SkipTimes)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("GetXmlDocNode", ex)
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' 快速获取一个元素的值|Get the value of an element quickly
    ''' </summary>
    ''' <param name="XmlKey">键值，例如：元素1.元素2.[...元素N]|Key value, for example: element 1 Element 2 [... Element n]</param>
    ''' <returns></returns>
    Public Function GetXmlDocText(XmlKey As String) As String
        Try
            GetXmlDocText = ""
            Dim strRet As String = Me.mGetXmlDoc(EmnGetXmlDoc.Text, XmlKey,,, GetXmlDocText)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("GetXmlDocText", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 快速获取一个元素的值|Get the value of an element quickly
    ''' </summary>
    ''' <param name="XmlKey">键值，例如：元素1.元素2.[...元素N]|Key value, for example: element 1 Element 2 [... Element n]</param>
    ''' <param name="SkipTimes">最后一个元素，如有名称重复，此为跳过的次数|If the name of the last element is repeated, this is the number of times to skip</param>
    ''' <returns></returns>
    Public Function GetXmlDocText(XmlKey As String, SkipTimes As Integer) As String
        Try
            GetXmlDocText = ""
            Dim strRet As String = Me.mGetXmlDoc(EmnGetXmlDoc.Text, XmlKey,,, GetXmlDocText, SkipTimes)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("GetXmlDocText", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 快速获取一个元素的属性|Get the attributes of an element quickly
    ''' </summary>
    ''' <param name="XmlKey">键值，例如：元素1.元素2.[...元素N].属性|Key value, for example: element 1 Element 2 [... Element n].attribute</param>
    ''' <returns></returns>
    Public Function GetXmlDocAttribute(XmlKey As String) As String
        Try
            GetXmlDocAttribute = ""
            Dim strRet As String = Me.mGetXmlDoc(EmnGetXmlDoc.Attribute, XmlKey,,, GetXmlDocAttribute)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("GetXmlDocAttribute", ex)
            Return ""
        End Try
    End Function

    ''' <summary>
    ''' 快速获取一个元素的属性|Get the attributes of an element quickly
    ''' </summary>
    ''' <param name="XmlKey">键值，例如：元素1.元素2.[...元素N].属性|Key value, for example: element 1 Element 2 [... Element n].attribute</param>
    ''' <param name="SkipTimes">最后一个元素，如有名称重复，此为跳过的次数|If the name of the last element is repeated, this is the number of times to skip</param>
    ''' <returns></returns>
    Public Function GetXmlDocAttribute(XmlKey As String, SkipTimes As Integer) As String
        Try
            GetXmlDocAttribute = ""
            Dim strRet As String = Me.mGetXmlDoc(EmnGetXmlDoc.Attribute, XmlKey,,, GetXmlDocAttribute, SkipTimes)
            If strRet <> "OK" Then Throw New Exception(strRet)
        Catch ex As Exception
            Me.SetSubErrInf("GetXmlDocAttribute", ex)
            Return ""
        End Try
    End Function

    Private Function mGetDotCnt(XmlKey As String) As Integer
        Dim intBeginPos As Integer = 1
        Dim intPos As Integer = 0
        mGetDotCnt = 0
        Do While True
            intPos = InStr(intBeginPos, XmlKey, ".")
            If intPos = 0 Then Exit Do
            intBeginPos = intPos + 1
            mGetDotCnt += 1
        Loop
    End Function

    Private Function mGetXmlDoc(WhatGetXmlDoc As EmnGetXmlDoc, XmlKey As String, Optional FromXmlNode As XmlNode = Nothing, Optional ByRef OutNode As XmlNode = Nothing, Optional ByRef OutTextAttribute As String = "", Optional SkipTimes As Integer = 0) As String
        Dim LOG As New PigStepLog("mGetXmlDocText")
        Try
            LOG.StepName = "Check XmlKey"
            Select Case Right(XmlKey, 1)
                Case ".", "#"
                    Throw New Exception("Invalid XmlKey")
            End Select
            OutTextAttribute = ""
            Select Case WhatGetXmlDoc
                Case EmnGetXmlDoc.Text, EmnGetXmlDoc.Node
                    XmlKey &= "."
                Case EmnGetXmlDoc.Attribute
                    XmlKey &= "#"
                Case Else
                    Throw New Exception("Invalid WhatGetXmlDoc is " & WhatGetXmlDoc.ToString)
            End Select
            Dim oParentNode As XmlNode = Nothing
            LOG.StepName = "Set ChildNodes(XmlDocument)"
            Dim oXmlNodeList As XmlNodeList
            If FromXmlNode Is Nothing Then
                oXmlNodeList = Me.XmlDocument.ChildNodes
            Else
                oXmlNodeList = FromXmlNode.ChildNodes
            End If
            Dim bolIsFind As Boolean = False
            Do While True
                Select Case WhatGetXmlDoc
                    Case EmnGetXmlDoc.Text, EmnGetXmlDoc.Node
                        If InStr(XmlKey, ".") = 0 Or XmlKey = "." Then Exit Do
                    Case EmnGetXmlDoc.Attribute
                        If InStr(XmlKey, "#") = 0 Or XmlKey = "#" Then Exit Do
                    Case Else
                        Throw New Exception("Invalid WhatGetXmlDoc is " & WhatGetXmlDoc.ToString)
                End Select
                Dim strNode As String = mPigFunc.GetStr(XmlKey, "", ".")
                If oParentNode IsNot Nothing Then
                    LOG.StepName = "Set ChildNodes(ParentNode)"
                    oXmlNodeList = oParentNode.ChildNodes
                End If
                LOG.StepName = "Find RootNode"
                For i = 0 To oXmlNodeList.Count - 1
                    If oXmlNodeList.Item(i).Name = strNode Then
                        bolIsFind = True
                        Select Case WhatGetXmlDoc
                            Case EmnGetXmlDoc.Text
                                If Me.mGetDotCnt(XmlKey) = 0 Then
                                    For j = i To oXmlNodeList.Count - 1
                                        If oXmlNodeList.Item(j).Name = strNode Then
                                            If SkipTimes <= 0 Then
                                                If oXmlNodeList.Item(j).FirstChild IsNot Nothing Then
                                                    Select Case oXmlNodeList.Item(j).FirstChild.NodeType
                                                        Case XmlNodeType.Text, XmlNodeType.CDATA
                                                            OutTextAttribute = oXmlNodeList.Item(j).InnerText
                                                    End Select
                                                    Exit For
                                                End If
                                            End If
                                            SkipTimes -= 1
                                        End If
                                    Next
                                    Exit Do
                                Else
                                    oParentNode = oXmlNodeList.Item(i)
                                End If
                                'If Me.mGetDotCnt(XmlKey) = 0 Then
                                '    For j = i To oXmlNodeList.Count - 1
                                '        If SkipTimes <= 0 Then
                                '            If oXmlNodeList.Item(j).Name = strNode Then
                                '                If oXmlNodeList.Item(j).FirstChild IsNot Nothing Then
                                '                    Select Case oXmlNodeList.Item(j).FirstChild.NodeType
                                '                        Case XmlNodeType.Text, XmlNodeType.CDATA
                                '                            OutTextAttribute = oXmlNodeList.Item(j).InnerText
                                '                    End Select
                                '                    Exit For
                                '                End If
                                '            End If
                                '        End If
                                '        SkipTimes -= 1
                                '    Next
                                'Else
                                '    oParentNode = oXmlNodeList.Item(i)
                                'End If
                            Case EmnGetXmlDoc.Node
                                If Me.mGetDotCnt(XmlKey) = 0 Then
                                    For j = i To oXmlNodeList.Count - 1
                                        If oXmlNodeList.Item(j).Name = strNode Then
                                            If SkipTimes <= 0 Then
                                                OutNode = oXmlNodeList.Item(j)
                                                Exit For
                                            End If
                                            SkipTimes -= 1
                                        End If
                                    Next
                                    Exit Do
                                Else
                                    oParentNode = oXmlNodeList.Item(i)
                                End If
                            Case EmnGetXmlDoc.Attribute
                                If Me.mGetDotCnt(XmlKey) = 0 Then
                                    For j = i To oXmlNodeList.Count - 1
                                        If oXmlNodeList.Item(j).Name = strNode Then
                                            If SkipTimes <= 0 Then
                                                If InStr(XmlKey, "#") > 0 Then
                                                    Dim strName As String = Me.mPigFunc.GetStr(XmlKey, "", "#")
                                                    If oXmlNodeList.Item(j).Attributes IsNot Nothing Then
                                                        For k = 0 To oXmlNodeList.Item(j).Attributes.Count - 1
                                                            If oXmlNodeList.Item(j).Attributes.Item(k).Name = strName Then
                                                                OutTextAttribute = oXmlNodeList.Item(j).Attributes.Item(k).InnerText
                                                                Exit For
                                                            End If
                                                        Next
                                                    End If
                                                End If
                                            End If
                                            If XmlKey = "" Then Exit For
                                            SkipTimes -= 1
                                        End If
                                    Next
                                    Exit Do
                                Else
                                    oParentNode = oXmlNodeList.Item(i)
                                End If
                            Case Else
                                Throw New Exception("Invalid WhatGetXmlDoc is " & WhatGetXmlDoc.ToString)
                        End Select
                        Exit For
                    End If
                Next
                If bolIsFind = False Then Exit Do
            Loop
            If OutTextAttribute <> "" Then
                LOG.StepName = "mUnEscapeXmlValue"
                LOG.Ret = Me.mUnEscapeXmlValue(OutTextAttribute)
                If LOG.Ret <> "OK" Then Throw New Exception(LOG.Ret)
                If Me.IsChgCtrlChar = True Then Me.mCtlStr2Src(OutTextAttribute)
            End If
            Return "OK"
        Catch ex As Exception
            OutTextAttribute = ""
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function InitXmlDocument() As String
        Return Me.mInitXmlDocument
    End Function

    Public Function InitXmlDocument(XmlFilePath As String) As String
        Return Me.mInitXmlDocument(XmlFilePath)
    End Function

    Private Function mInitXmlDocument(Optional XmlFilePath As String = "") As String
        Dim LOG As New PigStepLog("mInitXmlDocument")
        Try
            LOG.StepName = "Check XmlFilePath"
            If XmlFilePath = "" Then
                If Me.mstrMainXml = "" Then
                    Me.mstrMainXml = Me.MainXmlStr
                    If Me.mstrMainXml = "" Then Throw New Exception("Please call SetMainXml设置 to set MainXmlStr first.")
                End If
                LOG.StepName = "XmlDocument.LoadXml"
                If Me.IsDebug = True Then LOG.AddStepNameInf(Me.mstrMainXml)
                Me.XmlDocument = New XmlDocument
                Me.XmlDocument.LoadXml(Me.mstrMainXml)
            ElseIf Me.mPigFunc.IsFileExists(XmlFilePath) = False Then
                LOG.AddStepNameInf(XmlFilePath)
                Throw New Exception("File not found.")
            Else
                LOG.StepName = "XmlDocument.Load(" & XmlFilePath & ")"
                Me.XmlDocument = New XmlDocument
                Me.XmlDocument.Load(XmlFilePath)
            End If
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function XmlDocGetInt(XmlKey As String, Optional IsAttribute As Boolean = False) As Integer
        Try
            If IsAttribute = True Then
                Return CInt(Me.GetXmlDocAttribute(XmlKey))
            Else
                Return CInt(Me.GetXmlDocText(XmlKey))
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetInt", ex)
            Return 0
        End Try
    End Function

    Public Function XmlDocGetLong(XmlKey As String, Optional IsAttribute As Boolean = False) As Long
        Try
            If IsAttribute = True Then
                Return CLng(Me.GetXmlDocAttribute(XmlKey))
            Else
                Return CLng(Me.GetXmlDocText(XmlKey))
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetLong", ex)
            Return 0
        End Try
    End Function

    Public Function XmlDocGetDec(XmlKey As String, Optional IsAttribute As Boolean = False) As Decimal
        Try
            If IsAttribute = True Then
                Return CDec(Me.GetXmlDocAttribute(XmlKey))
            Else
                Return CDec(Me.GetXmlDocText(XmlKey))
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetDec", ex)
            Return 0
        End Try
    End Function

    Public Function XmlDocGetBool(XmlKey As String, Optional IsAttribute As Boolean = False) As Boolean
        Try
            If IsAttribute = True Then
                Return CBool(Me.GetXmlDocAttribute(XmlKey))
            Else
                Return CBool(Me.GetXmlDocText(XmlKey))
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetBool", ex)
            Return 0
        End Try
    End Function

    Public Function XmlDocGetStr(XmlKey As String, Optional IsAttribute As Boolean = False) As String
        Try
            If IsAttribute = True Then
                Return Me.GetXmlDocAttribute(XmlKey)
            Else
                Return Me.GetXmlDocText(XmlKey)
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetStr", ex)
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' 获取Xml元素的值或属性转换成布尔型（空表示真）|Gets the value or attribute of an XML element and converts it to Boolean (null means true)
    ''' </summary>
    ''' <param name="XmlKey">XML键值|XML key value</param>
    ''' <param name="IsAttribute">是否属性|Attribute or not</param>
    ''' <returns></returns>
    Public Function XmlDocGetBoolEmpTrue(XmlKey As String, Optional IsAttribute As Boolean = False) As Boolean
        Try
            Dim strData As String
            If IsAttribute = True Then
                strData = Me.GetXmlDocAttribute(XmlKey)
                If strData = "" Then
                    Return True
                Else
                    Return CBool(strData)
                End If
            Else
                strData = Me.GetXmlDocText(XmlKey)
                If strData = "" Then
                    Return True
                Else
                    Return CBool(strData)
                End If
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetBoolEmpTrue", ex)
            Return 0
        End Try
    End Function

    Public Function XmlDocGetDate(XmlKey As String, Optional IsAttribute As Boolean = False) As DateTime
        Try
            If IsAttribute = True Then
                Return CDate(Me.GetXmlDocAttribute(XmlKey))
            Else
                Return CDate(Me.GetXmlDocText(XmlKey))
            End If
        Catch ex As Exception
            Me.SetSubErrInf("XmlDocGetDate", ex)
            Return DateTime.MinValue
        End Try
    End Function

    Private Sub mCtlStr2Src(ByRef CtlStr As String)
        Try
            If CtlStr.IndexOf("\r\n") > 0 Then CtlStr = Replace(CtlStr, "\r\n", vbCrLf)
            If CtlStr.IndexOf("\r") > 0 Then CtlStr = Replace(CtlStr, "\r", vbCr)
            If CtlStr.IndexOf("\t") > 0 Then CtlStr = Replace(CtlStr, "\t", vbTab)
            If CtlStr.IndexOf("\b") > 0 Then CtlStr = Replace(CtlStr, "\b", vbBack)
            If CtlStr.IndexOf(vbFormFeed) > 0 Then CtlStr = Replace(CtlStr, "\f", vbFormFeed)
            If CtlStr.IndexOf(vbVerticalTab) > 0 Then CtlStr = Replace(CtlStr, "\v", vbVerticalTab)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mCtlStr2Src", ex)
        End Try
    End Sub

End Class
