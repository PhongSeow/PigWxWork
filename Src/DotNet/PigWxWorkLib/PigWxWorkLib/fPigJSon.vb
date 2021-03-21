'*******************************************************
'* Name: fPigJSon from PigJSon
'* Author: Seow Phong
'* Describe: Simple JSON class, which can assemble and parse JSON definitions without components.
'* Home Url: http://www.seowphong.com
'* Version: 1.0.9
'* Create Time: 8/8/2019
'* 1.0.2    10/8/2020   Code changed from VB6 to VB.NET
'* 1.0.3    12/8/2020   Some Function debugging 
'* 1.0.4    16/8/2020   mAddJSonStr debugging
'* 1.0.5    18/8/2020   AddArrayEle debugging
'* 1.0.6    18/8/2020   Add overloaded function AddArrayEle
'* 1.0.7    18/9/2020   Fix AddEle bug 
'* 1.0.8    19/9/2020   Fix AddArrayEle,mAddEle bug and add AddOneArrayEle
'* 1.0.9    1/10/2020   Fix AddArrayEleValue,add AddArrayEleBegin
'* 1.0.10   10/3/2020   Use PigBaseMini，and add IsGetValueErrRetNothing
'*******************************************************
Imports System.Text
Friend Class fPigJSon
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.0.10"

    ''' <summary>The type of the JSON element</summary>
    Public Enum xpJSonEleType
        ''' <summary>First element</summary>
        FristEle = 0
        ''' <summary>First array element</summary>
        FristArrayEle = 5
        ''' <summary>Not the first element</summary>
        NotFristEle = 10
        ''' <summary>Element, need to be escaped</summary>
        EleValue = 20
        ''' <summary>The value of the array element, which does not need to be escaped</summary>
        ArrayValue = 100
    End Enum

    ''' <summary>Separator or control symbol in JSON</summary>
    Public Enum xpSymbolType
        ''' <summary>The end flag of the element</summary>
        EleEndFlag = 0
        '''' <summary>The begin flag of the array</summary>
        'ArrayBeginFlag = 5
        ''' <summary>The end flag of the array</summary>
        ArrayEndFlag = 10
        '''' <summary>The separator for the array</summary>
        'ArraySeparator = 20
    End Enum

    Private mbolIsParse As Boolean
    Private moSc As Object
    Private msbMain As System.Text.StringBuilder
    Private mstrLastErr As String
    Private mstrClsName = Me.GetType.Name.ToString()

    ''' <summary>
    ''' If Get value Error, return Nothing 
    ''' </summary>
    Private mbolIsGetValueErrRetNothing As Boolean = False
    Public Property IsGetValueErrRetNothing() As Boolean
        Get
            Return mbolIsGetValueErrRetNothing
        End Get
        Friend Set(ByVal value As Boolean)
            mbolIsGetValueErrRetNothing = value
        End Set
    End Property


    ''' <summary>Full JSON string, the assembled JSON string is in compact format and can be displayed in a third-party format.</summary>
    Public ReadOnly Property MainJSonStr As String
        Get
            MainJSonStr = msbMain.ToString
        End Get
    End Property


    ''' <summary>Whether the JSON has been resolved or not, the value of JSON can be read through the key.</summary>
    Public ReadOnly Property IsParse As Boolean
        Get
            IsParse = mbolIsParse
        End Get
    End Property

    ''' <summary>Parsing with assembled JSON string</summary>
    Public Function ParseJSON() As String
        Return Me.mParseJSON(Me.MainJSonStr)
    End Function


    ''' <summary>Parsing with JSON string</summary>
    ''' <param name="JSonStr">JSON string</param>
    Public Function ParseJSON(JSonStr As String) As String
        Return Me.mParseJSON(JSonStr)
    End Function

    ''' <summary>Parsing with JSON string</summary>
    ''' <param name="JSonStr">JSON string</param>
    Private Function mParseJSON(JSonStr As String) As String
        Dim strStepName As String = ""
        Try
            If Not moSc Is Nothing Then moSc = Nothing
            strStepName = "CreateObject"
            moSc = CreateObject("MSScriptControl.ScriptControl")
            With moSc
                .Language = "javascript"
                strStepName = "Reset"
                .Reset()
                JSonStr = "var json = " & JSonStr & ";"
                strStepName = "AddCode"
                .AddCode(JSonStr)
            End With
            mbolIsParse = True
            Return "OK"
        Catch ex As Exception
            mbolIsParse = False
            Me.SetSubErrInf("mParseJSON", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Private Function mDate2Lng(DateValue As DateTime, IsLocalTime As Boolean) As Long
        Dim dteStart As New DateTime(1970, 1, 1)
        Dim mtsTimeDiff As TimeSpan = DateValue - dteStart
        Try
            If IsLocalTime = True Then
                mDate2Lng = mtsTimeDiff.TotalMilliseconds
            Else
                mDate2Lng = mtsTimeDiff.TotalMilliseconds - System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours * 3600000
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mDate2Lng", ex)
            Return 0
        End Try
    End Function

    ''' <summary>Add a non array element (long value)</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="IntValue">The long or integer value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, IntValue As Long, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet As String = Me.mAddEle(EleKey, IntValue.ToString, IsFirstEle, False)
            If strRet <> "" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.IntValue", ex)
        End Try
    End Sub

    ''' <summary>Add a non array element (string value)</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="StrValue">The string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, StrValue As String, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet As String = Me.mAddEle(EleKey, StrValue, IsFirstEle, True)
            If strRet <> "" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.StrValue", ex)
        End Try
    End Sub


    ''' <summary>Add a non array element (boolean value)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="BoolValue">The boolean value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, BoolValue As Boolean, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet = Me.mAddEle(EleKey, BoolValue.ToString, IsFirstEle, False)
            If strRet <> "" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.BoolValue", ex)
        End Try
    End Sub

    ''' <summary>Add a non array element (date value)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="DecValue">The long value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, DecValue As Decimal, Optional IsFirstEle As Boolean = False)
        Try
            Dim strRet = Me.mAddEle(EleKey, DecValue.ToString, IsFirstEle, False)
            If strRet <> "" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.DateValue", ex)
        End Try
    End Sub

    ''' <summary>Add a non array element (date value)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="DateValue">The date value of the element, default is local time.</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, DateValue As DateTime, Optional IsFirstEle As Boolean = False)
        Try
            Dim lngDate As Long = Me.mDate2Lng(DateValue, True)
            If Me.LastErr <> "" Then Err.Raise(-1, , Me.LastErr)
            Dim strRet = Me.mAddEle(EleKey, lngDate.ToString, IsFirstEle, False)
            If strRet <> "" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.DateValue", ex)
        End Try
    End Sub

    ''' <summary>Add a non array element (date value, need to specify whether it is a local time)</summary>
    ''' <param name="EleKey">The key of the element, An empty string represents an array element without a key value.</param>
    ''' <param name="DateValue">The date value of the element</param>
    ''' <param name="IsLocalTime">Is it local time</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddEle(EleKey As String, DateValue As DateTime, IsLocalTime As Boolean, Optional IsFirstEle As Boolean = False)
        Try
            Dim lngDate As Long = Me.mDate2Lng(DateValue, IsLocalTime)
            If Me.LastErr <> "" Then Err.Raise(-1, , Me.LastErr)
            Dim strRet = Me.mAddEle(EleKey, lngDate.ToString, IsFirstEle, False)
            If strRet <> "" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddEle.DateValue.IsLocalTime", ex)
        End Try
    End Sub

    ''' <summary>Gets the value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    ''' <param name="OutJSonValue">The JSON value of the output</param>
    Private Function mGetJSonValue(JSonKey As String, ByRef OutJSonValue As String) As String
        Try
            JSonKey = "json." & JSonKey
            OutJSonValue = moSc.Eval(JSonKey)
            If OutJSonValue Is Nothing Then OutJSonValue = ""
            Return "OK"
        Catch ex As Exception
            OutJSonValue = ""
            Return Me.GetSubErrInf("mGetJSonValue", ex)
        End Try
    End Function

    ''' <summary>Gets the string value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetStrValue(JSonKey As String) As String
        Dim strStepName As String = "", strRet As String
        Try
            GetStrValue = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, GetStrValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetStrValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return ""
            End If
        End Try
    End Function

    ''' <summary>Gets the boolean value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetBoolValue(JSonKey As String) As Boolean
        Dim strStepName As String = "", strRet As String
        Try
            Dim strValue As String = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, strValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Convert string to boolean"
            GetBoolValue = CBool(strValue)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetBoolValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return False
            End If
        End Try
    End Function

    ''' <summary>Gets the long value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetDecValue(JSonKey As String) As Decimal
        Dim strStepName As String = "", strRet As String
        Try
            Dim strValue As String = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, strValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Convert string to long"
            GetDecValue = CDec(strValue)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetDecValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return 0
            End If
        End Try
    End Function

    ''' <summary>Gets the date value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Overloads Function GetDateValue(JSonKey As String) As DateTime
        Dim strStepName As String = "", strRet As String
        Try
            Dim strValue As String = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, strValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Convert string to datetime"
            GetDateValue = Me.mLng2Date(CLng(strValue), False)
            If Me.LastErr <> "" Then Err.Raise(-1, , Me.LastErr)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetDateValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return DateTime.MinValue
            End If
        End Try
    End Function

    ''' <summary>Gets the date value of JSON, has IsLocalTime option.</summary>
    ''' <param name="JSonKey">JSON key</param>
    ''' <param name="IsLocalTime">Is it local time</param>
    Public Overloads Function GetDateValue(JSonKey As String, IsLocalTime As Boolean) As DateTime
        Dim strStepName As String = "", strRet As String
        Try
            Dim strValue As String = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, strValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Convert string to datetime"
            GetDateValue = Me.mLng2Date(CLng(strValue), IsLocalTime)
            If Me.LastErr <> "" Then Err.Raise(-1, , Me.LastErr)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetDateValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return DateTime.MinValue
            End If
        End Try
    End Function

    Private Function mLng2Date(LngValue As Long, IsLocalTime As Boolean) As DateTime
        Dim dteStart As New DateTime(1970, 1, 1)
        Try
            If IsLocalTime = True Then
                mLng2Date = dteStart.AddMilliseconds(LngValue - System.TimeZone.CurrentTimeZone.GetUtcOffset(Now).Hours * 3600000)
            Else
                mLng2Date = dteStart.AddMilliseconds(LngValue)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetDateValue", ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return DateTime.MinValue
            End If
        End Try
    End Function


    ''' <summary>Gets the long value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetLngValue(JSonKey As String) As Long
        Dim strStepName As String = "", strRet As String
        Try
            Dim strValue As String = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, strValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Convert string to long"
            GetLngValue = CLng(strValue)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetLngValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return 0
            End If
        End Try
    End Function

    ''' <summary>Gets the integer value of JSON</summary>
    ''' <param name="JSonKey">JSON key</param>
    Public Function GetIntValue(JSonKey As String) As Integer
        Dim strStepName As String = "", strRet As String
        Try
            Dim strValue As String = ""
            strStepName = "mGetJSonValue"
            strRet = Me.mGetJSonValue(JSonKey, strValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Convert string to integer"
            GetIntValue = CInt(strValue)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("GetIntValue", strStepName, ex)
            If Me.IsGetValueErrRetNothing = True Then
                Return Nothing
            Else
                Return 0
            End If
        End Try
    End Function

    ''' <summary>Add a JSON symbol</summary>
    ''' <param name="SymbolType">Symbol type</param>
    Public Sub AddSymbol(SymbolType As xpSymbolType)
        Try
            Select Case SymbolType
                'Case xpSymbolType.ArrayBeginFlag
                '    msbMain.Append("[")
                Case xpSymbolType.ArrayEndFlag
                    msbMain.Append("]")
                'Case xpSymbolType.ArraySeparator
                '    msbMain.Append(",")
                Case xpSymbolType.EleEndFlag
                    msbMain.Append("}")
                Case Else
                    Err.Raise(-1, , "Invalid SymbolType")
            End Select
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddSymbol", ex)
        End Try
    End Sub


    ''' <summary>Add a non array JSON element</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="EleValue">The string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Private Function mAddEle(EleKey As String, EleValue As String, IsFirstEle As Boolean, IsChgCtlStr As Boolean) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "Add EleKey"
            If IsFirstEle = True Then
                If EleKey <> "" Then
                    strRet = mAddJSonStr(msbMain, xpJSonEleType.FristEle, EleKey, "", False)
                Else
                    strRet = "OK"
                End If
            Else
                strRet = mAddJSonStr(msbMain, xpJSonEleType.NotFristEle, EleKey, "", False)
            End If
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.EleValue, "", EleValue, IsChgCtlStr)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAddEle", strStepName, ex)
        End Try
    End Function

    ''' <summary>Add a array JSON element begin</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddArrayEleBegin(EleKey As String, IsFirstEle As Boolean)
        Dim strStepName As String = "", strRet As String = ""
        Try
            mSrc2JSonStr(EleKey)
            With msbMain
                If IsFirstEle = False Then .Append(",")
                .Append("""")
                .Append(EleKey)
                .Append(""":[")
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEle", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Add a array JSON element begin</summary>
    ''' <param name="EleKey">The key of the element</param>
    Public Overloads Sub AddArrayEleBegin(EleKey As String)
        Dim strStepName As String = "", strRet As String = ""
        Try
            mSrc2JSonStr(EleKey)
            With msbMain
                .Append(",")
                .Append("""")
                .Append(EleKey)
                .Append(""":[")
            End With
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEle", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Add a array JSON element</summary>
    ''' <param name="ArrayEleValue">The array string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddArrayEleValue(ArrayEleValue As String, IsFirstEle As Boolean)
        Dim strStepName As String = "", strRet As String = ""
        Try
            If IsFirstEle = False Then msbMain.Append(",")
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ArrayValue, "", ArrayEleValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEleValue", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Add one array JSON element</summary>
    ''' <param name="EleKey">The key of the element</param>
    ''' <param name="ArrayEleValue">The array string value of the element</param>
    ''' <param name="IsFirstEle">Is it the first element</param>
    Public Overloads Sub AddOneArrayEle(EleKey As String, ArrayEleValue As String, Optional IsFirstEle As Boolean = False)
        Dim strStepName As String = "", strRet As String = ""
        Try
            strStepName = "Check EleKey"
            If EleKey = "" Then Err.Raise(-1, , "Need EleKey")
            If IsFirstEle = True Then
                strRet = mAddJSonStr(msbMain, xpJSonEleType.FristEle, EleKey, "")
            Else
                strRet = mAddJSonStr(msbMain, xpJSonEleType.NotFristEle, EleKey, "")
            End If
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ArrayValue, "", "[" & ArrayEleValue & "]")
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEleValue", strStepName, ex)
        End Try
    End Sub


    ''' <summary>Add a array JSON element</summary>
    ''' <param name="ArrayEleValue">The array string value of the element</param>
    Public Overloads Sub AddArrayEleValue(ArrayEleValue As String)
        Dim strStepName As String = "", strRet As String = ""
        Try
            msbMain.Append(",")
            strStepName = "Add EleValue"
            strRet = mAddJSonStr(msbMain, xpJSonEleType.ArrayValue, "", ArrayEleValue)
            If strRet <> "OK" Then Err.Raise(-1, , strRet)
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("AddArrayEleValue", strStepName, ex)
        End Try
    End Sub

    ''' <summary>Reset so that JSON can be assembled.</summary>
    Public Sub Reset()
        Try
            If Not msbMain Is Nothing Then msbMain = Nothing
            msbMain = New System.Text.StringBuilder("")
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("Reset", ex)
        End Try
    End Sub

    Private Function mAddJSonStr(ByRef sbJSonStr As System.Text.StringBuilder, JSonEleType As xpJSonEleType, ColName As String, ColValue As String, Optional IsChgCtlStr As Boolean = False) As String
        Dim strStepName As String = "", strRet As String = ""
        Try
10:         Select Case JSonEleType
                Case xpJSonEleType.FristEle, xpJSonEleType.NotFristEle, xpJSonEleType.FristArrayEle
                    Select Case JSonEleType
                        Case xpJSonEleType.FristEle
                            sbJSonStr.Append("{")
                        Case xpJSonEleType.FristArrayEle
                            sbJSonStr.Append("[")
                        Case xpJSonEleType.NotFristEle
                            sbJSonStr.Append(",")
                    End Select
                    If ColName <> "" Then
                        mSrc2JSonStr(ColName)
                        sbJSonStr.Append("""")
                        sbJSonStr.Append(ColName)
                        sbJSonStr.Append(""":")
                    End If
                Case xpJSonEleType.EleValue, xpJSonEleType.ArrayValue
                    Select Case JSonEleType
                        Case xpJSonEleType.EleValue
                            If IsChgCtlStr = True Then mSrc2CtlStr(ColValue)
                            mSrc2JSonStr(ColValue)
                            sbJSonStr.Append("""")
                            sbJSonStr.Append(ColValue)
                            sbJSonStr.Append("""")
                        Case xpJSonEleType.ArrayValue
                            sbJSonStr.Append(ColValue)
                    End Select
                Case Else
                    strRet = "Invalid jsoneletype:" & JSonEleType.ToString
                    Err.Raise(-1, , strRet)
            End Select
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAddJSonStr", strStepName, ex)
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

    Private Sub mSrc2JSonStr(SrcStr As String)
        Try
            If InStr(SrcStr, "\") <> 0 Then SrcStr = Replace(SrcStr, "\", "\\")
            If InStr(SrcStr, """") <> 0 Then SrcStr = Replace(SrcStr, """", "\""")
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mSrc2JSonStr", ex)
        End Try
    End Sub

    Private Sub mJSonStr2Src(JSonStr As String)
        Try
            If InStr(JSonStr, "\""") <> 0 Then JSonStr = Replace(JSonStr, "\""", """")
            If InStr(JSonStr, "\\") <> 0 Then JSonStr = Replace(JSonStr, "\\", "\")
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("mJSonStr2Src", ex)
        End Try
    End Sub

    Protected Overrides Sub Finalize()
        moSc = Nothing
        MyBase.Finalize()
    End Sub

    Public Sub New()
        MyBase.New(CLS_VERSION)
        Me.Reset()
    End Sub

    Public Sub New(JSonStr As String)
        MyBase.New(CLS_VERSION)
        Dim strRet As String = Me.mParseJSON(JSonStr)
        If strRet <> "OK" Then
            mstrLastErr = strRet
        Else
            mstrLastErr = ""
        End If
    End Sub
End Class
