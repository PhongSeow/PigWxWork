'**********************************
'* Name: PigWebReq
'* Author: Seow Phong
'* License: Copyright (c) 2020 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Http Web Request operation
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'* Create Time: 5/2/2021
'*1.0.2  25/2/2021   Add Me.ClearErr()
'*1.0.3  9/3/2021   Modify GetText,GetTextAuth,PostText,PostTextAuth
'**********************************
Imports System.Net
Imports System.IO
Imports System.Text
Public Class PigWebReq
    Inherits PigBaseMini
    Const CLS_VERSION As String = "1.0.3"
    Private mstrUrl As String
    Private mstrPara As String
    Private muriMain As System.Uri
    Private mhwrMain As HttpWebRequest
    Private msrRes As StreamReader
    Private mstrResString As String
    Private mstrUserAgent As String
    Public UseTimeItem As New UseTime

    Public Sub New(Url As String, Para As String, UserAgent As String)
        MyBase.New(CLS_VERSION)
        Me.MainNew(Url, Para, UserAgent)
    End Sub

    Public Sub New(Url As String)
        MyBase.New(CLS_VERSION)
        Me.MainNew(Url, "", "")
    End Sub

    Public Sub New(Url As String, UserAgent As String)
        MyBase.New(CLS_VERSION)
        Me.MainNew(Url, "", UserAgent)
    End Sub

    Public ReadOnly Property ResString As String
        Get
            ResString = mstrResString
        End Get
    End Property

    Public ReadOnly Property UserAgent As String
        Get
            UserAgent = mstrUserAgent
        End Get
    End Property

    Private Sub MainNew(Url As String, Para As String, UserAgent As String)
        Dim strStepName As String = ""
        Try
            If Len(Para) = 0 Then
                strStepName = "设置Uri" & Url & ":"
10:             muriMain = New System.Uri(Url)
            Else
                strStepName = "设置Uri" & Url & "?" & Para & ":"
20:             muriMain = New System.Uri(Url & "?" & Para)
            End If
            If Len(UserAgent) > 0 Then
                mstrUserAgent = UserAgent
            Else
                mstrUserAgent = "GEHttpWebItem"
            End If
            strStepName = "HttpWebRequest.Create:"
30:         mhwrMain = System.Net.HttpWebRequest.Create(muriMain)
            mhwrMain.UserAgent = mstrUserAgent
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf("MainNew", strStepName, ex)
        End Try
    End Sub

    Public Function GetText() As String
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        GetText = ""
        Try
10:         mhwrMain.Method = "GET"
            strStepName = "New StreamReader"
20:         msrRes = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
            strStepName = "ReadToEnd"
30:         mstrResString = msrRes.ReadToEnd
40:         msrRes.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("GetText", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function GetTextAuth(AccessToken As String) As String
        'AccessToken：在接入端中设置
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        GetTextAuth = ""
        Try
            'mhwrMain.UseDefaultCredentials = False
            With mhwrMain
                .Headers("Authorization") = "Bearer " & AccessToken
                '                .Headers("value") = "Bearer " & AccessToken
                .Method = "GET"
                .PreAuthenticate = True
            End With
            strStepName = "New StreamReader"
20:         msrRes = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
            strStepName = "ReadToEnd"
30:         mstrResString = msrRes.ReadToEnd
40:         msrRes.Close()
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("GetTextAuth", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function PostTextAuth(Para As String, AccessToken As String) As String
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        PostTextAuth = ""
        Try
10:         mhwrMain.Method = "POST"
            mhwrMain.ContentType = "application/x-www-form-urlencoded"
            mhwrMain.Headers("AUTHORIZATION") = "Bearer " & AccessToken
            mhwrMain.PreAuthenticate = True
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(Para)
20:         mhwrMain.ContentLength = bys.Length
            strStepName = "mhwrMain.GetRequestStream():"
            Dim newStream As Stream = mhwrMain.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
30:         newStream.Close()
            strStepName = "mhwrMain.GetResponse().GetResponseStream():"
            Dim sr As StreamReader = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
40:         mstrResString = sr.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("PostTextAuth", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

    Public Function PostText(Para As String) As String
        Dim strStepName As String = ""
        Me.UseTimeItem.GoBegin()
        PostText = ""
        Try
10:         mhwrMain.Method = "POST"
            mhwrMain.ContentType = "application/x-www-form-urlencoded"
            Dim encoding As New UTF8Encoding()
            Dim bys As Byte() = encoding.GetBytes(Para)
20:         mhwrMain.ContentLength = bys.Length
            strStepName = "mhwrMain.GetRequestStream():"
            Dim newStream As Stream = mhwrMain.GetRequestStream()
            newStream.Write(bys, 0, bys.Length)
30:         newStream.Close()
            strStepName = "mhwrMain.GetResponse().GetResponseStream():"
            Dim sr As StreamReader = New StreamReader(mhwrMain.GetResponse().GetResponseStream)
40:         mstrResString = sr.ReadToEnd
            Me.UseTimeItem.ToEnd()
            Me.ClearErr()
            Return "OK"
        Catch ex As Exception
            Me.UseTimeItem.ToEnd()
            Me.SetSubErrInf("PostText", strStepName, ex)
            Return Me.LastErr
        End Try
    End Function

End Class
