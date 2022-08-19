'**********************************
'* Name: UpdateCheck
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: Whether the processing attribute has been updated
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.0.3
'*1.1  26/6/2022    Add HasUpdated
'* Create Time: 25/6/2019
'************************************

Public Class UpdateCheck
    Private Property mKeyList As String = ""
    Public Property LastUpdateTime As DateTime = #1/1/1900#

    Public Sub Clear()
        Me.mKeyList = ""
        Me.LastUpdateTime = #1/1/1900#
    End Sub

    Public Sub Add(KeyName As String)
        KeyName = "<" & KeyName & ">"
        If InStr(Me.mKeyList, KeyName) = 0 Then
            Me.mKeyList &= KeyName
        End If
    End Sub

    Public Sub Remove(KeyName As String)
        KeyName = "<" & KeyName & ">"
        If InStr(Me.mKeyList, KeyName) > 0 Then
            Me.mKeyList = Replace(Me.mKeyList, KeyName, "")
        End If
    End Sub

    Public Function IsUpdated(KeyName As String) As Boolean
        KeyName = "<" & KeyName & ">"
        If InStr(Me.mKeyList, KeyName) > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function HasUpdated() As Boolean
        If Me.mKeyList = "" Then
            Return False
        Else
            Return True
        End If
    End Function

End Class
