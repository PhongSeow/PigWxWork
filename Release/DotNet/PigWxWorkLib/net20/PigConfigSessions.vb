'**********************************
'* Name: PigConfigSessions
'* Author: Seow Phong
'* License: Copyright (c) 2021 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigConfigSession 的集合类|Collection class of PigConfigSession
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.7
'* Create Time: 21/12/2021
'* 1.1    23/12/2020   Add Parent, modify New,Add
'* 1.2    24/12/2020   Add Clear
'* 1.3    25/12/2021   Add AddOrGet
'* 1.4    26/12/2021   Modify AddOrGet
'* 1.5    3/1/2022   Modify AddOrGet,Add
'* 1.6    21/2/2022   Modify mAdd,Remove,Clear
'* 1.7    9/5/2022   Remove IsChange
'************************************
Public Class PigConfigSessions
    Inherits PigBaseMini
    Implements IEnumerable(Of PigConfigSession)
    Private Const CLS_VERSION As String = "1.7.2"

    Friend Property Parent As PigConfigApp
    Private ReadOnly moList As New List(Of PigConfigSession)

    Public Sub New(Parent As PigConfigApp)
        MyBase.New(CLS_VERSION)
        Me.Parent = Parent
    End Sub

    Public ReadOnly Property Count() As Integer
        Get
            Try
                Return moList.Count
            Catch ex As Exception
                Me.SetSubErrInf("Count", ex)
                Return -1
            End Try
        End Get
    End Property
    Public Function GetEnumerator() As IEnumerator(Of PigConfigSession) Implements IEnumerable(Of PigConfigSession).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigConfigSession
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(SessionName As String) As PigConfigSession
        Get
            Try
                Item = Nothing
                For Each oPigConfigSession As PigConfigSession In moList
                    If oPigConfigSession.SessionName = SessionName Then
                        Item = oPigConfigSession
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.SessionName", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(SessionName) As Boolean
        Try
            IsItemExists = False
            For Each oPigConfigSession As PigConfigSession In moList
                If oPigConfigSession.SessionName = SessionName Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As PigConfigSession) As String
        Dim LOG As New PigStepLog("mAdd")
        Try
            If Me.IsItemExists(NewItem.SessionName) = True Then
                LOG.StepName = "Check IsItemExists"
                Throw New Exception(NewItem.SessionName & " already exists.")
            End If
            LOG.StepName = "List.Add"
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AddOrGet(SessionName As String) As PigConfigSession
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(SessionName) = True Then
                AddOrGet = Me.Item(SessionName)
            Else
                AddOrGet = Me.Add(SessionName)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function AddOrGet(SessionName As String, SessionDesc As String) As PigConfigSession
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(SessionName) = True Then
                AddOrGet = Me.Item(SessionName)
            Else
                AddOrGet = Me.Add(SessionName, SessionDesc)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(SessionName As String) As PigConfigSession
        Dim LOG As New PigStepLog("Add")
        Try
            LOG.StepName = "New PigConfigSession"
            Add = New PigConfigSession(SessionName, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(SessionName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(SessionName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Add(SessionName As String, SessionDesc As String) As PigConfigSession
        Dim LOG As New PigStepLog("Add")
        Try
            LOG.StepName = "New PigConfigSession"
            Add = New PigConfigSession(SessionName, SessionDesc, Me.Parent)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(SessionName)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(SessionName)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

    Public Function Remove(SessionName As String) As String
        Dim LOG As New PigStepLog("Remove.SessionName")
        Try
            LOG.StepName = "For Each To Remove"
            For Each oPigConfigSession As PigConfigSession In moList
                If oPigConfigSession.SessionName = SessionName Then
                    LOG.AddStepNameInf(SessionName)
                    moList.Remove(oPigConfigSession)
                    Exit For
                End If
            Next
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function Remove(Index As Integer) As String
        Dim LOG As New PigStepLog("Remove.Index")
        Try
            LOG.StepName = "Index=" & Index.ToString
            moList.RemoveAt(Index)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf(Log.SubName, Log.StepName, ex)
        End Try
    End Function

    Public Function Clear() As String
        Try
            moList.Clear()
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("Clear", ex)
        End Try
    End Function

End Class


