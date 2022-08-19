'**********************************
'* Name: PigProcs
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: PigProc 的集合类|Collection class of PigProc
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 20/3/2022
'* 1.1    26/3/2022   Modify New
'************************************
Public Class PigProcs
    Inherits PigBaseMini
    Implements IEnumerable(Of PigProc)
    Private Const CLS_VERSION As String = "1.0.1"
    Private ReadOnly moList As New List(Of PigProc)

    Public Sub New()
        MyBase.New(CLS_VERSION)
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
    Public Function GetEnumerator() As IEnumerator(Of PigProc) Implements IEnumerable(Of PigProc).GetEnumerator
        Return moList.GetEnumerator()
    End Function

    Public Function GetEnumerator1() As IEnumerator Implements IEnumerable.GetEnumerator
        Return Me.GetEnumerator()
    End Function

    Public ReadOnly Property Item(Index As Integer) As PigProc
        Get
            Try
                Return moList.Item(Index)
            Catch ex As Exception
                Me.SetSubErrInf("Item.Index", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property Item(ProcessID As String) As PigProc
        Get
            Try
                Item = Nothing
                For Each oPigProc As PigProc In moList
                    If oPigProc.ProcessID = ProcessID Then
                        Item = oPigProc
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Me.SetSubErrInf("Item.ProcessID", ex)
                Return Nothing
            End Try
        End Get
    End Property

    Public Function IsItemExists(ProcessID) As Boolean
        Try
            IsItemExists = False
            For Each oPigProc As PigProc In moList
                If oPigProc.ProcessID = ProcessID Then
                    IsItemExists = True
                    Exit For
                End If
            Next
        Catch ex As Exception
            Me.SetSubErrInf("IsItemExists", ex)
            Return False
        End Try
    End Function

    Private Function mAdd(NewItem As PigProc) As String
        Try
            If Me.IsItemExists(NewItem.ProcessID) = True Then Throw New Exception(NewItem.ProcessID & " already exists.")
            moList.Add(NewItem)
            Return "OK"
        Catch ex As Exception
            Return Me.GetSubErrInf("mAdd", ex)
        End Try
    End Function



    Public Function Add(ProcessID As String) As PigProc
        Dim LOG As New PigStepLog("Remove.ProcessID")
        Try
            LOG.StepName = "New PigProc"
            Add = New PigProc(ProcessID)
            If Add.LastErr <> "" Then
                LOG.AddStepNameInf(ProcessID)
                Throw New Exception(Add.LastErr)
            End If
            LOG.StepName = "mAdd"
            LOG.Ret = Me.mAdd(Add)
            If LOG.Ret <> "OK" Then
                LOG.AddStepNameInf(ProcessID)
                Throw New Exception(LOG.Ret)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function


    Public Function Remove(ProcessID As String) As String
        Dim LOG As New PigStepLog("Remove.ProcessID")
        Try
            LOG.StepName = "For Each"
            For Each oPigProc As PigProc In moList
                If oPigProc.ProcessID = ProcessID Then
                    LOG.AddStepNameInf(ProcessID)
                    moList.Remove(oPigProc)
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
            Return Me.GetSubErrInf(LOG.SubName, LOG.StepName, ex)
        End Try
    End Function

    Public Function AddOrGet(ProcessID As String) As PigProc
        Dim LOG As New PigStepLog("AddOrGet")
        Try
            If Me.IsItemExists(ProcessID) = True Then
                AddOrGet = Me.Item(ProcessID)
            Else
                AddOrGet = Me.Add(ProcessID)
            End If
            Me.ClearErr()
        Catch ex As Exception
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
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

