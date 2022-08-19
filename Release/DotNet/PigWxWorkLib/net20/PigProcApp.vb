'**********************************
'* Name: 豚豚进程应用|PigProcApp
'* Author: Seow Phong
'* License: Copyright (c) 2022 Seow Phong, For more details, see the MIT LICENSE file included with this distribution.
'* Describe: 进程相关处理|Process related processing
'* Home Url: https://www.seowphong.com or https://en.seowphong.com
'* Version: 1.1
'* Create Time: 20/3/2022
'* 1.1    26/3/2022   Modify GetPigProc(PID), Add GetPigProcs
'**********************************

Public Class PigProcApp
    Inherits PigBaseMini
    Private Const CLS_VERSION As String = "1.1.3"

    Public Sub New()
        MyBase.New(CLS_VERSION)
    End Sub

    Public Function GetPigProc(PID As Long) As PigProc
        Try
            GetPigProc = New PigProc(PID)
            If GetPigProc.LastErr <> "" Then Throw New Exception(GetPigProc.LastErr)
        Catch ex As Exception
            Me.SetSubErrInf("GetPigProc", ex)
            Return Nothing
        End Try
    End Function

    Public Function GetPigProcs(ProcName As String) As PigProcs
        Dim LOG As New PigStepLog("GetPigProcs")
        Try
            LOG.StepName = "GetProcessesByName"
            Dim abProcess As Process() = Process.GetProcessesByName(ProcName)
            LOG.StepName = "New PigProcs"
            GetPigProcs = New PigProcs
            For Each oProcess As Process In abProcess
                LOG.StepName = "Add"
                GetPigProcs.Add(oProcess.Id)
                If GetPigProcs.LastErr <> "" Then
                    LOG.AddStepNameInf(oProcess.Id)
                    Throw New Exception(GetPigProcs.LastErr)
                End If
            Next
        Catch ex As Exception
            LOG.AddStepNameInf(ProcName)
            Me.SetSubErrInf(LOG.SubName, LOG.StepName, ex)
            Return Nothing
        End Try
    End Function

End Class
