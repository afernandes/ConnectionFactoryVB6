VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBCommandAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private bClassInit As Boolean
Private bCommandFinished As Boolean
Private bRet As Boolean

Private WithEvents DBConn As DBConnection
Attribute DBConn.VB_VarHelpID = -1
Private WithEvents oRS As ADODB.Recordset
Attribute oRS.VB_VarHelpID = -1
Private WithEvents oTmr As timer
Attribute oTmr.VB_VarHelpID = -1

Public Event ReturnExecuteAsync(rs As ADODB.Recordset)
Attribute ReturnExecuteAsync.VB_MemberFlags = "200"
Public Event ErrorExecuteAsync(ByVal SQLerr As ErrObject)
Public Event ChangeStatus(status As String)

Private Sub Class_Initialize()
    bClassInit = False
    bCommandFinished = False
End Sub
Private Sub Class_Terminate()
    Call ExecuteCancel
    bClassInit = False
    bCommandFinished = False
End Sub

Friend Function Init(oDBConn As DBConnection) As DBCommandAsync
    Set DBConn = oDBConn

    Set oRS = New Recordset

    oRS.CursorLocation = adUseClient
    oRS.CursorType = adOpenStatic
    oRS.LockType = adLockBatchOptimistic
    oRS.Properties("Initial Fetch Size") = 0
    oRS.Properties("Background Fetch Size") = 1000

    bRet = False

    bClassInit = True
    Set Init = Me
End Function

Public Sub ExecuteCancel()
    On Error Resume Next

    If oRS.State = adStateClosed Then Exit Sub

    RaiseEvent ChangeStatus("Cancelando tarefa...")
    Do While oRS.State <> adStateClosed
        oRS.Cancel
        oRS.Close
        DoEvents
    Loop

    oTmr.Enabled = False
    RaiseEvent ChangeStatus("Tarefa cancelada...")
End Sub

Public Function ExecuteReaderAsync(cmdType As CommandTypeEnum, ByVal cmdText As String, Optional ByVal oParams As DBParameters = Nothing, Optional RecordsAffected As Long = 0) As Recordset
Attribute ExecuteReaderAsync.VB_UserMemId = 0
    Dim ocmd As New ADODB.Command

    On Error GoTo ExecuteReaderAsync_Error

    If Not bClassInit Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".ExecuteReaderAsync()", "Class uninitialized." & vbCrLf & _
                                                                             "Use: DBConnection.GetDBCommand()"
        Exit Function
    End If

    Set ocmd = PrepareCommand(cmdType, cmdText, oParams)
    
    If RecordsAffected > 0 Then
        Set oRS = ocmd.Execute(RecordsAffected, , adAsyncExecute + adAsyncFetch + adAsyncConnect)
    Else
        Set oRS = ocmd.Execute(, , adAsyncExecute + adAsyncFetch + adAsyncConnect)
    End If

    Do While ocmd.State And adStateFetching Or ocmd.State And adStateExecuting
        DoEvents
    Loop

    '    Do While Not bCommandFinished
    '        DoEvents
    '    Loop

    Set ExecuteReaderAsync = oRS

    On Error GoTo 0
    Exit Function

ExecuteReaderAsync_Error:
    Debug.Print "Error: &h" & Hex$(Err.Number) & " :: " & Err.Description & vbCrLf & _
                "--Trace: " & Err.Source & "<-ExecuteReaderAsync<-DBCommandAsync"

    Debug.Assert False
    Err.Raise Err.Number, "Class.DBCommandAsync.ExecuteReaderAsync", Err.Description

End Function


Private Sub oRS_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'#AndersonN - 14/12/2012
    Dim oErr As New ErrObject

    On Error GoTo oRS_FetchComplete_Error

    oTmr.Enabled = False
    RaiseEvent ChangeStatus("Tarefa concluida...")
    If Not bRet Then
        DoEvents
        If Not pError Is Nothing Then
            oErr.Description = pError.Description & vbCrLf & "NativeError: " & pError.NativeError & vbCrLf & "SQLState: " & pError.SQLState
            oErr.Number = pError.Number
            oErr.HelpContext = pError.HelpContext
            oErr.HelpFile = pError.HelpFile
            oErr.Source = pError.Source
            RaiseEvent ErrorExecuteAsync(oErr)
        Else
            If pRecordset.State = adStateOpen Then
                'm_lMaxTop = pRecordset.RecordCount
                RaiseEvent ReturnExecuteAsync(pRecordset)
                If oRS.State = adStateOpen Then oRS.Close
            End If
        End If
    End If
    bRet = True

    On Error GoTo 0
    Exit Sub

oRS_FetchComplete_Error:
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        SQLAsync->oRS_FetchComplete"

    Debug.Assert False

    oTmr.Enabled = False
    RaiseEvent ChangeStatus("Erro na consulta...")
    RaiseEvent ErrorExecuteAsync(Err)

End Sub

Private Sub oRS_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
'#AndersonN - 15/08/2013
    On Error GoTo oRS_FetchProgress_Error

    '    With prg1
    '        .Max = IIf(m_lMaxTop > MaxProgress, m_lMaxTop, MaxProgress)
    '        .Value = Progress
    '        .Visible = True
    '    End With

    'SU_ChangeStatus "Obtendo dados ( " & Fix(Progress / IIf(m_lMaxTop > MaxProgress, m_lMaxTop, MaxProgress) * 100) & "% )"
    RaiseEvent ChangeStatus("Obtendo dados ( " & Fix(Progress / MaxProgress * 100) & "% )")

    On Error GoTo 0
    Exit Sub

oRS_FetchProgress_Error:
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        SQLAsync->oRS_FetchProgress"

    Debug.Assert False
End Sub

Private Sub oTmr_Timer()
    On Error GoTo oTmr_Timer_Error

    If oRS.State = adStateClosed Then
        oTmr.Enabled = False
        RaiseEvent ChangeStatus("Consulta Cancelada...")
        RaiseEvent ReturnExecuteAsync(oRS)
    ElseIf oRS.State = adStateOpen Then
        oTmr.Enabled = False
        If Not bRet Then
            bRet = True
            DoEvents
            RaiseEvent ChangeStatus("Tarefa concluida...")
            'm_lMaxTop = oRS.RecordCount
            RaiseEvent ReturnExecuteAsync(oRS)
            If oRS.State = adStateOpen Then oRS.Close
        End If
    End If

    On Error GoTo 0
    Exit Sub

oTmr_Timer_Error:
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        SQLAsync->oTmr_Timer"
    Debug.Assert False

    oTmr.Enabled = False
    RaiseEvent ChangeStatus("Erro na consulta...")
    RaiseEvent ErrorExecuteAsync(Err)
End Sub


Private Sub DBConn_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    CursorType = adOpenStatic
    LockType = adLockBatchOptimistic
    Options = adAsyncExecute + adAsyncFetch + adAsyncConnect

    bCommandFinished = False
    RaiseEvent ChangeStatus("Excecute initialized...")
End Sub
Private Sub DBConn_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    bCommandFinished = True
    RaiseEvent ChangeStatus("Excecute complete...")
End Sub

Private Function PrepareCommand(cmdType As CommandTypeEnum, cmdText As String, cmdParams As DBParameters) As ADODB.Command
    Dim cmd As New ADODB.Command
    Dim param As New DBParameter

    On Error GoTo PrepareCommand_Error

    cmd.ActiveConnection = DBConn.getConnection()
    cmd.CommandText = cmdText
    cmd.CommandType = cmdType

    If Not cmdParams Is Nothing Then
        For Each param In cmdParams
            cmd.Parameters.Append cmd.CreateParameter(param.Name, adVarChar, adParamInput, Len(param.Value), Value:=param.Value)
            With cmd.Parameters.Item(param.Name)
                Select Case param.TypeParam
                Case enmTypeParameter.Character
                    .Type = adVarChar
                Case enmTypeParameter.Date
                    .Type = adDate
                Case enmTypeParameter.Numeric
                    .Type = adNumeric
                    Dim p As Byte
                    Dim s As Byte
                    fixPrecisionAndScale param.Value, p, s
                    .Precision = p
                    .NumericScale = s
                End Select
                .Value = param.Value
            End With
        Next
    End If

    Set PrepareCommand = cmd

    On Error GoTo 0
    Exit Function

PrepareCommand_Error:
    Debug.Print "Error: &h" & Hex$(Err.Number) & " :: " & Err.Description & vbCrLf & _
                "--Trace: " & Err.Source & "<-PrepareCommand<-DBCommandAsync"

    Debug.Assert False
    Err.Raise Err.Number, "Class.DBCommandAsync.PrepareCommand", Err.Description

End Function

'Auxiliar methods
Private Sub fixPrecisionAndScale(ByVal pValue As String, ByRef retPrecision As Byte, ByRef retScale As Byte)
    Dim sep As String
    sep = getDecimalSeparator()

    If IsNumeric(pValue) Then
        If InStr(1, pValue, sep) > 0 Then
            retPrecision = Len(Replace(pValue, sep, ""))
            retScale = Len(Split(pValue, sep)(1))
        End If
    End If
End Sub

Private Function getDecimalSeparator() As String
    getDecimalSeparator = Mid(CStr(1.1), 2, 1)
End Function

