VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBReaderAsync"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents cTmr As CTimer
Attribute cTmr.VB_VarHelpID = -1

Private WithEvents DBConn As DBConnection
Attribute DBConn.VB_VarHelpID = -1
Private WithEvents oRS As ADODB.Recordset
Attribute oRS.VB_VarHelpID = -1

Public Event SQLret(ByVal oRS As ADODB.Recordset, ByVal bSucess As Boolean)
Attribute SQLret.VB_MemberFlags = "200"
Public Event SQLStatus(ByVal sMessage As String, percentage As Integer)
Public Event SQLError(ByVal SQLerr As ErrObject)

Private bClassInit As Boolean
Private bRet As Boolean

Private Sub Class_Initialize()
    Set oRS = New Recordset

    oRS.CursorLocation = adUseClient
    oRS.CursorType = adOpenStatic
    oRS.LockType = adLockBatchOptimistic
    oRS.Properties("Initial Fetch Size") = 0
    oRS.Properties("Background Fetch Size") = 1000

    Set cTmr = New CTimer

    bRet = False
End Sub

Private Sub UserControl_Terminate()
    Call ExecCancel
End Sub

Friend Function Init(oDBConn As DBConnection) As DBReaderAsync
    Set DBConn = oDBConn

    bRet = False

    bClassInit = True
    Set Init = Me
End Function

Private Sub oRS_FetchProgress(ByVal Progress As Long, ByVal MaxProgress As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    On Error GoTo oRS_FetchProgress_Error

    If MaxProgress > 0 Then
        RaiseEvent SQLStatus("Recebendo ", Fix(Progress / MaxProgress * 100))
    Else
        RaiseEvent SQLStatus("Processando... ", 0)
    End If

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


Private Sub oRS_FetchComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
    Dim oErr As New ErrObject

    On Error GoTo oRS_FetchComplete_Error

    cTmr.Interval = 0
    RaiseEvent SQLStatus("Tarefa concluida...", 0)
    If Not bRet Then
        DoEvents
        If Not pError Is Nothing Then
            oErr.Description = pError.Description & vbCrLf & "NativeError: " & pError.NativeError & vbCrLf & "SQLState: " & pError.SQLState
            oErr.Number = pError.Number
            oErr.HelpContext = pError.HelpContext
            oErr.HelpFile = pError.HelpFile
            oErr.Source = pError.Source
            RaiseEvent SQLError(oErr)
        Else
            If pRecordset.State = adStateOpen Then
                RaiseEvent SQLret(pRecordset, True)
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

    cTmr.Interval = 0
    RaiseEvent SQLStatus("Erro na consulta...", 0)
    RaiseEvent SQLError(Err)

End Sub

Public Sub ExecCancel()
    On Error Resume Next

    If oRS.State = adStateClosed Then Exit Sub

    RaiseEvent SQLStatus("Cancelando tarefa...", 0)
    Do While oRS.State <> adStateClosed
        oRS.Cancel
        oRS.Close
        DoEvents
    Loop

    cTmr.Interval = 0
    RaiseEvent SQLStatus("Tarefa cancelada...", 0)
End Sub

Public Sub ExecuteReaderAsync(ByVal cmdText As String, Optional ByVal oParams As DBParameters = Nothing, Optional ByVal FetchSize = 1000)
Attribute ExecuteReaderAsync.VB_UserMemId = 0
    If Not bClassInit Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".ExecuteReaderAsync()", "Class uninitialized." & vbCrLf & _
                                                                             "Use: DBConnection.GetDBReaderAsync()"
        Exit Sub
    End If

    On Error GoTo ExecuteReaderAsync_Error

    bRet = False

    Call ExecCancel

    If oRS.State <> adStateClosed Then oRS.Close

    If DBConn.getConnection().State <> adStateOpen Then
        RaiseEvent SQLStatus("Conexão inativa...", 0)
        Exit Sub
    End If

    oRS.Properties("Background Fetch Size") = FetchSize

    RaiseEvent SQLStatus("Executando tarefa...", 0)
    cTmr.Interval = 100

    cmdText = PrepareCommand(oParams) & cmdText

    oRS.Open cmdText, DBConn.getConnection(), adOpenStatic, adLockBatchOptimistic, adAsyncExecute + adAsyncFetch

    On Error GoTo 0
    Exit Sub

ExecuteReaderAsync_Error:
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        SQLAsync->ExecuteReaderAsync"
    Debug.Assert False

    cTmr.Interval = 0
    RaiseEvent SQLStatus("Erro na execução...", 0)
    RaiseEvent SQLError(Err)

End Sub

Private Sub cTmr_ThatTime()
    On Error GoTo TimerProc_Error

    If oRS.State = adStateClosed Then
        cTmr.Interval = 0
        RaiseEvent SQLStatus("Consulta Cancelada...", 0)
        RaiseEvent SQLret(oRS, False)
    ElseIf oRS.State = adStateOpen Then
        cTmr.Interval = 0
        If Not bRet Then
            bRet = True
            DoEvents
            RaiseEvent SQLStatus("Tarefa concluida...", 0)
            RaiseEvent SQLret(oRS, True)
            If oRS.State = adStateOpen Then oRS.Close
        End If
    End If

    On Error GoTo 0
    Exit Sub

TimerProc_Error:
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        SQLAsync->TimerProc"
    Debug.Assert False

    cTmr.Interval = 0
    RaiseEvent SQLStatus("Erro na consulta...", 0)
    RaiseEvent SQLError(Err)
End Sub

Private Function PrepareCommand(cmdParams As DBParameters) As String
    Dim param As New DBParameter
    Dim returnString As String

    On Error GoTo PrepareCommand_Error

    If Not cmdParams Is Nothing Then
        For Each param In cmdParams
            With param
                Select Case param.TypeParam
                Case enmTypeParameter.Character
                    returnString = returnString & _
                                   "DECLARE " & .Name & " VARCHAR(" & Len(.Value) & "); " & vbCrLf & _
                                   "SET " & .Name & " = " & SQL.FU_Asp(.Value, True) & "; " & vbCrLf
                Case enmTypeParameter.Date
                    returnString = returnString & _
                                "DECLARE " & .Name & " DATETIME; " & vbCrLf & _
                                   "SET " & .Name & " = " & SQL.FU_Date(.Value) & "; " & vbCrLf
                Case enmTypeParameter.Numeric
                    Dim p As Byte
                    Dim s As Byte
                    fixPrecisionAndScale .Value, p, s
                    
                    returnString = returnString & _
                                   "DECLARE " & .Name & " NUMERIC(" & p & "," & s & "); " & vbCrLf & _
                                   "SET " & .Name & " = " & SQL.FU_Asp(.Value, True) & "; " & vbCrLf
                End Select
            End With
        Next
    End If

    PrepareCommand = returnString

    On Error GoTo 0
    Exit Function

PrepareCommand_Error:
    Debug.Print "Error: &h" & Hex$(Err.Number) & " :: " & Err.Description & vbCrLf & _
                "--Trace: " & Err.Source & "<-PrepareCommand<-" & TypeName(Me)

    Debug.Assert False
    Err.Raise Err.Number, "Class." & TypeName(Me) & ".PrepareCommand", Err.Description

End Function

Private Sub fixPrecisionAndScale(ByVal pValue As String, ByRef retPrecision As Byte, ByRef retScale As Byte)
    Dim sep As String
    sep = getDecimalSeparator()

    If IsNumeric(pValue) Then
        If InStr(1, pValue, sep) > 0 Then
            retPrecision = Len(Replace(pValue, sep, ""))
            retScale = Len(Split(pValue, sep)(1))
        Else
            retPrecision = Len(pValue)
            retScale = 0
        End If
    End If
End Sub

Private Function getDecimalSeparator() As String
    getDecimalSeparator = Mid(CStr(1.1), 2, 1)
End Function

