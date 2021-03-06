VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private WithEvents oDB As ADODB.Connection
Attribute oDB.VB_VarHelpID = -1

Public Enum TransactionType
    enmTransactionOpen = 1
    enmTransactionCommit = 2
    enmTransactionRollback = 3
End Enum

Private m_ConnectionString As String
Private m_Alias As String

Private lTranLevel As Long


Public Event WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
Public Event ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
Attribute ExecuteComplete.VB_MemberFlags = "200"

'[Default]
Public Function Init(sConnectionString As String) As Boolean
Attribute Init.VB_UserMemId = 0
    m_ConnectionString = sConnectionString

    EstablishFactoryConnection
    Init = True
End Function

'Get DBCommand
Public Function GetDBCommand() As DBCommand
    Dim cmd As DBCommand
    Set cmd = New DBCommand

    Set GetDBCommand = cmd.Init(Me)

End Function

'Get DBReaderAsync
Public Function GetDBReaderAsync() As DBReaderAsync
    Dim reader As DBReaderAsync
    Set reader = New DBReaderAsync

    Set GetDBReaderAsync = reader.Init(Me)

End Function


'Get DBCommandAsync
Public Function GetDBCommandAsync() As DBCommandAsync
    Dim cmd As DBCommandAsync
    Set cmd = New DBCommandAsync

    Set GetDBCommandAsync = cmd.Init(Me)

End Function

'Get DBTransaction
Public Function GetDBTransaction() As DBTransaction
    Dim tran As DBTransaction
    Set tran = New DBTransaction

    Set GetDBTransaction = tran.Init(Me)

End Function


Private Sub Class_Initialize()
    lTranLevel = 0
End Sub
Private Sub Class_Terminate()
    Me.CloseFactoryConnection
    lTranLevel = 0
End Sub

Friend Function getTransactionLevel() As Long
    getTransactionLevel = lTranLevel
End Function

Friend Property Get getConnection() As ADODB.Connection
    Set getConnection = oDB
End Property

Friend Function TransactionHandler(veTransactionType As TransactionType) As Long
    On Error GoTo TransactionHandler_Error

    TransactionHandler = True

    Select Case veTransactionType
    Case enmTransactionOpen
        TransactionHandler = oDB.BeginTrans
    Case enmTransactionCommit
        oDB.CommitTrans
    Case enmTransactionRollback
        oDB.RollbackTrans
    End Select

    On Error GoTo 0
    Exit Function

TransactionHandler_Error:
    TransactionHandler = False
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        DBHelper->TransactionHandler"

    Debug.Assert False
    Err.Raise 2, "Class.DBHelper.TransactionHandler", "DBHelper->TransactionHandler(" & veTransactionType & ") :" & Err.Description
End Function

Friend Property Get ConnectionString() As String
    ConnectionStrings = oDB.ConnectionString
End Property
Friend Property Let ConnectionStrings(ByVal vData As String)
    oDB.ConnectionString = vData
End Property

Friend Property Let Alias(ByVal NewAlias As String)
    m_Alias = NewAlias
End Property
Friend Property Get Alias() As String
    Alias = m_Alias
End Property

'---------------------------------------------------------------------------------------
' Procedure : EstablishFactoryConnection
' Author    : ander
' Date      : 27/04/2016
' Purpose   :
'---------------------------------------------------------------------------------------
'
Friend Sub EstablishFactoryConnection()

   On Error GoTo EstablishFactoryConnection_Error

    If oDB Is Nothing Then
        Set oDB = New ADODB.Connection
    End If

    If oDB.State = adStateClosed Then
        oDB.Open m_ConnectionString
    End If

   On Error GoTo 0
   Exit Sub

EstablishFactoryConnection_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure EstablishFactoryConnection of M�dulo de classe DBConnection"
End Sub

Friend Sub CloseFactoryConnection()
    If Not oDB Is Nothing Then
        If oDB.State <> adStateClosed Then
            oDB.Close
        End If
        Set oDB = Nothing
    End If
End Sub


Private Sub oDB_BeginTransComplete(ByVal TransactionLevel As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If Not pError Is Nothing Then
        If pError.Number <> 0 Then
            Err.Raise pError.Number, "Class.DBConneciont.Event.BeginTransComplete [" & pError.Source & "]", pError.Description & vbCrLf & "SQLState: " & pError.SQLState
            Exit Sub
        End If
    Else
        If adStatus = adStatusOK Then
            If TransactionLevel > 0 Then
                lTranLevel = TransactionLevel
            Else
                Err.Raise 1, "Class.DBConneciont.Event.BeginTransComplete", "Nivel invalido de transa��o! " & vbCrLf & "TransactionLevel: " & TransactionLevel
            End If
        Else
            Err.Raise 1, "Class.DBConneciont.Event.BeginTransComplete", "N�o foi possivel iniciar a transa��o! " & vbCrLf & "Status: " & adStatus
        End If
    End If
End Sub

Private Sub oDB_CommitTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If Not pError Is Nothing Then
        If pError.Number <> 0 Then
            Err.Raise pError.Number, "Class.DBConneciont.Event.CommitTransComplete [" & pError.Source & "]", pError.Description & "[" & pError.SQLState & "]"
            Exit Sub
        End If
    Else
        If adStatus = adStatusOK Then
            lTranLevel = lTranLevel - 1
        Else
            Err.Raise 1, "Class.DBConneciont.Event.CommitTransComplete", "N�o foi possivel concluir a transa��o com sucesso! " & vbCrLf & "Status: " & adStatus
        End If
    End If
End Sub

Private Sub oDB_RollbackTransComplete(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    If Not pError Is Nothing Then
        If pError.Number <> 0 Then
            Err.Raise pError.Number, "Class.DBConneciont.Event.RollbackTransComplete [" & pError.Source & "]", pError.Description & "[" & pError.SQLState & "]"
            Exit Sub
        End If
    Else
        If adStatus = adStatusOK Then
            lTranLevel = lTranLevel - 1
        Else
            Err.Raise 1, "Class.DBConneciont.Event.RollbackTransComplete", "Erro ao cancelar transa��o! " & vbCrLf & "Status: " & adStatus
        End If
    End If
End Sub

'Using DBCommandAsync
Private Sub oDB_ExecuteComplete(ByVal RecordsAffected As Long, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    RaiseEvent ExecuteComplete(ByVal RecordsAffected, ByVal pError, adStatus, ByVal pCommand, ByVal pRecordset, ByVal pConnection)
End Sub

'Using DBCommandAsync
Private Sub oDB_WillExecute(Source As String, CursorType As ADODB.CursorTypeEnum, LockType As ADODB.LockTypeEnum, Options As Long, adStatus As ADODB.EventStatusEnum, ByVal pCommand As ADODB.Command, ByVal pRecordset As ADODB.Recordset, ByVal pConnection As ADODB.Connection)
    RaiseEvent WillExecute(Source, CursorType, LockType, Options, adStatus, ByVal pCommand, ByVal pRecordset, ByVal pConnection)
End Sub

Private Sub oDB_InfoMessage(ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pConnection As ADODB.Connection)
    Debug.Print pError.Description
End Sub
