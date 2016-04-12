VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBCommand"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private DBConn As DBConnection
Private bClassInit As Boolean

Private Sub Class_Initialize()
    bClassInit = False
End Sub
Private Sub Class_Terminate()
    bClassInit = False
End Sub

Friend Function Init(oDBConn As DBConnection) As DBCommand
    Set DBConn = oDBConn
    bClassInit = True
    Set Init = Me
End Function

Public Function ExecuteNonQuery(cmdType As CommandTypeEnum, ByVal cmdText As String, Optional ByVal oParams As DBParameters = Nothing, Optional RecordsAffected As Long = 0) As Boolean
'#AndersonN - 28/11/2014
    Dim ocmd As New ADODB.Command

    On Error GoTo ExecuteNonQuery_Error

    ExecuteNonQuery = False
    
    If Not bClassInit Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".ExecuteNonQuery()", "Class uninitialized." & vbCrLf & _
                                                                          "Use: DBConnection.GetDBCommand()"
        Exit Function
    End If
    
    Set ocmd = PrepareCommand(cmdType, cmdText, oParams)
    If RecordsAffected <= 0 Then
        ocmd.Execute
    Else
        ocmd.Execute RecordsAffected
    End If
    ExecuteNonQuery = True

    On Error GoTo 0
    Exit Function

ExecuteNonQuery_Error:
    Debug.Print "Error: &h" & Hex$(Err.Number) & " :: " & Err.Description & vbCrLf & _
                "--Trace: " & Err.Source & "<-ExecuteNonQuery<-DBCommand"

    Debug.Assert False
    Err.Raise Err.Number, "Class.DBCommand.ExecuteNonQuery", Err.Description

    ExecuteNonQuery = False

End Function

Public Function ExecuteReader(cmdType As CommandTypeEnum, ByVal cmdText As String, Optional ByVal oParams As DBParameters = Nothing, Optional RecordsAffected As Long = 0) As Recordset
    Dim ocmd As New ADODB.Command

    On Error GoTo ExecuteReader_Error

    If Not bClassInit Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".ExecuteNonQuery()", "Class uninitialized." & vbCrLf & _
                                                                          "Use: DBConnection.GetDBCommand()"
        Exit Function
    End If

    Set ocmd = PrepareCommand(cmdType, cmdText, oParams)

    If RecordsAffected > 0 Then
        Set ExecuteReader = ocmd.Execute(RecordsAffected)
    Else
        Set ExecuteReader = ocmd.Execute()
    End If

    On Error GoTo 0
    Exit Function

ExecuteReader_Error:
    Debug.Print "Error: &h" & Hex$(Err.Number) & " :: " & Err.Description & vbCrLf & _
                "--Trace: " & Err.Source & "<-ExecuteReader<-DBCommand"

    Debug.Assert False
    Err.Raise Err.Number, "Class.DBCommand.ExecuteReader", Err.Description

End Function

Public Function ExecuteScalar(cmdType As CommandTypeEnum, ByVal cmdText As String, Optional ByVal oParams As DBParameters = Nothing) As Variant
    Dim ocmd As New ADODB.Command
    Dim oRS As New ADODB.Recordset

    On Error GoTo ExecuteScalar_Error
    
    If Not bClassInit Then
        Err.Raise vbObjectError + 1, TypeName(Me) & ".ExecuteNonQuery()", "Class uninitialized." & vbCrLf & _
                                                                          "Use: DBConnection.GetDBCommand()"
        Exit Function
    End If

    Set ocmd = PrepareCommand(cmdType, cmdText, oParams)
    Set oRS = ocmd.Execute(1)
    If oRS.EOF Then
        ExecuteScalar = Empty
    Else
        ExecuteScalar = IIf(IsNull(oRS(0).Value), Empty, oRS(0).Value)
    End If

    On Error GoTo 0
    Exit Function

ExecuteScalar_Error:
    Debug.Print "Error: &h" & Hex$(Err.Number) & " :: " & Err.Description & vbCrLf & _
                "--Trace: " & Err.Source & "<-ExecuteScalar<-DBCommand"

    Debug.Assert False
    Err.Raise Err.Number, "Class.DBCommand.ExecuteScalar", Err.Description


End Function

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
                "--Trace: " & Err.Source & "<-PrepareCommand<-DBCommand"

    Debug.Assert False
    Err.Raise Err.Number, "Class.DBCommand.PrepareCommand", Err.Description

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
