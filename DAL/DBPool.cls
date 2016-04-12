VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBPool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private m_colConn As Collection

Private Sub Class_Initialize()
   Set m_colConn = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_colConn = Nothing
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = m_colConn.[_NewEnum]
End Property

Public Function MapConnection(Alias As String, ConnectionString As String) As DBConnection
    'create a new object
    Dim objNewConn As DBConnection
    
    If Len(Trim$(Alias)) = 0 Then
        Err.Raise vbObjectError + 1, "DBPool.MapConnection", "Alias da conexão obrigatório!"
        Exit Function
    Else
        If Exists(Alias) Then
            Err.Raise vbObjectError + 1, "DBPool.MapConnection", "Já existe uma conexão mapeada para este alias!"
            Exit Function
        End If
    End If
        
    Set objNewConn = New DBConnection

    'set the properties passed into the method
    objNewConn.Alias = Alias
    
    If objNewConn.Init(ConnectionString) Then
        m_colConn.Add objNewConn, Alias
    End If

    'return the object created
    Set MapConnection = objNewConn
    Set objNewConn = Nothing
End Function

Public Function Exists(Alias As String) As Boolean
    On Error GoTo Exists_Error
    
    Exists = False
    
    If Not m_colConn.Item(Alias) Is Nothing Then
        Exists = True
    End If
    
    Exit Function
Exists_Error:
    On Error GoTo 0

End Function

Public Function GetConn(Alias As String) As DBConnection
Attribute GetConn.VB_UserMemId = 0
    On Error GoTo Exists_Error
    
    Set GetConn = m_colConn(Alias)

    Exit Function
Exists_Error:
    On Error GoTo 0
End Function


Public Sub Remove(Alias As String)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)

    m_colConn.Remove Alias
End Sub



'Public Function getConnection(ConnectionString As String) As ADODB.Connection
'   Dim tCD        As DBConnectionData
'   Dim returnValue As ADODB.Connection
'   Dim conStr     As String
'   Dim iteration  As Long
'   Dim bClear     As Boolean
'   Dim bFound     As Boolean
'
'   bClear = False
'   bFound = False
'
'   conStr = connectionStringFIX(ConnectionString)
'
'   For iteration = 1 To m_colConn.Count
'      Set tCD = m_colConn.Item(iteration)
'      If tCD Is Nothing Then
'         bClear = True
'      Else
'         If tCD.Connection Is Nothing Then
'            bClear = True
'         Else
'            If conStr = tCD.ConnectionString Then
'               If tCD.Connection.State = adStateClosed Then
'                  tCD.Connection.Open
'               End If
'               bFound = True
'               tCD.UpdateTime = Now
'               Call m_colConn.Remove(iteration)
'               Call m_colConn.Add(tCD)
'               Set returnValue = tCD.Connection
'               Exit For
'            End If
'         End If
'      End If
'   Next
'
'   If Not bFound Then
'      Set tCD = New DBConnectionData
'      Set tCD.Connection = New ADODB.Connection
'      tCD.Connection.Open conStr
'      tCD.ConnectionString = conStr
'      tCD.UpdateTime = Now
'      m_colConn.Add tCD, conStr
'      Set returnValue = tCD.Connection
'   End If
'
'   If bClear Then
'      Call releaseFreeConnections
'   End If
'
'   Set getConnection = returnValue
'End Function

'Private Sub releaseFreeConnections()
'   Dim iteration  As Long
'   Dim tCD        As DBConnectionData
'
'   For iteration = 1 To m_colConn.Count
'      tCD = m_colConn(iteration)
'      If tCD Is Nothing Then
'         m_colConn.Remove iteration
'         iteration = iteration - 1
'      ElseIf tCD.Connection Is Nothing Then
'         m_colConn.Remove iteration
'         iteration = iteration - 1
'      End If
'   Next
'End Sub

'Private Function connectionStringFIX(ConnectionString As String) As String
'   Dim returnValue As String
'   Dim elementsFix() As String
'   Dim elements() As String
'   Dim elementAux As String
'   Dim properties() As String
'   Dim i, y       As Integer
'
'   elements = Split(Trim$(ConnectionString), ";")
'   ReDim elementsFix(UBound(elements))
'
'   For i = LBound(elements) To UBound(elements)
'      properties = Split(Trim$(elements(i)), "=", 2)
'      For y = LBound(properties) To UBound(properties)
'         properties(y) = Trim$(properties(y))
'      Next
'      elementsFix(i) = Join(properties, "=")
'   Next
'
'   For i = LBound(elementsFix) To UBound(elementsFix) - 1
'      For y = LBound(elementsFix) To UBound(elementsFix) - 1
'         If UCase(elementsFix(y)) > UCase(elementsFix(y + 1)) Then
'            elementAux = elementsFix(y + 1)
'            elementsFix(y + 1) = elementsFix(y)
'            elementsFix(y) = elementAux
'         End If
'      Next
'   Next
'
'   returnValue = Join(elementsFix, ";")
'   Do While Left$(returnValue, 1) = ";"
'      returnValue = Mid$(returnValue, 2)
'   Loop
'
'   connectionStringFIX = returnValue
'
'End Function

