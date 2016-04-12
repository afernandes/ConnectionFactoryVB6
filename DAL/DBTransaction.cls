VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Private DBConn As DBConnection
Private m_bTranOpen As Boolean
Private m_tranLevel As Long

'[BEGIN TRANSACTION]
Friend Function Init(oDBConn As DBConnection) As DBTransaction
    If m_bTranOpen Then
        Err.Raise 1, "Class.DBTransaction.Init", "Transação já inicializada anteriormente."
    Else
        Set DBConn = oDBConn
        m_tranLevel = DBConn.TransactionHandler(enmTransactionOpen)
        m_bTranOpen = True
        Set Init = Me
    End If
End Function

'[COMMIT TRANSACTION]
Public Sub Commit()
    If m_bTranOpen Then
        Call DBConn.TransactionHandler(enmTransactionCommit)
        m_bTranOpen = False
    Else
        Err.Raise vbObjectError + 1, TypeName(Me) & ".Commit()", "There is no open transaction." & vbCrLf & _
                                                                 "Use: DBConnection.GetDBTransaction()"
        Exit Sub

    End If
End Sub

Private Sub Class_Initialize()
    m_bTranOpen = False
End Sub

'[AUTO ROLLBACK TRANSACTION IF NOT COMMIT]
Private Sub Class_Terminate()
    If m_bTranOpen Then
        Call DBConn.TransactionHandler(enmTransactionRollback)
    End If
End Sub
