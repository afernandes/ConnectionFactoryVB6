VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Connection Factory VB6"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   6225
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExecuteScalar 
      Caption         =   "ExecuteScalar"
      Height          =   435
      Left            =   30
      TabIndex        =   3
      Top             =   1200
      Width           =   1485
   End
   Begin VB.TextBox Text1 
      Height          =   1515
      Left            =   1620
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   210
      Width           =   4545
   End
   Begin VB.CommandButton cmdExecuteReader 
      Caption         =   "ExecuteReader"
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   720
      Width           =   1485
   End
   Begin VB.CommandButton cmdExecuteNonQuery 
      Caption         =   "ExecuteNonQuery"
      Height          =   435
      Left            =   30
      TabIndex        =   0
      Top             =   240
      Width           =   1485
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private conn As DBConnection

Private Sub Form_Load()
    Set conn = gDBPool(DB_ALIAS_DEFAULT)
End Sub

Private Sub cmdExecuteNonQuery_Click()
    Dim cmd As DBCommand
    Dim tran As DBTransaction
    Dim params As New DBParameters
    Dim ret As Boolean

    On Error GoTo cmdExecuteNonQuery_Click_Error
    Screen.MousePointer = vbHourglass

    Set tran = conn.GetDBTransaction()    '<-- Begin Transaction
    Set cmd = conn.GetDBCommand()

    params.Add "@colString", "insert with transaction", Character
    params.Add "@colNumeric", 10.12, Numeric
    params.Add "@colDate", "09/18/2015", Date

    ret = cmd.ExecuteNonQuery(adCmdStoredProc, "[dbo].spc_TesteConnectionFactory", params)

    tran.Commit    '<-- Commit Transaction

    If ret Then
        MsgBox "Sucesso!", vbInformation
    Else
        MsgBox "Falhou!", vbExclamation
    End If

    Screen.MousePointer = vbDefault
    On Error GoTo 0
    Exit Sub

cmdExecuteNonQuery_Click_Error:        '<-- Auto Roolback Transaction
    Screen.MousePointer = vbDefault
    Debug.Print "Error:        " & Err.Number & vbCrLf & _
                "Description:  " & Err.Description & vbCrLf & _
                "Source:       " & Err.Source & vbCrLf & _
                "LastDllError: " & Err.LastDllError & vbCrLf & _
                "Trace:        Form1->cmdExecuteNonQuery_Click"
    Debug.Assert False
    MsgBox Err.Description, vbExclamation, "Atenção!"
End Sub

Private Sub cmdExecuteReader_Click()
    Dim oRS As ADODB.Recordset
    Dim cmd As DBCommand

   On Error GoTo cmdExecuteReader_Click_Error

    Set cmd = conn.GetDBCommand()

    Set oRS = cmd.ExecuteReader(adCmdText, "SELECT * FROM SIS_USER")

    Me.Text1.Text = ""
    Do While Not oRS.EOF
        Me.Text1.Text = Me.Text1.Text & FU_Null(oRS(0).Value) & vbCrLf
        oRS.MoveNext
    Loop

   On Error GoTo 0
   Exit Sub

cmdExecuteReader_Click_Error:
    Debug.Assert False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExecuteReader_Click of Formulário Form1"

End Sub

Private Sub cmdExecuteScalar_Click()
    Dim cmd As DBCommand
   On Error GoTo cmdExecuteScalar_Click_Error

    Set cmd = conn.GetDBCommand()

    Text1.Text = cmd.ExecuteScalar(adCmdText, "SELECT * FROM SIS_USER")

    Set cmd = Nothing

   On Error GoTo 0
   Exit Sub

cmdExecuteScalar_Click_Error:
    Debug.Assert False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExecuteScalar_Click of Formulário Form1"
End Sub
