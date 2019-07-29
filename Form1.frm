VERSION 5.00
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Test Connection Factory VB6"
   ClientHeight    =   2505
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4155
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   4155
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtStatus 
      Height          =   1905
      Left            =   1890
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   240
      Width           =   2145
   End
   Begin VB.CommandButton cmdExecuteReaderAsync 
      Caption         =   "ExecuteReaderAsync"
      Height          =   435
      Left            =   30
      TabIndex        =   3
      Top             =   1680
      Width           =   1785
   End
   Begin VB.CommandButton cmdExecuteScalar 
      Caption         =   "ExecuteScalar"
      Height          =   435
      Left            =   30
      TabIndex        =   2
      Top             =   1200
      Width           =   1785
   End
   Begin VB.CommandButton cmdExecuteReader 
      Caption         =   "ExecuteReader"
      Height          =   435
      Left            =   30
      TabIndex        =   1
      Top             =   720
      Width           =   1785
   End
   Begin VB.CommandButton cmdExecuteNonQuery 
      Caption         =   "ExecuteNonQuery"
      Height          =   435
      Left            =   30
      TabIndex        =   0
      Top             =   240
      Width           =   1785
   End
   Begin VB.Label lblProgress 
      Caption         =   "Status (ExecuteReaderAsync)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   60
      TabIndex        =   5
      Top             =   2220
      Width           =   2235
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Private WithEvents commandAsync As DBCommandAsync
Private WithEvents oReaderAsync As DBReaderAsync
Attribute oReaderAsync.VB_VarHelpID = -1
Private conn As DBConnection

Private Sub Form_Load()
    Set conn = gDBPool(DB_ALIAS_DEFAULT)
End Sub

'Example ExecuteNonQuery
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

'Example ExecuteReader
Private Sub cmdExecuteReader_Click()
    Dim oRS As ADODB.Recordset
    Dim cmd As DBCommand

    On Error GoTo cmdExecuteReader_Click_Error

    Set cmd = conn.GetDBCommand()

    Set oRS = cmd.ExecuteReader(adCmdText, "SELECT * FROM SIS_USER")

    Me.txtStatus.Text = ""
    Do While Not oRS.EOF
        Me.txtStatus.Text = Me.txtStatus.Text & FU_Null(oRS(0).Value) & vbCrLf
        oRS.MoveNext
    Loop

    On Error GoTo 0
    Exit Sub

cmdExecuteReader_Click_Error:
    Debug.Assert False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExecuteReader_Click of Formulário Form1"

End Sub

'Example ExecuteScalar
Private Sub cmdExecuteScalar_Click()
    Dim cmd As DBCommand
    On Error GoTo cmdExecuteScalar_Click_Error

    Set cmd = conn.GetDBCommand()

    txtStatus.Text = cmd.ExecuteScalar(adCmdText, "SELECT * FROM SIS_USER")

    Set cmd = Nothing

    On Error GoTo 0
    Exit Sub

cmdExecuteScalar_Click_Error:
    Debug.Assert False
    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdExecuteScalar_Click of Formulário Form1"
End Sub


'Example ExecuteReaderAsync
Private Sub cmdExecuteReaderAsync_Click()
    Dim sSQL As String
    Dim params As New DBParameters
    Set oReaderAsync = gDBPool(DB_ALIAS_SIS).GetDBReaderAsync()

    sSQL = ""
    sSQL = sSQL & "SELECT DISTINCT TOP 10000 SUBSTRING(CP.NumProcesso, 1, 5) + '/' " & vbCrLf
    sSQL = sSQL & "                   + SUBSTRING(CP.NumProcesso, 7, 2), " & vbCrLf
    sSQL = sSQL & "                   C.CodFab, " & vbCrLf
    sSQL = sSQL & "                   C.NumConhec, " & vbCrLf
    sSQL = sSQL & "                   C.DatConhec, " & vbCrLf
    sSQL = sSQL & "                   C.NumDI, " & vbCrLf
    sSQL = sSQL & "                   C.NumConhecMaster, " & vbCrLf
    sSQL = sSQL & "                   SUBSTRING(CP.NumProcesso, 7, 2), " & vbCrLf
    sSQL = sSQL & "                   SUBSTRING(CP.NumProcesso, 1, 5) " & vbCrLf
    sSQL = sSQL & "FROM      tbl_Conhecimento AS C " & vbCrLf
    sSQL = sSQL & "LEFT JOIN tbl_ConhecimentoProcesso AS CP " & vbCrLf
    sSQL = sSQL & "       ON c.NumConhec = cp.NumConhec " & vbCrLf
    sSQL = sSQL & "      AND c.Codfab = cp.CodFab " & vbCrLf
    sSQL = sSQL & "LEFT JOIN tbl_ConhecimentoContainer CC " & vbCrLf
    sSQL = sSQL & "       ON C.NumConhec = CC.NumConhec " & vbCrLf
    sSQL = sSQL & "LEFT JOIN tbl_Fatura F " & vbCrLf
    sSQL = sSQL & "       ON C.NumConhec = F.NumConhec " & vbCrLf
    sSQL = sSQL & "LEFT JOIN tbl_faturaItem FI " & vbCrLf
    sSQL = sSQL & "       ON F.NumFatura = FI.NumFatura " & vbCrLf
    sSQL = sSQL & "ORDER     BY SUBSTRING(CP.NumProcesso, 7, 2) DESC, " & vbCrLf
    sSQL = sSQL & "             SUBSTRING(CP.NumProcesso, 1, 5) DESC "

    'sSQL = "SELECT NAME FROM SIS_PERSON_TYPE WHERE ID = @ID"
    'Call params.Add("@ID", 1, Numeric)
    
    Me.txtStatus.Text = ""
    
    Call oReaderAsync.ExecuteReaderAsync(sSQL, params, 1000)

End Sub

Private Sub oReaderAsync_SQLError(ByVal SQLerr As ErrObject)
    txtStatus.Text = txtStatus.Text & SQLerr.Description & vbCrLf
End Sub

Private Sub oReaderAsync_SQLret(ByVal oRS As ADODB.Recordset, ByVal bSucess As Boolean)
    If bSucess Then
        If oRS.EOF Then
            txtStatus = txtStatus & "Nenhum registro encontrado!"
        Else
            txtStatus = txtStatus & FU_Null(oRS(0).Value) & vbCrLf
            lblProgress.Caption = "Concluido..."
        End If
    Else
        txtStatus = txtStatus & "Erro no retorno" & vbCrLf
    End If

End Sub

Private Sub oReaderAsync_SQLStatus(ByVal sMessage As String, percentage As Integer)
    If percentage = 0 Then
        txtStatus = txtStatus & sMessage & vbCrLf
    Else
        lblProgress.Caption = sMessage & IIf(percentage > 0, " [" & percentage & "%]", "")
        lblProgress.Refresh
    End If
    DoEvents
End Sub
