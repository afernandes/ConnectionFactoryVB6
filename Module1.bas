Attribute VB_Name = "Util"
Option Explicit

Public Const DB_ALIAS_DEFAULT = "DEFAULT"

'[CONNECTION POOL]
Public gDBPool As DBPool

Sub Main()
    Dim frm As frmTest

    Set gDBPool = New DBPool

    'MAP CONNECTION
    Call gDBPool.MapConnection(DB_ALIAS_DEFAULT, "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=C:\Users\ander\Dropbox\Firebird DB\db1.fdb;")
    
    Set frm = New frmTest
    Load frm
    frm.Show
End Sub


Function FU_Asp(ByVal sCampo As String, Optional pNull As Boolean) As String
    If pNull And Trim$(sCampo) = "" Then
        FU_Asp = "NULL"
        Exit Function
    End If

    FU_Asp = Chr$(39) & Replace(sCampo, Chr$(39), Chr$(39) & Chr(39)) & Chr$(39)
End Function

Function FU_Null(ByVal Campo, Optional pNumeric As Boolean) As Variant
    If IsNull(Campo) Then
        FU_Null = IIf(pNumeric, 0, Empty)
    Else
        FU_Null = IIf(pNumeric, Campo, Trim$(Campo))
    End If
End Function

Function FU_Date(ByVal p_Data As Variant) As Variant

    If IsNull(p_Data) Then
        FU_Date = "NULL"
    ElseIf Trim$(p_Data) = "/  /" Or Trim$(p_Data) = "" Then
        FU_Date = "NULL"
    Else
        FU_Date = Chr$(39) & Format$(p_Data, "mm/dd/yyyy") & Chr$(39)
    End If

End Function

'Public Function getConnectionString_MsSqlSSPI(Optional ByVal sServer As String = "CJRVM12\Homolog", Optional ByVal sCatalog As String = "dtb_SIM") As String
'    getConnStr_MsSqlSSPI = "Provider=SQLOLEDB;" & _
'                           "Integrated Security=SSPI;" & _
'                           "Initial Catalog=" & sCatalog & ";" & _
'                           "App=" & App.Title & ";" & _
'                           "Data Source=" & sServer
'End Function

