Attribute VB_Name = "MMain"
Option Explicit

Public Const DB_ALIAS_DEFAULT = "DEFAULT"
Public Const DB_ALIAS_SIS = "SIS"
Public Const DB_ALIAS_SIM = "SIM"

'[CONNECTION POOL]
Public gDBPool As DBPool

Sub Main()
    Dim frm As frmTest

    Set gDBPool = New DBPool

    'MAP CONNECTION
    ' Call gDBPool.MapConnection(DB_ALIAS_DEFAULT, "DRIVER=Firebird/InterBase(r) driver;UID=SYSDBA;PWD=masterkey;DBNAME=C:\Users\ander\Dropbox\Firebird DB\db1.fdb;")
    'Call gDBPool.MapConnection(DB_ALIAS_SIS, "Provider=SQLNCLI11;Server=(localdb)\.\MSSQLLocalDBShare;Uid=stilo;Pwd=stistilo;")

    Call gDBPool.MapConnection(DB_ALIAS_SIS, "Provider=SQLOLEDB;" & _
                                             "Server=CJRVM12\HOMOLOG;" & _
                                             "Initial Catalog=dtb_SIM;" & _
                                             "Trusted_Connection=yes;" & _
                                             "App=ConnectionFactoryVB6;")

    Set frm = New frmTest
    Load frm
    frm.Show 1
End Sub


'Public Function getConnectionString_MsSqlSSPI(Optional ByVal sServer As String = "CJRVM12\Homolog", Optional ByVal sCatalog As String = "dtb_SIM") As String
'    getConnStr_MsSqlSSPI = "Provider=SQLOLEDB;" & _
     '                           "Integrated Security=SSPI;" & _
     '                           "Initial Catalog=" & sCatalog & ";" & _
     '                           "App=" & App.Title & ";" & _
     '                           "Data Source=" & sServer
'End Function

