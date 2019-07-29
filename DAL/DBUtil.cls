VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBUtil"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function FU_Asp(ByVal sCampo As String, Optional pNull As Boolean) As String
    If pNull And Trim$(sCampo) = "" Then
        FU_Asp = "NULL"
        Exit Function
    End If

    FU_Asp = Chr$(39) & Replace(sCampo, Chr$(39), Chr$(39) & Chr(39)) & Chr$(39)
End Function

Public Function FU_Date(ByVal p_Data As Variant) As Variant

    If IsNull(p_Data) Then
        FU_Date = "NULL"
    ElseIf Trim$(p_Data) = "/  /" Or Trim$(p_Data) = "" Then
        FU_Date = "NULL"
    Else
        FU_Date = Chr$(39) & Format$(p_Data, "yyyy-mm-dd") & Chr$(39)
    End If

End Function

Public Function FU_Null(ByVal Campo, Optional pNumeric As Boolean) As Variant
    If IsNull(Campo) Then
        FU_Null = IIf(pNumeric, 0, Empty)
    Else
        FU_Null = IIf(pNumeric, Campo, Trim$(Campo))
    End If
End Function


