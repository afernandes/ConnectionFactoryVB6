VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DBParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Lista de parametros"
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"DBParameter"
Attribute VB_Ext_KEY = "Member0" ,"DBParameter"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mCol As Collection

Public Function Add(Name As String, Value As String, Optional TypeParam As enmTypeParameter) As DBParameter
   Dim objNewMember As DBParameter

   If Len(Name) = 0 Then
      Err.Raise 1, "Class.DBParameter.Add", "Nome do parametro não pode ser nulo"
      Exit Function
   End If

   Set objNewMember = New DBParameter

   With objNewMember
      .Name = Name
      .Value = Value
      .TypeParam = TypeParam
   End With

   mCol.Add objNewMember, Name

   Set Add = objNewMember
   Set objNewMember = Nothing

End Function

Public Property Get Item(vntIndexKey As Variant) As DBParameter
Attribute Item.VB_UserMemId = 0
'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
   Set Item = mCol(vntIndexKey)
End Property

Public Property Get Count() As Long
   Count = mCol.Count
End Property

Public Sub Remove(vntIndexKey As Variant)
   mCol.Remove vntIndexKey
End Sub

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
'this property allows you to enumerate
'this collection with the For...Each syntax
   Set NewEnum = mCol.[_NewEnum]
End Property

Private Sub Class_Initialize()
   Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
   Set mCol = Nothing
End Sub

