VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "oSuperCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private pCtr As ComboBox

Private Type tItem
  Caption As String
  Value   As String
End Type

Dim ItemLst() As tItem
Dim ItemNbr   As Long

Public Sub AddItem(Value As String, Caption As String)

  ItemNbr = ItemNbr + 1
  ReDim Preserve ItemLst(0 To (ItemNbr - 1))
  With ItemLst(ItemNbr - 1)
    .Caption = Caption
    .Value = Value
  End With

  If Not (pCtr Is Nothing) Then
    pCtr.AddItem Caption, (ItemNbr - 1)
  End If

End Sub

Public Sub SetCtrObj(Ctr As ComboBox, Optional Clear As Boolean)

  Dim i As Long

  Set pCtr = Ctr
  
  If Clear Then
    ItemNbr = 0
    pCtr.Clear
  Else
    If (pCtr.ListCount = 0) Then
      For i = 0 To (ItemNbr - 1)
        pCtr.AddItem ItemLst(i).Caption, i
      Next i
    End If
  End If
  
End Sub

Public Sub SetCtrVal(Value As String, Optional FirstIfNotFound As Boolean)

  pCtr.ListIndex = Me.GetItemIndex(Value)
  If (pCtr.ListIndex = c_Nothing) And FirstIfNotFound Then
    pCtr.ListIndex = 0
  End If

End Sub

Public Function CurrValue() As String

  If pCtr.ListIndex >= 0 Then
    CurrValue = ItemLst(pCtr.ListIndex).Value
  Else
    CurrValue = vbNullString
  End If

End Function

Public Function GetItemIndex(Value As String) As Long

  Dim Found As Boolean
  Dim i     As Long
  
  i = 0
  Found = False
  Do Until Found Or (i > (ItemNbr - 1))
    If ItemLst(i).Value = Value Then
      Found = True
    Else
      i = i + 1
    End If
  Loop

  If Found Then
    GetItemIndex = i
  Else
    GetItemIndex = c_Nothing
  End If

End Function

Private Sub Class_Initialize()

  ItemNbr = 0

End Sub
