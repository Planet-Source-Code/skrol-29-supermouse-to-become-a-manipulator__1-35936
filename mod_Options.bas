Attribute VB_Name = "mod_Options"
Option Explicit
Option Compare Text

Public Const c_Opt_Main_StayOnTop = "StayOnTop"
Public Const c_Opt_Main_AutoQuit = "AutoQuit"
Public Const c_Opt_Main_RestorePos = "MainWinPos"
Public Const c_Opt_Mouse_RestorePos = "SaveUserMousePos"
Public Const c_Opt_DateCrea = "CreationDate"
Public Const c_Opt_DateModif = "ModificationDate"
Public Const c_Opt_Version = "Version"

Public OptionLst() As String
Public OptionNbr   As Long

Public Function m_Option_Get(OptType As String, OptName As String, OptPrm As String) As String
'Returns the value of a parameter for the specified option
  
  Dim i     As Long

  i = m_Option_Found(OptType, OptName)

  If i = c_Nothing Then
    m_Option_Get = vbNullString
  Else
    m_Option_Get = m_ItemLst_Get(OptionLst(i), OptPrm)
  End If

End Function

Public Sub m_Option_Set(OptType As String, OptName As String, OptPrm As String, OptValue As String)
'Set the value of a parameter for the specified option

  Dim i     As Long

  i = m_Option_Found(OptType, OptName)

  If i = c_Nothing Then
    'We add an item at the end of the option list
    OptionNbr = OptionNbr + 1
    ReDim Preserve OptionLst(0 To (OptionNbr - 1))
    'We point the the last item to set the value
    i = OptionNbr - 1
    'We save the name of the item
    OptionLst(i) = m_ItemLst_Set(OptionLst(i), c_Prm_Type, OptType)
    If OptName <> vbNullString Then
      OptionLst(i) = m_ItemLst_Set(OptionLst(i), c_Prm_Name, OptName)
    End If
  End If

  OptionLst(i) = m_ItemLst_Set(OptionLst(i), OptPrm, OptValue)
  
End Sub

Public Function m_Option_Found(OptType As String, OptName As String) As Long
'Search the index of a specific option
'Returns -1 if nothing is found
  
  Dim i     As Long
  Dim Found As Boolean

  'We search for the item that has the specified option name
  i = 0
  Found = False
  Do Until Found Or (i > OptionNbr - 1)
    If m_ItemLst_Get(OptionLst(i), c_Prm_Type) = OptType Then
      If OptName = vbNullString Then
        Found = True
      Else
        If m_ItemLst_Get(OptionLst(i), c_Prm_Name) = OptName Then
          Found = True
        End If
      End If
    Else
      i = i + 1
    End If
  Loop

  If Found Then
    m_Option_Found = i
  Else
    m_Option_Found = c_Nothing
  End If

End Function
