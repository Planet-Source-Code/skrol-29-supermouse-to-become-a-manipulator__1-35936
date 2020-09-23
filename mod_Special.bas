Attribute VB_Name = "mod_Special"
Option Explicit
Option Compare Text

Public g_Pos_Frm       As Form
Public g_Pos_SaveX     As String
Public g_Pos_SaveY     As String
Public g_Pos_SaveColor As String
Public g_Pos_Ok        As Boolean
Public g_Pos_EditXY    As Boolean
Public g_Pos_EditColor As Boolean

Public Sub m_Pos_LocatePoint(Frm As Form, EditXY As Boolean, EditColor As Boolean)

  Frm.lbl_LocInfo.Visible = True
  
  Set g_Pos_Frm = Frm
  g_Pos_EditXY = EditXY
  g_Pos_EditColor = EditColor

  frm_Target_Point.Show vbModal

  Frm.lbl_LocInfo.Visible = False

End Sub

Public Function m_HexToDec(HexVal As String) As Long

  Dim i    As Integer
  Dim d    As Long
  Dim X    As String
  
  X = "000000" & Trim$(HexVal)
  d = 0
  
  For i = 0 To 5
    Select Case Mid$(X, Len(X) - i, 1)
    Case "1": d = d + 1 * (16 ^ i)
    Case "2": d = d + 2 * (16 ^ i)
    Case "3": d = d + 3 * (16 ^ i)
    Case "4": d = d + 4 * (16 ^ i)
    Case "5": d = d + 5 * (16 ^ i)
    Case "6": d = d + 6 * (16 ^ i)
    Case "7": d = d + 7 * (16 ^ i)
    Case "8": d = d + 8 * (16 ^ i)
    Case "9": d = d + 9 * (16 ^ i)
    Case "A": d = d + 10 * (16 ^ i)
    Case "B": d = d + 11 * (16 ^ i)
    Case "C": d = d + 12 * (16 ^ i)
    Case "D": d = d + 13 * (16 ^ i)
    Case "E": d = d + 14 * (16 ^ i)
    Case "F": d = d + 15 * (16 ^ i)
    End Select
  Next i

  m_HexToDec = d

End Function

Public Sub m_RelativeInfo_Get(ByRef Relative As String, ByRef ActiveWinHwnd As Long, ByRef Mediane As Boolean)

  Select Case Relative
  Case c_Relative_Screen
    ActiveWinHwnd = 0
    Mediane = False
  Case c_Relative_ScreenM
    ActiveWinHwnd = 0
    Mediane = True
  Case c_Relative_ActiveWindow
    ActiveWinHwnd = m_Win_GetNextActive(0, g_ForbiddenWin)
    Mediane = False
  Case c_Relative_ActiveWindowM
    ActiveWinHwnd = m_Win_GetNextActive(0, g_ForbiddenWin)
    Mediane = True
  Case Else
    ActiveWinHwnd = 0
    Mediane = False
  End Select

End Sub

Public Sub m_RelativeInfo_Actualize(Frm As Form, scRelative As oSuperCombo)

  Dim ActiveWinHwnd As Long
  Dim Mediane       As Boolean

  With Frm
    m_RelativeInfo_Get scRelative.CurrValue, ActiveWinHwnd, Mediane
    .ActiveWinHwnd = ActiveWinHwnd
    .Mediane = Mediane
    If ActiveWinHwnd = 0 Then
      .lbl_ActiveWin = vbNullString
    Else
      .lbl_ActiveWin = m_Win_GetTitle(ActiveWinHwnd)
    End If
  End With
  
End Sub

Public Sub m_RelativeInfo_FeedLst(scRelative As oSuperCombo, LstCtr As ComboBox)

  With scRelative
    .SetCtrObj LstCtr, True
    .AddItem c_Relative_Screen, "Screen"
    .AddItem c_Relative_ScreenM, "Screen Mediane"
    .AddItem c_Relative_ActiveWindow, "Active window"
    .AddItem c_Relative_ActiveWindowM, "Active window Mediane"
  End With

End Sub


