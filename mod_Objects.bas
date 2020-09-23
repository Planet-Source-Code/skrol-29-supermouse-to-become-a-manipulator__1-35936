Attribute VB_Name = "mod_Actions"
Option Explicit
Option Compare Text

Public g_AskForStop   As Boolean
Public g_StartTime    As Date
Public g_RestorePos   As Boolean
Public g_StayOnTop    As Boolean
Public g_ForbiddenWin As Long

Public g_CurrAct_Prm As String
Public g_CurrAct_Id  As Long
Public g_CurrAct_Add As Boolean
Public g_CurrAct_Ok  As Boolean

Public g_LastGoto_Id As Long

Public Const c_Nothing = -1

Public Const c_Txt_New = "new"

Public Const c_Prm_Type = "Type"
Public Const c_Prm_Name = "Name"
Public Const c_Prm_Enabled = "Enabled"
Public Const c_Prm_X = "X"
Public Const c_Prm_Y = "Y"
Public Const c_Prm_Relative = "RelativeTo"
Public Const c_Prm_FrmX = "WinX"
Public Const c_Prm_FrmY = "WinY"
Public Const c_Prm_FrmW = "WinW" 'width
Public Const c_Prm_FrmH = "WinH" 'height
Public Const c_Prm_Speed = "Speed"
Public Const c_Prm_Radius = "Radius"
Public Const c_Prm_Move = "Move"
Public Const c_Prm_Button1 = "Button1"
Public Const c_Prm_Button2 = "Button2"
Public Const c_Prm_Nbr = "Nbr"
Public Const c_Prm_Unit = "Unit"
Public Const c_Prm_Value = "Val"
Public Const c_Prm_Color = "Color"
Public Const c_Prm_Mode = "Mode"
Public Const c_Prm_Then = "Then"
Public Const c_Prm_Else = "Else"

Public Const c_Act_Loop = "Loop"
Public Const c_Act_Wait = "Wait"
Public Const c_Act_Mouse = "Mouse"
Public Const c_Act_Keys = "Keys"
Public Const c_Act_Execute = "Execute"
Public Const c_Act_Comment = "Comment"
Public Const c_Act_Label = "Label"
Public Const c_Act_Goto = "Goto"
Public Const c_Act_Condition = "Condition"
Public Const c_Act_Message = "Message"

Public Const c_Val_Yes = "Yes"
Public Const c_Val_No = "No"
Public Const c_Val_Return = "Return"
Public Const c_Val_End = "End"

Public Const c_Relative_Screen = "Screen"
Public Const c_Relative_ScreenM = "ScreenMediane"
Public Const c_Relative_ActiveWindow = "ActiveWindow"
Public Const c_Relative_ActiveWindowM = "ActiveWindowMediane"

Public Const c_Unit_Occ = "nbr"
Public Const c_Unit_Ms = "ms"
Public Const c_Unit_Sec = "sec"
Public Const c_Unit_Min = "min"
Public Const c_Unit_Hour = "hr"

Public Const c_Move_None = "None"
Public Const c_Move_Line = "Line"
Public Const c_Move_Circle = "Circle"

Public Const c_Button_Up = "U"
Public Const c_Button_Down = "D"
Public Const c_Button_Wait = "W"
Public Const c_Button_Click = c_Button_Down & c_Button_Up
Public Const c_Button_DblClick = c_Button_Click & c_Button_Click

Public Const c_Todo_Skip = "skip"
Public Const c_Todo_Goto = "goto"
Public Const c_Todo_Next = "next"

Public Type Action
  Type      As String
  Enabled   As Boolean
  Prameters As String
  StartTime As Long
  Nbr       As Long
  LoopLevel As Integer
  LoopBeg   As Long
  LoopEnd   As Long
End Type

Public ActionLst() As Action '(0 To ActionNbr - 1)
Public ActionNbr   As Long

Public Type mLabel
  Name  As String
  ActId As Long
End Type

Public LabelLst() As mLabel
Public LabelNbr   As Long

Public ErrTxt      As String
Public ErrActId    As Long

Public Sub m_Err_Init()

  ErrTxt = ""
  ErrActId = c_Nothing
  
End Sub

Public Function m_Action_IsLoopE(ActId As Long) As Boolean
'Return True if the action is a loop-end action
  With ActionLst(ActId)
    If .Type = c_Act_Loop Then
      m_Action_IsLoopE = (ActId = .LoopEnd)
    End If
  End With
End Function

Public Function m_Action_IsLoopB(ActId As Long) As Boolean
'Return True if the action is a loop-begin action
  With ActionLst(ActId)
    If .Type = c_Act_Loop Then
      m_Action_IsLoopB = (ActId = .LoopBeg)
    End If
  End With
End Function

Public Function m_Action_Caption(ActId As Long) As String

  Dim X       As String
  Dim s       As String
  Dim Name    As Boolean

  With ActionLst(ActId)
  
    Select Case .Type
    Case c_Act_Loop
      If m_Action_IsLoopB(ActId) Then
        X = "{ (" & m_ItemLst_Get(.Prameters, c_Prm_Nbr) & m_ItemLst_Get(.Prameters, c_Prm_Unit) & ") "
        Name = True
      Else
        X = "}"
      End If
    Case c_Act_Mouse
      s = m_ItemLst_Get(.Prameters, c_Prm_Speed)
      If Val(s) > 0 Then
        s = "(s" & s & ") "
      Else
        s = ""
      End If
      X = m_ItemLst_Get(.Prameters, c_Prm_Button1) & m_ItemLst_Get(.Prameters, c_Prm_Move) & m_ItemLst_Get(.Prameters, c_Prm_Button2)
      X = Replace(X, c_Move_None, "")
      X = Replace(X, c_Move_Line, "->")
      X = Replace(X, c_Move_Circle, "-O")
      X = Replace(X, c_Button_Click, "[.]")
      X = Replace(X, c_Button_Down, "[")
      X = Replace(X, c_Button_Up, "]")
      X = s & X & " "
      Name = True
    Case c_Act_Keys
      X = "xyz - "
      Name = True
    Case c_Act_Execute
      X = "exec - "
      Name = True
    Case c_Act_Comment
      X = "// " & m_ItemLst_Get(.Prameters, c_Prm_Value)
    Case c_Act_Wait
      X = "... (" & m_ItemLst_Get(.Prameters, c_Prm_Nbr) & m_ItemLst_Get(.Prameters, c_Prm_Unit) & ") - "
      Name = True
    Case c_Act_Goto
      X = "=> " & m_ItemLst_Get(.Prameters, c_Prm_Value)
    Case c_Act_Label
      X = ":" & m_ItemLst_Get(.Prameters, c_Prm_Value) & ":"
    Case c_Act_Condition
      X = "(~) " & m_ItemLst_Get(.Prameters, c_Prm_Name)
    Case c_Act_Message
      X = "(!) " & m_ItemLst_Get(.Prameters, c_Prm_Value)
    Case Else
      X = "??? " & .Type
      Name = True
    End Select
  
    If Name Then
      X = X & m_ItemLst_Get(.Prameters, c_Prm_Name)
    End If
    
    If Not .Enabled Then
      X = "# " & X
    End If
  
  End With

  m_Action_Caption = X

End Function

Public Function m_DoAct_Any(ActId As Long) As Long
'This function runs

  Dim NextId  As Long
  
  With ActionLst(ActId)
  
    If m_Action_IsLoopB(ActId) Then
      'It's a loop-begin
      If .Enabled Then
        NextId = m_DoAct_LoopBegin(ActId)
      Else
        NextId = .LoopEnd + 1
      End If
    ElseIf m_Action_IsLoopE(ActId) Then
        'It's a loop-end
        NextId = .LoopBeg
    Else
    
      NextId = ActId + 1
      
      If .Enabled Then
        Select Case .Type
        Case c_Act_Wait
          m_DoAct_Wait (ActId)
        Case c_Act_Mouse
          m_DoAct_Mouse ActId
        Case c_Act_Keys
          m_DoAct_Keys ActId
        Case c_Act_Execute
          m_DoAct_Execute ActId
        Case c_Act_Goto
          NextId = m_DoAct_Goto(ActId)
        Case c_Act_Condition
          NextId = m_DoAct_Condition(ActId)
        Case c_Act_Message
          m_DoAct_Message ActId
        Case c_Act_Comment, c_Act_Label
        Case Else
          ErrActId = ActId
          ErrTxt = "The action type '" & .Type & "' is unknown."
        End Select
      End If
      
    End If
  
  End With

  m_DoAct_Any = NextId

End Function

Public Function m_DoAct_LoopBegin(ActId As Long) As Long

  Dim NextId As Long
  
  Dim IsOcc    As Boolean
  Dim Duration As Long
  Dim Continue As Boolean
  Dim X        As String
  Dim z        As String
  
  NextId = NextId + 1
  
  With ActionLst(ActId)
    
    If .Nbr = 0 Then
      .StartTime = api_GetTickCount()
    End If
    .Nbr = .Nbr + 1
    
    Duration = m_Action_Duration(ActId, True)
        
    If ErrActId = c_Nothing Then
      
      If Duration <= 0 Then
        IsOcc = True
        Duration = Abs(Duration)
      Else
        IsOcc = False
      End If
      
      'See if the loop continue
      If IsOcc Then
        Continue = (.Nbr <= Duration)
      Else
        Continue = (.StartTime + Duration >= api_GetTickCount())
      End If
    
      'Decide wich is next Action
      If Continue Then
      
        NextId = ActId + 1
        
        'show the count int the main list
        X = m_Action_Caption(.LoopEnd)
        z = Format$(TimeSerial(0, 0, (api_GetTickCount() - .StartTime) / 1000), "hh:nn:ss")
        X = X & " nbr=" & .Nbr & " , durée=" & z
        frm_Main.lst_Action.List(.LoopEnd) = X
      
      Else
      
        .Nbr = 0
        .StartTime = 0
        NextId = .LoopEnd + 1
        
      End If
    
    End If

  End With

  m_DoAct_LoopBegin = NextId

End Function

Public Sub m_DoAct_Message(ActId As Long)

  Dim Txt As String

  Txt = m_ItemLst_Get(ActionLst(ActId).Prameters, c_Prm_Value)

  MsgBox Txt

End Sub

Public Sub m_DoAct_Wait(ActId As Long)

  Dim Ticks0 As Long
  Dim NbrMilliSec As Long
  
  Dim PosX   As Long
  Dim PosY   As Long
  Dim ColHex As String
  Dim Color  As Long
  
  Dim WaitColor As Boolean
  Dim WaitTime  As Boolean
  Dim Ok        As Boolean
  
  'Get action's parametres in variables
  NbrMilliSec = m_Action_Duration(ActId)
  With ActionLst(ActId)
    ColHex = m_ItemLst_Get(.Prameters, c_Prm_Color)
    Color = m_HexToDec(ColHex)
  End With
  
  'Initialize other variables
  WaitColor = (Trim$(ColHex) <> vbNullString)
  WaitTime = (NbrMilliSec > 0)
  Ticks0 = api_GetTickCount()
  Ok = False
  
  'Start the waiting
  Do Until Ok
  
    If WaitTime Then
      If api_GetTickCount() >= Ticks0 + NbrMilliSec Then
        Ok = True
      End If
    End If
    
    If WaitColor Then
      m_Action_GetPos ActionLst(ActId).Prameters, PosX, PosY
      If m_Pixel_GetColor(PosX, PosY) = Color Then
        Ok = True
      End If
    End If
    
    If g_AskForStop Then
      Ok = True
    End If
    
    DoEvents
    
  Loop
  
End Sub

Public Sub m_DoAct_Execute(ActId As Long)

  Dim Prm      As String
  Dim WinStyle As VbAppWinStyle

  On Error GoTo Shell_Err

  With ActionLst(ActId)
    Prm = m_ItemLst_Get(.Prameters, c_Prm_Mode)
    If Prm = "" Then
      WinStyle = vbNormalFocus
    Else
      WinStyle = Val(Prm)
    End If
    Shell m_ItemLst_Get(.Prameters, c_Prm_Value), WinStyle
  End With

Shell_End:
  Exit Sub

Shell_Err:
  ErrActId = ActId
  ErrTxt = Err.Description
  Resume Shell_End

End Sub

Public Sub m_DoAct_Keys(ActId As Long)

  On Error GoTo Keys_Err
  
  With ActionLst(ActId)
    SendKeys m_ItemLst_Get(.Prameters, c_Prm_Value), True
  End With

Keys_End:
  Exit Sub

Keys_Err:
  ErrTxt = Err.Description
  ErrActId = ActId
  Resume Keys_End

End Sub

Public Sub m_DoAct_Mouse(ActId As Long)

  Dim X As Long
  Dim Y As Long
  Dim Speed As Long
  Dim Radius As Long
  Dim Button1 As String
  Dim Button2 As String
  Dim Move    As String

  With ActionLst(ActId)
    
    m_Action_GetPos .Prameters, X, Y
    
    Speed = Val(m_ItemLst_Get(.Prameters, c_Prm_Speed))
    Radius = Val(m_ItemLst_Get(.Prameters, c_Prm_Radius))
    Move = m_ItemLst_Get(.Prameters, c_Prm_Move)
    Button1 = m_ItemLst_Get(.Prameters, c_Prm_Button1)
    Button2 = m_ItemLst_Get(.Prameters, c_Prm_Button2)
    
    'Add a pause between double-clicks
    Button1 = Replace(Button1, c_Button_DblClick, c_Button_Click & c_Button_Wait & c_Button_Click)
    Button2 = Replace(Button2, c_Button_DblClick, c_Button_Click & c_Button_Wait & c_Button_Click)
    
    'Performs mouse actions
    m_Mouse_Action Move, X, Y, Button1, Button2, Speed, Radius
  
  End With

End Sub

Public Function m_DoAct_Goto(ActId As Long) As Long

  Dim LabelName As String
  Dim NextId    As Long
  
  LabelName = m_ItemLst_Get(ActionLst(ActId).Prameters, c_Prm_Value)
  Select Case LabelName
  Case c_Val_End
    NextId = ActionNbr
  Case c_Val_Return
    If g_LastGoto_Id = c_Nothing Then
      ErrActId = ActId
      ErrTxt = "Return not possible in this context."
    Else
      NextId = g_LastGoto_Id + 1 'We go to the action after the goto
      g_LastGoto_Id = c_Nothing
    End If
  Case Else
    NextId = m_Action_LabelId(LabelName)
    If NextId = c_Nothing Then
      ErrActId = ActId
      ErrTxt = "Label '" & LabelName & "' not found."
    Else
      g_LastGoto_Id = ActId
    End If
  End Select

  m_DoAct_Goto = NextId

End Function

Public Function m_DoAct_Condition(ActId As Long) As Long

  Dim NextId As Long
  Dim Todo   As String
  Dim Value  As String
  
  Dim PosX   As Long
  Dim PosY   As Long
  Dim ColHex As String
  Dim Color  As Long
  
  Dim Ok        As Boolean
  
  'Get action's parametres in variables
  With ActionLst(ActId)
    ColHex = m_ItemLst_Get(.Prameters, c_Prm_Color)
    Color = m_HexToDec(ColHex)
    m_Action_GetPos .Prameters, PosX, PosY
  End With
  
  'Check condition
  Ok = False
  If m_Pixel_GetColor(PosX, PosY) = Color Then
    Ok = True
  End If
    
  DoEvents
    
  'Do the action todo
  If Ok Then
    m_Action_Todo False, ActionLst(ActId).Prameters, c_Prm_Then, Todo, Value
  Else
    m_Action_Todo False, ActionLst(ActId).Prameters, c_Prm_Else, Todo, Value
  End If
  Select Case Todo
  Case c_Todo_Next
    NextId = ActId + 1
  Case c_Todo_Goto
    NextId = 1 + m_Action_LabelId(Value)
  Case c_Todo_Skip
    NextId = ActId + 1 + Val(Value)
  Case Else
    ErrActId = ActId
    ErrTxt = "Condition unknown."
  End Select
  
  m_DoAct_Condition = NextId
  
End Function

Public Function m_Action_Name(ActId As Long) As String
'Returns the name of an action

  Dim X As String

  With ActionLst(ActId)
    X = m_ItemLst_Get(.Prameters, c_Prm_Name)
    If X = vbNullString Then
      X = m_ItemLst_Get(.Prameters, c_Prm_Type)
    End If
    If X = vbNullString Then
      X = "???"
    End If
  End With

  m_Action_Name = X

End Function

Public Function m_Action_LabelId(LabelName As String) As Long
'Return the id of the label

  Dim i  As Long
  Dim Ok As Boolean

  'Search the lablel
  i = 0
  Ok = False
  Do Until (i >= LabelNbr) Or Ok
    If LabelLst(i).Name = LabelName Then
      Ok = True
    Else
      i = i + 1
    End If
  Loop

  If Ok Then
    m_Action_LabelId = LabelLst(i).ActId
  Else
    m_Action_LabelId = c_Nothing
  End If

End Function

Public Sub m_ActLst_Init()

  Dim i     As Long
  Dim j     As Long
  Dim Lev   As Integer
  Dim Found As Boolean
  Dim Ok    As Boolean
  
  LabelNbr = 0
  
  'Init simple values
  Lev = 0
  For i = 0 To ActionNbr - 1
  
    With ActionLst(i)
    
      .Nbr = 0
      .StartTime = 0
      .Enabled = Not (m_ItemLst_Get(.Prameters, c_Prm_Enabled) = c_Val_No)
      
      If .Type = c_Act_Loop Then
        If m_ItemLst_Get(.Prameters, c_Prm_Unit) = vbNullString Then
          'It's a loop-end
          Lev = Lev - 1
          .LoopLevel = Lev
          .LoopBeg = c_Nothing
          .LoopEnd = i
        Else
          'It's a loop-begin
          .LoopLevel = Lev
          .LoopBeg = i
          .LoopEnd = c_Nothing
          Lev = Lev + 1
        End If
      Else
        .LoopLevel = Lev
        .LoopBeg = c_Nothing
        .LoopEnd = c_Nothing
        If .Type = c_Act_Label Then
          ReDim Preserve LabelLst(0 To LabelNbr)
          LabelLst(LabelNbr).Name = m_ItemLst_Get(.Prameters, c_Prm_Value)
          LabelLst(LabelNbr).ActId = i
          LabelNbr = LabelNbr + 1
        End If
      End If
      
    End With
    
  Next i
  
  'Search loop-end action for each loop-begin action.
  For i = 0 To ActionNbr - 1
    With ActionLst(i)
      If .LoopBeg = i Then 'Loop-Begin actions have been initialisated like this just before
        
        j = i + 1
        Lev = .LoopLevel
        Found = False
        Do Until Found Or (j > ActionNbr - 1)
          With ActionLst(j)
            If (.LoopEnd = j) And (.LoopLevel = Lev) Then 'Loop-End actions have been initialisated like this just before
              .LoopBeg = i
              .Enabled = ActionLst(i).Enabled
              Found = True
            Else
              j = j + 1
            End If
          End With
        Loop
        
        If Found Then
          .LoopEnd = j
        Else
          ErrActId = i
          ErrTxt = "The loop has no end."
        End If
        
      End If
    End With
  Next i
  
End Sub

Public Sub m_ActLst_Insert(ByVal ActId As Long, ActPrm As String)
'Insert a blanc action in the list at the specified index
  
  Dim i As Long
  
  'Insert a new action add the end of the list
  ActionNbr = ActionNbr + 1
  ReDim Preserve ActionLst(0 To (ActionNbr - 1))
  
  'Move up actions in order to free the ActId specified
  For i = (ActionNbr - 1) To (ActId + 1) Step -1
    ActionLst(i) = ActionLst(i - 1)
  Next i
  
  With ActionLst(ActId)
    .Type = m_ItemLst_Get(ActPrm, c_Prm_Type)
    .Prameters = ActPrm
  End With


End Sub

Public Sub m_ActLst_Swap(ActId1 As Long, ActId2 As Long)
'Move the bloc of actions in the list

  Dim Act As Action
  
  Act = ActionLst(ActId1)
  ActionLst(ActId1) = ActionLst(ActId2)
  ActionLst(ActId2) = Act
  
End Sub

Public Sub m_ActLst_Delete(ByVal ActId As Long)
'Delete the action item

  Dim i As Long

  For i = ActId + 1 To (ActionNbr - 1)
    ActionLst(i - 1) = ActionLst(i)
  Next i

  ActionNbr = ActionNbr - 1
  'Do not RedimPreserve because an error occurs when ActionNbr=0

End Sub

Public Function m_Action_Duration(ByVal ActId, Optional EnablesOcc As Boolean) As Long
'Returns the duration in millisecond corresponding to the parameters of the Action
'Returns a negative result if the duration is in occurence in stade of millisecond

  Dim Duration As Long
  Dim Nbr      As Long

  With ActionLst(ActId)
  
    Nbr = Abs(Val(m_ItemLst_Get(.Prameters, c_Prm_Nbr)))
  
    Select Case m_ItemLst_Get(.Prameters, c_Prm_Unit)
    Case c_Unit_Hour
      Duration = Nbr * CLng(1000) * 60 * 60
    Case c_Unit_Min
      Duration = Nbr * CLng(1000) * 60
    Case c_Unit_Sec
      Duration = Nbr * CLng(1000)
    Case c_Unit_Ms
      Duration = Nbr
    Case c_Unit_Occ
      If EnablesOcc Then
        Duration = -Nbr
      Else
        Duration = 0
        ErrTxt = "This Action doesn't allow the 'nbr' Unit."
        ErrActId = ActId
      End If
    Case Else
      Duration = 0
      ErrTxt = "The specified Unit is not correct."
      ErrActId = ActId
    End Select
  
  End With

  m_Action_Duration = Duration

End Function

Public Function m_Frm_SavePosition(ByRef Frm As Form, ByRef Parameters As String, Optional Size As Boolean = False)
'Save the position of the Form in the action's parameters

  Parameters = m_ItemLst_Set(Parameters, c_Prm_FrmX, Frm.Left)
  Parameters = m_ItemLst_Set(Parameters, c_Prm_FrmY, Frm.Top)
  If Size Then
    Parameters = m_ItemLst_Set(Parameters, c_Prm_FrmW, Frm.Width)
    Parameters = m_ItemLst_Set(Parameters, c_Prm_FrmH, Frm.Height)
  End If

End Function

Public Function m_Frm_LoadPosition(ByRef Frm As Form, ByRef Parameters As String, Optional Size As Boolean = False)
'Load the position of the Form that is saved in the action's parameters
'The property 'StartUpPosition' of the Form must be set to 'Manual'

  Dim X As Long
  Dim Y As Long
  Dim w As Long
  Dim h As Long

  X = Val(m_ItemLst_Get(Parameters, c_Prm_FrmX))
  Y = Val(m_ItemLst_Get(Parameters, c_Prm_FrmY))
  If (X = 0) And (Y = 0) Then
    'Place it on the center of the screen
    Frm.Move (Screen.Width - Frm.Width) \ 2, (Screen.Height - Frm.Height) \ 2
  Else
    Frm.Move X, Y
  End If

  If Size Then
    w = Val(m_ItemLst_Get(Parameters, c_Prm_FrmW))
    h = Val(m_ItemLst_Get(Parameters, c_Prm_FrmH))
    If w > 0 Then Frm.Width = w
    If h > 0 Then Frm.Height = h
  End If

End Function

Public Sub m_Ctr_FeedWithLabels(Ctr As ComboBox)

  Dim i As Long

  With Ctr
    .AddItem c_Val_Return
    .AddItem c_Val_End
    For i = 0 To LabelNbr - 1
      .AddItem LabelLst(i).Name
    Next i
  End With

End Sub

Public Sub m_Action_Todo(Save As Boolean, ByRef PrmStr As String, ByRef PrmName As String, ByRef Todo As String, ByRef Value As String)
'Get or set the todo informations in the paremeters's string

  Const c_Sep = ":"
  
  Dim X As String
  Dim p As Integer
  
  If Save Then
    PrmStr = m_ItemLst_Set(PrmStr, PrmName, Todo & c_Sep & Value)
  Else
    X = m_ItemLst_Get(PrmStr, PrmName)
    p = InStr(X, c_Sep)
    If p > 0 Then
      Todo = Left$(X, p - 1)
      Value = Mid$(X, p + 1)
    Else
      Todo = c_Todo_Next
      Value = ""
    End If
  End If

End Sub

Public Sub m_Action_GetPos(ByRef Prameters As String, ByRef X As Long, ByRef Y As Long)

  Dim Hwdn    As Long
  Dim Mediane As Boolean
  Dim PrmVal  As String
  Dim WinX    As Long
  Dim WinY    As Long
  Dim WinW    As Long
  Dim WinH    As Long

  X = Val(m_ItemLst_Get(Prameters, c_Prm_X))
  Y = Val(m_ItemLst_Get(Prameters, c_Prm_Y))
    
  PrmVal = m_ItemLst_Get(Prameters, c_Prm_Relative)
  m_RelativeInfo_Get PrmVal, Hwdn, Mediane
  
  If Hwdn = 0 Then
    If Mediane Then X = X + (Screen.Width / Screen.TwipsPerPixelX) / 2
  Else
    m_Win_GetPosition Hwdn, WinX, WinY
    X = X + WinX
    Y = Y + WinY
    If Mediane Then
      m_Win_GetSize Hwdn, WinW, WinH
      X = X + WinW / 2
    End If
  End If

End Sub
