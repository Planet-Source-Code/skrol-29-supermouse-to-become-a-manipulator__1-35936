VERSION 5.00
Begin VB.Form frm_Main 
   Caption         =   "Supermouse"
   ClientHeight    =   6390
   ClientLeft      =   5385
   ClientTop       =   3555
   ClientWidth     =   4920
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   4920
   Begin VB.Frame frm_Add 
      Caption         =   "New Action"
      Height          =   1215
      Left            =   3360
      TabIndex        =   8
      Tag             =   "MargeD=;MargeH="
      Top             =   2760
      Width           =   1455
      Begin VB.CommandButton btn_Add 
         Caption         =   "Insert"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   1215
      End
      Begin VB.ComboBox lst_Add 
         Height          =   315
         ItemData        =   "frm_Main.frx":0442
         Left            =   120
         List            =   "frm_Main.frx":0444
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame frm_Action 
      Caption         =   "Action"
      Height          =   2175
      Left            =   3360
      TabIndex        =   3
      Tag             =   "MargeD=;MargeH="
      Top             =   360
      Width           =   1455
      Begin VB.CommandButton btn_Copy 
         Caption         =   "Copy"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton btn_Edit 
         Caption         =   "Edit"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.CommandButton btn_Delete 
         Caption         =   "Delete"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   1680
         Width           =   1215
      End
      Begin VB.CommandButton btn_MoveUp 
         Caption         =   "^"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton btn_MoveDown 
         Caption         =   "v"
         Height          =   375
         Left            =   600
         TabIndex        =   4
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton btn_Stop 
      Caption         =   "Stop"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Tag             =   "MargeD=;MargeB="
      Top             =   5400
      Width           =   1215
   End
   Begin VB.ListBox lst_Action 
      Height          =   5325
      ItemData        =   "frm_Main.frx":0446
      Left            =   120
      List            =   "frm_Main.frx":0448
      MultiSelect     =   2  'Extended
      TabIndex        =   0
      Tag             =   "MargeG=;MargeD=;MargeH=;MargeB="
      Top             =   480
      Width           =   3135
   End
   Begin VB.Label lbl_Duration 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Tag             =   "MargeG=;MargeD=;MargeB="
      Top             =   6120
      Width           =   4815
   End
   Begin VB.Label lbl_Info 
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Tag             =   "MargeG=;MargeD=;MargeB="
      Top             =   5880
      Width           =   4695
   End
   Begin VB.Label lbl_Action 
      Caption         =   "Actions"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Tag             =   "MargeG=;MargeH="
      Top             =   240
      Width           =   2055
   End
   Begin VB.Menu mnu_File 
      Caption         =   "&File"
      Begin VB.Menu mnu_New 
         Caption         =   "&New"
      End
      Begin VB.Menu mnu_Open 
         Caption         =   "&Open"
      End
      Begin VB.Menu mnu_Save 
         Caption         =   "&Save"
      End
      Begin VB.Menu mnu_SaveAs 
         Caption         =   "Save &as"
      End
      Begin VB.Menu mnu_space2 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Quit 
         Caption         =   "&Quit"
      End
   End
   Begin VB.Menu mnu_Tools 
      Caption         =   "&Tools"
      Begin VB.Menu mnu_Run 
         Caption         =   "&Run"
      End
      Begin VB.Menu mnu_Space3 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_ReplaceAllPoints 
         Caption         =   "&Change all Points"
      End
      Begin VB.Menu mnu_ReplaceAllSpeed 
         Caption         =   "Replace all &Speeds"
      End
      Begin VB.Menu mnu_ReplaceAllRadius 
         Caption         =   "Remplace all &Radius"
      End
      Begin VB.Menu mnu_Space4 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_OPtions 
         Caption         =   "&Options"
      End
   End
   Begin VB.Menu mnu_Question 
      Caption         =   "&?"
      Begin VB.Menu mnu_ActiveWinInfo 
         Caption         =   "Info on active Windows"
      End
      Begin VB.Menu mnu_Space5 
         Caption         =   "-"
      End
      Begin VB.Menu mnu_Help 
         Caption         =   "&Help"
      End
      Begin VB.Menu mnu_About 
         Caption         =   "&About Supermouse"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim scAdd As New oSuperCombo

Dim ActIdMin As Long
Dim ActIdMax As Long
Dim ActIdNbr As Long
Dim ActIdChg As Boolean

Private Sub p_Run()

  Dim ActId      As Long
  Dim PreviousId As Long
  
  ActId = 0
  PreviousId = c_Nothing
  
  g_StartTime = Now()
  g_AskForStop = False
  btn_Stop.Enabled = True
  
  m_Err_Init
  
  lbl_Info = "Script in progress..."
  
  g_RestorePos = (m_Option_Get(c_Opt_Mouse_RestorePos, "", c_Prm_Enabled) <> c_Val_No)
  If g_RestorePos Then m_Mouse_GetCurrPosPixels SmsPosX, SmsPosY
  
  g_StayOnTop = (m_Option_Get(c_Opt_Main_StayOnTop, "", c_Prm_Enabled) <> c_Val_No)
  m_Win_StayOnTop Me.Hwnd, g_StayOnTop
  
  Do Until p_Action_IsEnd(ActId) Or g_AskForStop Or (ErrActId <> c_Nothing)
  
    p_ActionLst_SelectOne ActId, PreviousId
    lbl_Duration = "Duration : " & Format$(Now() - g_StartTime, "hh:nn:ss")
    DoEvents
  
    PreviousId = ActId
    ActId = m_DoAct_Any(PreviousId)
  
  Loop

  If ErrActId <> c_Nothing Then
    lbl_Info = "Erreur : " & ErrTxt
  ElseIf g_AskForStop Then
    lbl_Info = "Script stoped by user."
  Else
    lbl_Info = "Script ended."
  End If

  lst_Action.SetFocus
  btn_Stop.Enabled = False

  m_Win_StayOnTop Me.Hwnd, False
  
  If g_RunMode = 2 Then
    If m_Option_Get(c_Opt_Main_AutoQuit, "", c_Prm_Enabled) <> c_Val_No Then
      Unload Me
    End If
  End If

End Sub

Private Sub btn_Copy_Click()

  Dim ActId As Long
  Dim X     As String
  
  ActId = lst_Action.ListIndex

  If p_Action_IsEnd(ActId) Then
    Exit Sub
  End If
  
  If ActionLst(ActId).Type = c_Act_Loop Then
    Exit Sub
  End If
  
  If MsgBox("Are you sure you want to copy this action ?", vbQuestion + vbYesNo) = vbYes Then
    'copy the current actions parameters int the global variable
    g_CurrAct_Prm = ActionLst(ActId).Prameters
    'Chnage the name
    X = m_ItemLst_Get(ActionLst(ActId).Prameters, c_Prm_Name)
    X = m_Txt_IncrementeCopie(X)
    ActionLst(ActId).Prameters = m_ItemLst_Set(ActionLst(ActId).Prameters, c_Prm_Name, X)
    'Inserte the new action before the current action
    m_ActLst_Insert ActId, g_CurrAct_Prm
    'Actualize the list
    p_ActionLst_Actualize ActId + 1
  End If

End Sub

Private Sub btn_Add_Click()

  p_ActionLst_AddModif ActIdMin, True, scAdd.CurrValue
  
End Sub

Private Sub btn_Delete_Click()

  Dim ActId As Long
  Dim i     As Long
  Dim ToDelLst() As Long
  Dim ToDelNbr   As Long
  Dim Nbr   As Long
  
  Nbr = 0
  
  If ActIdNbr = 1 Then
    If p_Action_IsEnd(ActIdMin) Then Exit Sub
    If m_Action_IsLoopE(ActIdMin) Then Exit Sub
  End If
  
  If MsgBox("Are you sure you want to delete " & ActIdNbr & " actions?", vbDefaultButton2 + vbYesNo + vbQuestion) = vbYes Then
    
    'Scan for End-Loop item to delete
    ToDelNbr = 0
    For ActId = ActIdMin To ActIdMax
      If m_Action_IsLoopB(ActId) Then
        i = ActionLst(ActId).LoopEnd
        If (i < ActIdMin) Or (i > ActIdMax) Then
          ReDim Preserve ToDelLst(0 To ToDelNbr)
          ToDelLst(ToDelNbr) = i
          ToDelNbr = ToDelNbr + 1
        End If
      End If
    Next ActId
    
    'Delete item in order from the higer to the lower
    '------------------------------------------------
    
    'Delete End-Loop item after the block to delete
    For i = ToDelNbr - 1 To 0 Step -1
      If ToDelLst(ToDelNbr) > ActIdMax Then
        m_ActLst_Delete ToDelLst(ToDelNbr)
        Nbr = Nbr + 1
      End If
    Next i
    
    'Delete items
    For ActId = ActIdMax To ActIdMin Step -1
      If m_Action_IsLoopE(ActId) Then
        i = ActionLst(ActId).LoopBeg
        If (i >= ActIdMin) And (i <= ActIdMax) Then
          m_ActLst_Delete ActId
          Nbr = Nbr + 1
        End If
      Else
        m_ActLst_Delete ActId
        Nbr = Nbr + 1
      End If
    Next ActId
    
    'Delete End-Loop item befor the block to delete
    For i = ToDelNbr - 1 To 0 Step -1
      If ToDelLst(ToDelNbr) < ActIdMin Then
        Nbr = Nbr + 1
        m_ActLst_Delete ToDelLst(ToDelNbr)
      End If
    Next i
    
    'Actualize display
    If Nbr > 0 Then
      ActId = ActIdMin
      p_ActionLst_Actualize ActId
    End If
  
  End If

End Sub

Private Sub btn_Edit_Click()

  p_ActionLst_AddModif ActIdMin, False
  
End Sub

Private Sub btn_MoveDown_Click()

  Dim ActId  As Long
  Dim IdMin  As Long
  Dim IdMax  As Long
  
  If ActIdMax >= (ActionNbr - 1) Then Exit Sub
  
  IdMin = ActIdMin + 1
  IdMax = ActIdMax + 1
  
  For ActId = ActIdMax To ActIdMin Step -1
    m_ActLst_Swap ActId, ActId + 1
  Next ActId
  
  p_ActionLst_Actualize IdMin, IdMax

End Sub

Private Sub btn_MoveUp_Click()

  Dim ActId  As Long
  Dim IdMin  As Long
  Dim IdMax  As Long
  
  If ActIdMin <= 0 Then Exit Sub
  
  IdMin = ActIdMin - 1
  IdMax = ActIdMax - 1
  
  For ActId = ActIdMin To ActIdMax
    m_ActLst_Swap ActId - 1, ActId
  Next ActId
  
  p_ActionLst_Actualize IdMin, IdMax

End Sub

Private Sub btn_Stop_Click()
  g_AskForStop = True
End Sub

Private Sub Form_Activate()

  If g_RunMode = 1 Then
    g_RunMode = 2
    p_Run
  End If
  
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

  'MsgBox "keycode=" & KeyCode & ";shift=" & Shift

End Sub

Private Sub Form_Load()

  Dim CmdLinePrm As String

  g_ForbiddenWin = Me.Hwnd

  m_Frm_PositionCtr Me, True

  'Initialize controls
  With scAdd
    .SetCtrObj lst_Add, True
    .AddItem c_Act_Mouse, c_Act_Mouse
    .AddItem c_Act_Wait, c_Act_Wait
    .AddItem c_Act_Loop, c_Act_Loop
    .AddItem c_Act_Condition, c_Act_Condition
    .AddItem c_Act_Label, c_Act_Label
    .AddItem c_Act_Goto, c_Act_Goto
    .AddItem c_Act_Execute, c_Act_Execute
    .AddItem c_Act_Keys, c_Act_Keys
    .AddItem c_Act_Message, c_Act_Message
    .AddItem c_Act_Comment, c_Act_Comment
    .SetCtrVal "", True
  End With

  'Actualize the action list
  p_ActionLst_Actualize 0
  g_Dirty = False

  'Load and execute the script in the commande line
  CmdLinePrm = Trim$("" & Command())
  If CmdLinePrm <> "" Then
    If Len(CmdLinePrm) > 1 Then
      If (Left$(CmdLinePrm, 1) = """") And (Right$(CmdLinePrm, 1) = """") Then
        CmdLinePrm = Mid$(CmdLinePrm, 2, Len(CmdLinePrm) - 2)
      End If
    End If
    If CmdLinePrm = "/?" Then
      MsgBox "Command line parameters :" & vbNewLine & vbNewLine & "/? : help information" & vbNewLine & "[file] : open and execute the script", vbInformation
      Unload Me
    ElseIf m_File_Exists(CmdLinePrm) Then
      g_RunMode = 1
      p_File_Load CmdLinePrm
    Else
      MsgBox "The parameter given in the command line is not a file path, or the file specified is unknown.", vbCritical
    End If
  End If

End Sub

Private Sub Form_Resize()

  m_Frm_PositionCtr Me, False
  
End Sub

Private Sub lst_Action_Click()

  p_ActionLst_Check

End Sub

Private Sub lst_Action_DblClick()
  
  btn_Edit_Click

End Sub

Private Sub lst_Action_KeyPress(KeyAscii As Integer)

  If KeyAscii = 13 Then 'key {Enter} pressed
    btn_Edit_Click
  End If

End Sub

Private Sub mnu_About_Click()
  MsgBox "Supermouse " & App.Major & "." & App.Minor & vbNewLine & vbNewLine & "----------" & vbNewLine & "Skrol 29" & vbNewLine & "skrol29@freesurf.fr", vbInformation
End Sub

Private Sub mnu_ActiveWinInfo_Click()

  m_Win_ShowActiveWinList 11
  
End Sub

Private Sub mnu_Help_Click()

  frm_Help.Show vbModal

End Sub

Private Sub mnu_ReplaceAllRadius_Click()

  Dim ActId As Long
  Dim X     As Long
  Dim Resp     As String
  Dim RadiusOld As Long
  Dim RadiusNew As Long
  Dim RadiusX   As Long
  
  ActId = lst_Action.ListIndex
  If Not p_Action_IsEnd(ActId) Then
    If ActionLst(ActId).Type = c_Act_Mouse Then
      RadiusOld = Val(m_ItemLst_Get(ActionLst(ActId).Prameters, c_Prm_Radius))
      If RadiusOld > 0 Then
      
        Resp = InputBox$("If you want to replace all radius values equale to " & RadiusOld & " into all actions, you have to enter the new radius :", "Replace radius", "" & RadiusOld)
        RadiusNew = Val(Resp)
        
        If RadiusNew > 0 Then
          
          For X = 0 To (ActionNbr - 1)
            If ActionLst(X).Type = c_Act_Mouse Then
              RadiusX = Val(m_ItemLst_Get(ActionLst(X).Prameters, c_Prm_Radius))
              If RadiusX = RadiusOld Then
                ActionLst(X).Prameters = m_ItemLst_Set(ActionLst(X).Prameters, c_Prm_Radius, "" & RadiusNew)
              End If
            End If
          Next X
          
          p_ActionLst_Actualize
         
        End If
        
      End If
    End If
  End If

End Sub

Private Sub mnu_ReplaceAllSpeed_Click()
  
  Dim ActId As Long
  Dim X     As Long
  Dim Resp     As String
  Dim SpeedOld As Long
  Dim SpeedNew As Long
  Dim SpeedX   As Long
  
  ActId = lst_Action.ListIndex
  If Not p_Action_IsEnd(ActId) Then
    If ActionLst(ActId).Type = c_Act_Mouse Then
      SpeedOld = Val(m_ItemLst_Get(ActionLst(ActId).Prameters, c_Prm_Speed))
      If SpeedOld > 0 Then
      
        Resp = InputBox$("If you want to replace all speed values equale to " & SpeedOld & " into all actions, you have to enter the new speed :", "Replace speeds", "" & SpeedOld)
        SpeedNew = Val(Resp)
        
        If SpeedNew > 0 Then
          
          For X = 0 To (ActionNbr - 1)
            If ActionLst(X).Type = c_Act_Mouse Then
              SpeedX = Val(m_ItemLst_Get(ActionLst(X).Prameters, c_Prm_Speed))
              If SpeedX = SpeedOld Then
                ActionLst(X).Prameters = m_ItemLst_Set(ActionLst(X).Prameters, c_Prm_Speed, "" & SpeedNew)
              End If
            End If
          Next X
          
          p_ActionLst_Actualize
          
        End If
        
      End If
    End If
  End If
  
End Sub

Private Sub mnu_ReplaceAllPoints_Click()

  Dim ActId As Long
  Dim i     As Long
  Dim X0    As Long
  Dim Y0    As Long
  Dim Tx    As Long
  Dim Ty    As Long
  
  ActId = lst_Action.ListIndex
  If Not p_Action_IsEnd(ActId) Then
    If ActionLst(ActId).Type = c_Act_Mouse Then
      If MsgBox("If you want to move all point, you have to first modify the position of the selected point." & vbNewLine & "After this modification, all other points will be modified with the same move." & vbNewLine & "Do you want to continue ?", vbDefaultButton2 + vbInformation + vbYesNo, "Move all points") = vbYes Then
        
        'Get the motion
        With ActionLst(ActId)
          X0 = Val(m_ItemLst_Get(.Prameters, c_Prm_X))
          Y0 = Val(m_ItemLst_Get(.Prameters, c_Prm_Y))
          p_ActionLst_AddModif ActId, False, , True
          Tx = Val(m_ItemLst_Get(.Prameters, c_Prm_X)) - X0
          Ty = Val(m_ItemLst_Get(.Prameters, c_Prm_Y)) - Y0
        End With
        
        If (Tx <> 0) Or (Ty <> 0) Then
          'Permorms the motion for all mouse actions
          For i = 0 To (ActionNbr - 1)
            If i <> ActId Then
              With ActionLst(i)
                If (.Type = c_Act_Mouse) Or (.Type = c_Act_Wait) Then
                  X0 = Val(m_ItemLst_Get(.Prameters, c_Prm_X))
                  Y0 = Val(m_ItemLst_Get(.Prameters, c_Prm_Y))
                  If (X0 <> 0) And (Y0 <> 0) Then
                    .Prameters = m_ItemLst_Set(.Prameters, c_Prm_X, "" & (X0 + Tx))
                    .Prameters = m_ItemLst_Set(.Prameters, c_Prm_Y, "" & (Y0 + Ty))
                  End If
                End If
              End With
            End If
          Next i
          MsgBox "All points have been moved.", vbInformation
        End If
      
        'Actualize even if ther is no translation because ther may be a
        p_ActionLst_Actualize
      
      End If
    End If
  End If

End Sub

Private Sub mnu_New_Click()

  If m_SmsFile_Check() Then
    g_File = ""
    ActionNbr = 0
    OptionNbr = 0
    p_ActionLst_Actualize 0
    g_Dirty = False
  End If

End Sub

Private Sub mnu_Open_Click()

  Dim File As String

  If m_SmsFile_Check() Then 'Check if some changes have been got since the last saving.
    File = m_File_Open("Open", c_Filter, True, , App.Path, Me.Hwnd)
    If File <> "" Then
      p_File_Load File
    End If
  End If

End Sub

Private Sub mnu_OPtions_Click()
  frm_Options.Show vbModal
End Sub

Private Sub mnu_Quit_Click()
  If m_SmsFile_Check() Then
    Unload Me
  End If
End Sub

Private Sub p_ActionLst_AddModif(ActId As Long, Add As Boolean, Optional ActType As String, Optional NoActualize As Boolean)
'Add or modify an action

  g_CurrAct_Id = ActId
  g_CurrAct_Add = Add
  g_CurrAct_Ok = False
  
  If Add Then
    g_CurrAct_Prm = ""
    g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Type, ActType)
    g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Name, "<" & c_Txt_New & " " & ActType & ">")
  Else
    'Check that it is not the last list item
    If p_Action_IsEnd(ActId) Then Exit Sub
    With ActionLst(ActId)
      'Check it is not a loop-end item
      If m_Action_IsLoopE(ActId) Then Exit Sub
      ActType = .Type
      g_CurrAct_Prm = .Prameters
    End With
  End If

  Select Case ActType
  Case c_Act_Mouse
    Me.Visible = False
    frm_Act_Mouse.Show vbModal
    Me.Visible = True
  Case c_Act_Wait
    Me.Visible = False
    frm_Act_Wait.Show vbModal
    Me.Visible = True
  Case c_Act_Loop
    frm_Act_Loop.Show vbModal
  Case c_Act_Comment
    p_Action_Misc "Comment", "Enter a comment"
  Case c_Act_Execute
    p_Action_Misc "Command", "Enter a command line to execute"
  Case c_Act_Keys
    p_Action_Misc "Keys", "Entrer the key string to send :", "[enter]=~ ; [ctrl]=^ ; [alt]=% ; [shift]=+"
  Case c_Act_Label
    p_Action_Misc "Label", "Enter a unique name for the label"
  Case c_Act_Goto
    p_Action_Misc "Goto", "Enter the name of an existing label or 'Return'."
  Case c_Act_Message
    p_Action_Misc "Message", "Enter the message to display:"
  Case c_Act_Condition
    Me.Visible = False
    frm_Act_Condition.Show vbModal
    Me.Visible = True
  Case Else
  End Select

  If g_CurrAct_Ok Then
    
    If Add Then
      m_ActLst_Insert ActId, g_CurrAct_Prm
      If ActType = c_Act_Loop Then
        m_ActLst_Insert ActId + 1, m_ItemLst_Set("", c_Prm_Type, c_Act_Loop)
      End If
    End If
    
    ActionLst(ActId).Prameters = g_CurrAct_Prm
    
    If Not NoActualize Then
      If Add Then
        p_ActionLst_Actualize ActId + 1
      Else
        p_ActionLst_Actualize
      End If
    End If
    
  End If

End Sub

Private Sub p_ActionLst_Actualize(Optional IdMin As Long = c_Nothing, Optional IdMax As Long = c_Nothing)
'Clear and refeed the list. Select the asked items

  Dim i As Long

  ActIdChg = True

  If IdMin = c_Nothing Then
    IdMin = ActIdMin
    IdMax = ActIdMax
  Else
    If IdMax = c_Nothing Then
      IdMax = IdMin
    End If
  End If
  
  m_ActLst_Init
  
  With lst_Action
  
    .Clear
    For i = 0 To (ActionNbr - 1)
      .AddItem String(Abs(ActionLst(i).LoopLevel) * 2, " ") & m_Action_Caption(i)
    Next i
    .AddItem c_Val_End
  
    If .MultiSelect = 0 Then
      .ListIndex = IdMin
    Else
      For i = IdMin To IdMax
        .Selected(i) = True
      Next i
    End If
    
  End With
    
  g_Dirty = True

  ActIdMin = IdMin
  ActIdMax = IdMax
  ActIdNbr = IdMax - IdMin + 1
  
  ActIdChg = False

End Sub

Private Function p_Action_IsEnd(ActId) As Boolean

  p_Action_IsEnd = (ActId >= ActionNbr)

End Function

Private Sub p_Action_Misc(TitleTxt As String, ValueTxt As String, Optional HelpTxt As String)

  With frm_Act_Misc
    .Caption = TitleTxt
    .lbl_Value = ValueTxt
    .lbl_Help = HelpTxt
    .Show vbModal
  End With

End Sub

Private Sub mnu_Run_Click()
  p_Run
End Sub

Private Sub mnu_Save_Click()

  If g_File = "" Then 'The current script is a new one
    g_File = m_File_Open("Enregistrer", c_Filter, False, "", App.Path, Me.Hwnd)
  End If

  If g_File <> "" Then
    p_File_Save g_File
  End If

End Sub

Private Sub mnu_SaveAs_Click()

  Dim X As String
  
  X = m_File_Open("Enregistrer", c_Filter, False, "", App.Path, Me.Hwnd)
  If X <> "" Then
    p_File_Save X
  End If

End Sub

Private Function p_File_Save(File As String)

  'Save some options
  m_Option_Set c_Opt_Main_RestorePos, "", c_Prm_FrmX, "" & Me.Left
  m_Option_Set c_Opt_Main_RestorePos, "", c_Prm_FrmY, "" & Me.Top
  m_Option_Set c_Opt_Main_RestorePos, "", c_Prm_FrmW, "" & Me.Width
  m_Option_Set c_Opt_Main_RestorePos, "", c_Prm_FrmH, "" & Me.Height
  
  If m_Option_Get(c_Opt_DateCrea, "", c_Prm_Value) = "" Then
    m_Option_Set c_Opt_DateCrea, "", c_Prm_Value, Format$(Now(), "yyyy-mm-dd hh:nn:ss")
  End If
  m_Option_Set c_Opt_DateModif, "", c_Prm_Value, Format$(Now(), "yyyy-mm-dd hh:nn:ss")
  m_Option_Set c_Opt_Version, "", c_Prm_Value, App.Major & "." & App.Minor
    
  'Write the file info from ActionLst() and OptionsLst() arrays
  m_SmsFile_Save File
  
  'Actualize
  Me.Caption = m_File_GetName(File) & " - " & App.ProductName
  g_Dirty = False
  g_File = File

End Function

Private Function p_File_Load(File As String)

  Dim i As Long
  
  
  'Load the file info into ActionLst() and OptionsLst() arrays
  m_SmsFile_Load (File)
  
  'Perform the 'RestorePos' option
  If m_Option_Get(c_Opt_Main_RestorePos, "", c_Prm_Enabled) <> c_Val_No Then
    i = m_Option_Found(c_Opt_Main_RestorePos, "")
    If i <> c_Nothing Then
      m_Frm_LoadPosition Me, OptionLst(i), True
    End If
  End If
  
  'Actualize
  p_ActionLst_Actualize 0
  Me.Caption = m_File_GetName(File) & " - " & App.ProductName
  g_Dirty = False
  g_File = File

End Function

Private Sub p_ActionLst_SelectOne(ActId As Long, Optional PreviousId As Long = c_Nothing)

  Dim i As Long
  
  ActIdChg = True
  
  With lst_Action
  
    If .MultiSelect = 0 Then
      .ListIndex = ActId
    Else
      If PreviousId = c_Nothing Then
        For i = 0 To .ListCount - 1
          .Selected(i) = (i = ActId)
        Next i
      Else
        If ActId <> c_Nothing Then
          .Selected(ActId) = True
        End If
        .Selected(PreviousId) = False
      End If
    End If
    
  End With

  ActIdChg = False

End Sub

Private Sub p_ActionLst_Check(Optional ActId As Long = c_Nothing)
'Calculate ActIdMin, ActIdMax and ActIdNbr

  If ActIdChg Then Exit Sub

  Dim i   As Long
  Dim Sel As Boolean
  
  With lst_Action
  
    If ActId = c_Nothing Then
      ActId = .ListIndex
    Else
      .Selected(ActId) = True
    End If
    
    ActIdMin = ActId
    ActIdMax = ActId
    ActIdNbr = 1
  
    If .MultiSelect <> 0 Then
  
      Sel = True
      For i = .ListIndex + 1 To .ListCount - 1
        If .Selected(i) Then
          If Sel Then
            ActIdMax = i
            ActIdNbr = ActIdNbr + 1
          Else
            .Selected(i) = False
          End If
        Else
          If Sel Then Sel = False
        End If
      Next i
    
      Sel = True
      For i = .ListIndex - 1 To 0 Step -1
        If .Selected(i) Then
          If Sel Then
            ActIdMin = i
            ActIdNbr = ActIdNbr + 1
          Else
            .Selected(i) = False
          End If
        Else
          If Sel Then Sel = False
        End If
      Next i
  
    End If
  
  End With

End Sub

