VERSION 5.00
Begin VB.Form frm_Act_Condition 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Condition"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txt_Else_Value 
      Height          =   315
      Left            =   2520
      TabIndex        =   24
      Text            =   "0"
      Top             =   3720
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox lst_Else_Value 
      Height          =   315
      Left            =   2520
      TabIndex        =   25
      Top             =   3720
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txt_Then_Value 
      Height          =   315
      Left            =   2520
      TabIndex        =   22
      Text            =   "0"
      Top             =   3240
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.ComboBox lst_Then_Value 
      Height          =   315
      Left            =   2520
      TabIndex        =   23
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ComboBox lst_Else_Todo 
      Height          =   315
      ItemData        =   "frm_Act_Condition.frx":0000
      Left            =   840
      List            =   "frm_Act_Condition.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   3720
      Width           =   1575
   End
   Begin VB.ComboBox lst_Then_Todo 
      Height          =   315
      ItemData        =   "frm_Act_Condition.frx":0004
      Left            =   840
      List            =   "frm_Act_Condition.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   3240
      Width           =   1575
   End
   Begin VB.CommandButton btn_Ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   13
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CheckBox chk_Enabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      Height          =   255
      Left            =   1680
      TabIndex        =   12
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txt_Name 
      Height          =   315
      Left            =   120
      TabIndex        =   11
      Top             =   360
      Width           =   2535
   End
   Begin VB.Frame frm_Color 
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   4455
      Begin VB.ComboBox lst_Relative 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   240
         Width           =   2175
      End
      Begin VB.TextBox txt_Color 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.CheckBox chk_ChangeXY 
         Caption         =   "Edit"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chk_ChangeColor 
         Caption         =   "Edit"
         Height          =   255
         Left            =   1800
         TabIndex        =   4
         Top             =   1320
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txt_X 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.TextBox txt_Y 
         Height          =   315
         Left            =   960
         TabIndex        =   2
         Text            =   "0"
         Top             =   960
         Width           =   735
      End
      Begin VB.CommandButton btn_Position 
         Caption         =   "Locate"
         Height          =   375
         Left            =   3240
         TabIndex        =   1
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lbl_Relative 
         Caption         =   "Relative to:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_Y 
         Caption         =   "Y"
         Height          =   255
         Left            =   960
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl_X 
         Caption         =   "X"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   735
      End
      Begin VB.Label lbl_Color 
         Caption         =   "Color Id"
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   720
         Width           =   855
      End
      Begin VB.Shape shp_Color 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   315
         Left            =   2760
         Top             =   960
         Width           =   375
      End
      Begin VB.Label lbl_LocInfo 
         Caption         =   "Click or press [Enter] to validate, or [Esc] to cancel."
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   1680
         Visible         =   0   'False
         Width           =   4215
      End
   End
   Begin VB.Label Label1 
      Caption         =   "If the pixel has the specified color:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   840
      Width           =   4455
   End
   Begin VB.Label lbl_Else 
      Caption         =   "Else:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   3720
      Width           =   735
   End
   Begin VB.Label lbl_Then 
      Caption         =   "Then:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   3240
      Width           =   735
   End
   Begin VB.Label lbl_Name 
      Caption         =   "Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label lbl_ActiveWin 
      Caption         =   " "
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   4560
      Width           =   4455
   End
End
Attribute VB_Name = "frm_Act_Condition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim scThen As New oSuperCombo
Dim scElse As New oSuperCombo
Dim scRelative As New oSuperCombo

Dim CtrThenValue As Control
Dim CtrElseValue As Control

Public ActiveWinHwnd As Long
Public Mediane As Boolean

Private Sub btn_Cancel_Click()

  Unload Me

End Sub

Private Sub btn_Ok_Click()

  Dim Todo  As String
  Dim Value As String

  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Name, txt_Name)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Enabled, IIf(chk_Enabled = vbChecked, c_Val_Yes, c_Val_No))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_X, txt_X)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Y, txt_Y)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Color, txt_Color)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Relative, scRelative.CurrValue)
  
  'Save 'Then' information
  Todo = scThen.CurrValue
  Value = p_CtrValue(CtrThenValue)
  m_Action_Todo True, g_CurrAct_Prm, c_Prm_Then, Todo, Value

  'Save 'Else' information
  Todo = scElse.CurrValue
  Value = ""
  Value = p_CtrValue(CtrElseValue)
  m_Action_Todo True, g_CurrAct_Prm, c_Prm_Else, Todo, Value
  
    
  m_Frm_SavePosition Me, g_CurrAct_Prm
  g_CurrAct_Ok = True
  Unload Me

End Sub

Private Sub btn_Position_Click()

  m_Pos_LocatePoint Me, (chk_ChangeXY.Value = vbChecked), (chk_ChangeColor.Value = vbChecked)

End Sub

Private Sub Form_Load()

  Dim Todo  As String
  Dim Value As String

  m_Win_StayOnTop Me.Hwnd, True

  txt_Name = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Name)
  chk_Enabled.Value = IIf(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Enabled) = c_Val_No, vbUnchecked, vbChecked)
  txt_X = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_X)
  txt_Y = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Y)
  txt_Color = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Color)

  'Valeurs par défaut
  If Val(txt_X) = 0 Then txt_X = "0"
  If Val(txt_Y) = 0 Then txt_Y = "0"

  With scThen
    .SetCtrObj lst_Then_Todo, True
    .AddItem c_Todo_Next, "Next action"
    .AddItem c_Todo_Goto, "Goto label"
    .AddItem c_Todo_Skip, "Skip a num of act"
    m_Action_Todo False, g_CurrAct_Prm, c_Prm_Then, Todo, Value
    .SetCtrVal Todo
    If Not (CtrThenValue Is Nothing) Then CtrThenValue = Value
  End With
  m_Ctr_FeedWithLabels lst_Then_Value

  With scElse
    .SetCtrObj lst_Else_Todo, True
    .AddItem c_Todo_Next, "Next action"
    .AddItem c_Todo_Goto, "Goto label"
    .AddItem c_Todo_Skip, "Skip a num of act"
    m_Action_Todo False, g_CurrAct_Prm, c_Prm_Else, Todo, Value
    .SetCtrVal Todo
    If Not (CtrElseValue Is Nothing) Then CtrElseValue = Value
  End With
  m_Ctr_FeedWithLabels lst_Else_Value

  m_RelativeInfo_FeedLst scRelative, lst_Relative
  scRelative.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Relative), True

  m_Frm_LoadPosition Me, g_CurrAct_Prm

End Sub

Private Sub lst_Else_Todo_Click()

  If Not (CtrElseValue Is Nothing) Then CtrElseValue.Visible = False
  
  Select Case scElse.CurrValue
  Case c_Todo_Goto: Set CtrElseValue = lst_Else_Value
  Case c_Todo_Skip: Set CtrElseValue = txt_Else_Value
  Case Else: Set CtrElseValue = Nothing
  End Select
  
  If Not (CtrElseValue Is Nothing) Then CtrElseValue.Visible = True

End Sub

Private Sub lst_Then_Todo_Click()

  If Not (CtrThenValue Is Nothing) Then CtrThenValue.Visible = False
  
  Select Case scThen.CurrValue
  Case c_Todo_Goto: Set CtrThenValue = lst_Then_Value
  Case c_Todo_Skip: Set CtrThenValue = txt_Then_Value
  Case Else: Set CtrThenValue = Nothing
  End Select
  
  If Not (CtrThenValue Is Nothing) Then CtrThenValue.Visible = True

End Sub

Private Function p_CtrValue(Ctr As Control) As String

  If (Ctr Is Nothing) Then
    p_CtrValue = ""
  Else
    p_CtrValue = Ctr
  End If

End Function

Private Sub txt_Color_Change()

  shp_Color.BackColor = m_HexToDec(txt_Color.Text)
  
End Sub

Private Sub lst_Relative_Click()

  m_RelativeInfo_Actualize Me, scRelative
  
End Sub

