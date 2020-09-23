VERSION 5.00
Begin VB.Form frm_Act_Mouse 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mouse action"
   ClientHeight    =   4350
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4350
   ScaleWidth      =   5160
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox lst_Move 
      Height          =   315
      ItemData        =   "frm_Position.frx":0000
      Left            =   960
      List            =   "frm_Position.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.ComboBox lst_Button2 
      Height          =   315
      ItemData        =   "frm_Position.frx":0004
      Left            =   2760
      List            =   "frm_Position.frx":0006
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   3240
      Width           =   1935
   End
   Begin VB.ComboBox lst_Button1 
      Height          =   315
      ItemData        =   "frm_Position.frx":0008
      Left            =   120
      List            =   "frm_Position.frx":000A
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3240
      Width           =   1935
   End
   Begin VB.Frame frm_Point 
      Height          =   1575
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   4935
      Begin VB.ComboBox lst_Relative 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton btn_Position 
         Caption         =   "Locate"
         Height          =   375
         Left            =   1800
         TabIndex        =   21
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_Radius 
         Height          =   315
         Left            =   4200
         TabIndex        =   6
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txt_Speed 
         Height          =   315
         Left            =   3360
         TabIndex        =   5
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txt_X 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox txt_Y 
         Height          =   315
         Left            =   960
         TabIndex        =   4
         Text            =   "0"
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lbl_Relative 
         Caption         =   "Relative to:"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_LocInfo 
         Caption         =   "Click or press [Enter] to validate, or [Esc] to cancel."
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Label lbl_T 
         Caption         =   "Y"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lbl_X 
         Caption         =   "X"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   255
      End
      Begin VB.Label lbl_Radius 
         Caption         =   "Radius"
         Height          =   255
         Left            =   4200
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.Label lbl_Speed 
         Caption         =   "Speed"
         Height          =   255
         Left            =   3360
         TabIndex        =   15
         Top             =   600
         Width           =   615
      End
   End
   Begin VB.TextBox txt_Name 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.CheckBox chk_Enabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3960
      TabIndex        =   10
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton btn_OK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label lbl_ActiveWin 
      Caption         =   " "
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   4080
      Width           =   4935
   End
   Begin VB.Label lbl_Move 
      Caption         =   "Move:"
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lbl_Button2 
      Caption         =   "Button action After move"
      Height          =   255
      Left            =   2760
      TabIndex        =   14
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lbl_Button1 
      Caption         =   "Button action Before move"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label lbl_Name 
      Caption         =   "Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_Act_Mouse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim scMove     As New oSuperCombo
Dim scButton1  As New oSuperCombo
Dim scButton2  As New oSuperCombo
Dim scRelative As New oSuperCombo

Dim X As Long
Dim y As Long

Public ActiveWinHwnd As Long
Public Mediane       As Boolean

Private Sub btn_Cancel_Click()

  Unload Me

End Sub

Private Sub btn_Ok_Click()

  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Name, txt_Name)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Enabled, IIf(chk_Enabled.Value = vbChecked, c_Val_Yes, c_Val_No))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_X, txt_X)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Y, txt_Y)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Move, scMove.CurrValue)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Speed, Val(txt_Speed))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Radius, Val(txt_Radius))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Button1, scButton1.CurrValue)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Button2, scButton2.CurrValue)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Relative, scRelative.CurrValue)
  m_Frm_SavePosition Me, g_CurrAct_Prm
  g_CurrAct_Ok = True
  Unload Me

End Sub

Private Sub btn_Position_Click()
    
  m_Pos_LocatePoint Me, True, False
  
End Sub

Private Sub lbl_ActiveWin_DblClick()

  m_Win_ShowActiveWinList 11

End Sub

Private Sub lst_Move_Click()

  Const c_Free = "free"

  Dim OkPoint   As Boolean
  Dim OkRadius  As Boolean
  Dim OkButton2 As Boolean
  Dim Button1   As String
  Dim Button2   As String
  
  Button1 = c_Free
  Button2 = c_Free
  
  Select Case scMove.CurrValue
  Case c_Move_Line
    frm_Point.Caption = "Target point"
    OkPoint = True
    OkButton2 = True
  Case c_Move_Circle
    frm_Point.Caption = "Center"
    OkPoint = True
    OkRadius = True
    OkButton2 = True
    Button1 = c_Button_Down
    Button2 = c_Button_Up
  End Select

  'Make visible or invisible
  frm_Point.Visible = OkPoint
  txt_Radius.Visible = OkRadius
  lbl_Radius.Visible = OkRadius
  lst_Button2.Visible = OkButton2
  lbl_Button2.Visible = OkButton2

  'Fixe or free action button
  
  If Button1 = c_Free Then
    lst_Button1.Enabled = True
  Else
    scButton1.SetCtrVal Button1
    lst_Button1.Enabled = False
  End If

  If Button2 = c_Free Then
    lst_Button2.Enabled = True
  Else
    scButton2.SetCtrVal Button2
    lst_Button2.Enabled = False
  End If

End Sub

Private Sub Form_Load()

  Dim X As String

  m_Win_StayOnTop Me.Hwnd, True
  
  With scMove
    .SetCtrObj lst_Move, True 'strange bug : when the frm is open for the second time, the combo box is empty but the oSuperCombo object is not.
    .AddItem c_Move_Line, c_Move_Line
    .AddItem c_Move_Circle, c_Move_Circle
    .AddItem c_Move_None, c_Move_None
  End With

  With scButton1
    .SetCtrObj lst_Button1, True
    .AddItem "", "None"
    .AddItem c_Button_Down, "Only Down"
    .AddItem c_Button_Up, "Only Up"
    .AddItem c_Button_Click, "Click"
    .AddItem c_Button_DblClick, "Double-click"
  End With

  With scButton2
    .SetCtrObj lst_Button2, True
    .AddItem "", "None"
    .AddItem c_Button_Down, "Only Down"
    .AddItem c_Button_Up, "Only Up"
    .AddItem c_Button_Click, "Click"
    .AddItem c_Button_DblClick, "Double-click"
  End With
  
  m_RelativeInfo_FeedLst scRelative, lst_Relative
  scRelative.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Relative), True
 
  'Retrieve Action's parameters
  '----------------------------
  
  m_Frm_LoadPosition Me, g_CurrAct_Prm
  
  txt_Name = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Name)
  chk_Enabled.Value = IIf(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Enabled) = c_Val_No, vbUnchecked, vbChecked)
  
  txt_X = "" & Val(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_X))
  txt_Y = "" & Val(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Y))
  txt_Speed = Val(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Speed))
  txt_Radius = Val(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Radius))
  
  'Can make bug of coded previously because value change
  scMove.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Move), True
  scButton1.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Button1), True
  scButton2.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Button2), True
  
End Sub

Private Sub lst_Relative_Click()

  m_RelativeInfo_Actualize Me, scRelative
  
End Sub
