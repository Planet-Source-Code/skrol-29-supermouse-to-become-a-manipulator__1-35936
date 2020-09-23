VERSION 5.00
Begin VB.Form frm_Act_Wait 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Wait"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chk_WaitColor 
      Caption         =   "Wait for a pixel to takes a specific color"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Value           =   1  'Checked
      Width           =   4455
   End
   Begin VB.CheckBox chk_WaitTime 
      Caption         =   "Wait for a duration"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   3120
      Width           =   2655
   End
   Begin VB.Frame frm_Time 
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ComboBox lst_Unit 
         Height          =   315
         ItemData        =   "frm_Act_Wait.frx":0000
         Left            =   960
         List            =   "frm_Act_Wait.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txt_Nbr 
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Text            =   "1"
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lbl_Unit 
         Caption         =   "Unit"
         Height          =   255
         Left            =   960
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_Nbr 
         Caption         =   "Nbr"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame frm_Color 
      Height          =   1815
      Left            =   120
      TabIndex        =   14
      Top             =   1080
      Width           =   4455
      Begin VB.ComboBox lst_Relative 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   240
         Width           =   2175
      End
      Begin VB.CommandButton btn_Position 
         Caption         =   "Locate"
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   840
         Width           =   1095
      End
      Begin VB.TextBox txt_Y 
         Height          =   315
         Left            =   960
         TabIndex        =   5
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.TextBox txt_X 
         Height          =   315
         Left            =   120
         TabIndex        =   3
         Text            =   "0"
         Top             =   840
         Width           =   735
      End
      Begin VB.CheckBox chk_ChangeColor 
         Caption         =   "Edit"
         Height          =   255
         Left            =   1800
         TabIndex        =   7
         Top             =   1200
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CheckBox chk_ChangeXY 
         Caption         =   "Edit"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.TextBox txt_Color 
         Height          =   315
         Left            =   1800
         TabIndex        =   6
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lbl_Relative 
         Caption         =   "Relative to:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_LocInfo 
         Caption         =   "Click or press [Enter] to validate, or [Esc] to cancel."
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   120
         TabIndex        =   23
         Top             =   1560
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.Shape shp_Color 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         Height          =   315
         Left            =   2760
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbl_Color 
         Caption         =   "Color Id"
         Height          =   255
         Left            =   1800
         TabIndex        =   17
         Top             =   600
         Width           =   855
      End
      Begin VB.Label lbl_X 
         Caption         =   "X"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl_Y 
         Caption         =   "Y"
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   600
         Width           =   735
      End
   End
   Begin VB.TextBox txt_Name 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2535
   End
   Begin VB.CheckBox chk_Enabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3480
      TabIndex        =   12
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton btn_Ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Label lbl_ActiveWin 
      Caption         =   " "
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   4920
      Width           =   4455
   End
   Begin VB.Label lbl_Name 
      Caption         =   "Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Act_Wait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim scUnit As New oSuperCombo
Dim scRelative As New oSuperCombo

Public ActiveWinHwnd As Long
Public Mediane As Boolean

Private Sub btn_Position_Click()

  m_Pos_LocatePoint Me, (chk_ChangeXY.Value = vbChecked), (chk_ChangeColor.Value = vbChecked)
  
End Sub



Private Sub btn_Cancel_Click()

  Unload Me

End Sub

Private Sub btn_Ok_Click()

  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Name, txt_Name)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Enabled, IIf(chk_Enabled = vbChecked, c_Val_Yes, c_Val_No))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Nbr, IIf(chk_WaitTime = vbChecked, txt_Nbr, "0"))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Unit, scUnit.CurrValue)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Color, IIf(chk_WaitColor = vbChecked, txt_Color, ""))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_X, IIf(chk_WaitColor = vbChecked, txt_X, "0"))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Y, IIf(chk_WaitColor = vbChecked, txt_Y, "0"))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Relative, scRelative.CurrValue)
  m_Frm_SavePosition Me, g_CurrAct_Prm
  g_CurrAct_Ok = True
  Unload Me

End Sub

Private Sub chk_WaitColor_Click()
  
  frm_Color.Visible = (chk_WaitColor.Value = vbChecked)
  
End Sub

Private Sub chk_WaitTime_Click()

  frm_Time.Visible = (chk_WaitTime.Value = vbChecked)

End Sub

Private Sub lst_Relative_Click()

  m_RelativeInfo_Actualize Me, scRelative
  
End Sub

Private Sub txt_Color_Change()

  shp_Color.BackColor = m_HexToDec(txt_Color.Text)

End Sub

Private Sub Form_Load()

  m_Win_StayOnTop Me.Hwnd, True

  With scUnit
    .SetCtrObj lst_Unit, True
    .AddItem c_Unit_Ms, c_Unit_Ms
    .AddItem c_Unit_Sec, c_Unit_Sec
    .AddItem c_Unit_Min, c_Unit_Min
    .AddItem c_Unit_Hour, c_Unit_Hour
  End With

  m_RelativeInfo_FeedLst scRelative, lst_Relative
  scRelative.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Relative), True

  'Retrieve Action's parameters
  '----------------------------
  
  m_Frm_LoadPosition Me, g_CurrAct_Prm
  
  txt_Name = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Name)
  chk_Enabled.Value = IIf(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Enabled) = c_Val_No, vbUnchecked, vbChecked)
  
  txt_Nbr = "" & Val(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Nbr))
  txt_X = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_X)
  txt_Y = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Y)
  txt_Color = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Color)

  'Valeurs par défaut
  If Val(txt_Nbr) = 0 Then txt_Nbr = "0"
  If Val(txt_X) = 0 Then txt_X = "0"
  If Val(txt_Y) = 0 Then txt_Y = "0"

  chk_WaitTime = IIf(txt_Nbr = "0", vbUnchecked, vbChecked)
  chk_WaitColor = IIf(txt_Color = "", vbUnchecked, vbChecked)

  scUnit.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Unit), True

End Sub

Private Sub txt_Nbr_Change()
  chk_WaitTime = IIf(Val(txt_Nbr) > 0, vbChecked, vbUnchecked)
End Sub
