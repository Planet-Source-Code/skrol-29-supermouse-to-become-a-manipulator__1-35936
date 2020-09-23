VERSION 5.00
Begin VB.Form frm_Act_Loop 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Loop"
   ClientHeight    =   1920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3510
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1920
   ScaleWidth      =   3510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chk_Enabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      Height          =   255
      Left            =   2400
      TabIndex        =   1
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   1440
      Width           =   1095
   End
   Begin VB.CommandButton btn_Ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1095
   End
   Begin VB.ComboBox lst_Unit 
      Height          =   315
      ItemData        =   "frm_Act_Loop.frx":0000
      Left            =   720
      List            =   "frm_Act_Loop.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txt_Nbr 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Text            =   "1"
      Top             =   960
      Width           =   495
   End
   Begin VB.TextBox txt_Name 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
   Begin VB.Label lbl_Unit 
      Caption         =   "Unit"
      Height          =   255
      Left            =   720
      TabIndex        =   8
      Top             =   720
      Width           =   975
   End
   Begin VB.Label lbl_Nbr 
      Caption         =   "Nbr"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lbl_Name 
      Caption         =   "Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_Act_Loop"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim scUnit As New oSuperCombo

Private Sub btn_Cancel_Click()
  Unload Me
End Sub

Private Sub btn_Ok_Click()

  If Val(txt_Nbr) = 0 Then
    MsgBox "You must enter a positive value.", vbInformation
    Exit Sub
  End If

  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Name, txt_Name)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Enabled, IIf(chk_Enabled.Value = vbChecked, c_Val_Yes, c_Val_No))
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Nbr, txt_Nbr)
  g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Unit, scUnit.CurrValue)
  g_CurrAct_Ok = True

  Unload Me

End Sub

Private Sub Form_Load()

  'Controls initialisation
  With scUnit
    .SetCtrObj lst_Unit, True
    .AddItem c_Unit_Occ, c_Unit_Occ
    .AddItem c_Unit_Sec, c_Unit_Sec
    .AddItem c_Unit_Min, c_Unit_Min
    .AddItem c_Unit_Hour, c_Unit_Hour
  End With
  
  'Retrieve Action's parameters
  '----------------------------
  
  txt_Name = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Name)
  chk_Enabled.Value = IIf(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Enabled) = c_Val_No, vbUnchecked, vbChecked)
  
  txt_Nbr = "" & Val(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Nbr))
  If Val(txt_Nbr) = 0 Then
    txt_Nbr = "1"
  End If
  
  scUnit.SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Unit), True

End Sub
