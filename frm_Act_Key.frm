VERSION 5.00
Begin VB.Form frm_Act_Misc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Misc"
   ClientHeight    =   2535
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox lst_Value 
      Height          =   315
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.ComboBox lst_Option 
      Height          =   315
      Left            =   1080
      TabIndex        =   3
      Top             =   1320
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txt_Value 
      Height          =   315
      Left            =   120
      TabIndex        =   2
      Top             =   930
      Width           =   4335
   End
   Begin VB.CommandButton B_OK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1095
   End
   Begin VB.CommandButton B_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
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
   Begin VB.Label lbl_Option 
      Caption         =   "Option:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label lbl_Help 
      Caption         =   "help"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   4335
   End
   Begin VB.Label lbl_Value 
      Caption         =   "Value"
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lbl_Name 
      Caption         =   "Caption"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "frm_Act_Misc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'{BACKSPACE}{BS}{BKSP}
'{BREAK}
'{CAPSLOCK}
'{DELETE}{DEL}
'{DOWN}
'{END}
'{ENTER}
'{ESC}
'{HELP}
'{HOME}
'{INSERT}{INS}
'{LEFT}
'{NUMLOCK}
'{PGDN}
'{PGUP}
'{PRTSC}
'{RIGHT}
'{SCROLLLOCK}

Dim ActType As String
Dim scOption As New oSuperCombo

Private Sub B_Cancel_Click()
  Unload Me
End Sub

Private Sub B_OK_Click()

  If txt_Name.Visible Then g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Name, txt_Name)
  If chk_Enabled.Visible Then g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Enabled, IIf(chk_Enabled.Value = vbChecked, c_Val_Yes, c_Val_No))
  If lst_Option.Visible Then g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Mode, scOption.CurrValue)
  If txt_Value.Visible Then g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Value, txt_Value)
  If lst_Value.Visible Then g_CurrAct_Prm = m_ItemLst_Set(g_CurrAct_Prm, c_Prm_Value, lst_Value)
  
  g_CurrAct_Ok = True

  Unload Me

End Sub

Private Sub Form_Load()

  Dim i As Long

  ActType = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Type)
  
  Select Case ActType
  Case c_Act_Comment, c_Act_Label
    txt_Name.Visible = False
    lbl_Name.Visible = False
    chk_Enabled.Visible = False
  Case c_Act_Message
    txt_Name.Visible = False
    lbl_Name.Visible = False
  Case c_Act_Goto
    txt_Name.Visible = False
    lbl_Name.Visible = False
    txt_Value.Visible = False
  Case Else 'Keys / Execute
  End Select
  
  'Special list for Execute Action
  If ActType = c_Act_Execute Then
    With lbl_Option
      .Caption = "Win Style:"
      .Visible = True
    End With
    With scOption
      .SetCtrObj lst_Option, True
      .AddItem vbNormalFocus, "Normal Focus"
      .AddItem vbNormalNoFocus, "Normal No Focus"
      .AddItem vbMaximizedFocus, "Maximized Focus"
      .AddItem vbMinimizedFocus, "Minimized Focus"
      .AddItem vbMinimizedNoFocus, "Minimized No Focus"
      .AddItem vbHide, "Hide"
      .SetCtrVal m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Mode), vbNormalFocus
    End With
    lst_Option.Visible = True
  End If

  'Special list for Goto Action
  If ActType = c_Act_Goto Then
    m_Ctr_FeedWithLabels lst_Value
    lst_Value.Visible = True
  End If
  
  'Warning: the following controls can be unvisible at that point even if they have been set to Visible
  txt_Name = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Name)
  chk_Enabled.Value = IIf(m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Enabled) = c_Val_No, vbUnchecked, vbChecked)
  txt_Value = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Value)
  lst_Value = m_ItemLst_Get(g_CurrAct_Prm, c_Prm_Value)
  
End Sub
