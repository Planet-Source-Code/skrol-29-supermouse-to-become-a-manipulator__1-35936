VERSION 5.00
Begin VB.Form frm_Options 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3135
   ClientLeft      =   5910
   ClientTop       =   3300
   ClientWidth     =   4770
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox chk_Main_AutoQuit 
      Caption         =   "Quit after run if script opened with command line"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   4575
   End
   Begin VB.CheckBox chk_Mouse_RestorePos 
      Caption         =   "Protection against user moves"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   4575
   End
   Begin VB.CheckBox chk_Main_StayOnTop 
      Caption         =   "Main window stays on top when script is running"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4575
   End
   Begin VB.CheckBox chk_Main_RestorePos 
      Caption         =   "Remember the main window position when script is opened"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4575
   End
   Begin VB.CommandButton B_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton B_OK 
      Caption         =   "Ok"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"frm_Options.frx":0000
      ForeColor       =   &H80000011&
      Height          =   735
      Left            =   360
      TabIndex        =   6
      Top             =   1440
      Width           =   4215
   End
End
Attribute VB_Name = "frm_Options"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub B_Cancel_Click()
  Unload Me
End Sub

Private Sub B_OK_Click()

  m_Option_Set c_Opt_Main_StayOnTop, "", c_Prm_Enabled, IIf(chk_Main_StayOnTop = vbChecked, c_Val_Yes, c_Val_No)
  m_Option_Set c_Opt_Main_AutoQuit, "", c_Prm_Enabled, IIf(chk_Main_AutoQuit = vbChecked, c_Val_Yes, c_Val_No)
  m_Option_Set c_Opt_Main_RestorePos, "", c_Prm_Enabled, IIf(chk_Main_RestorePos = vbChecked, c_Val_Yes, c_Val_No)
  m_Option_Set c_Opt_Mouse_RestorePos, "", c_Prm_Enabled, IIf(chk_Mouse_RestorePos = vbChecked, c_Val_Yes, c_Val_No)

  Unload Me

End Sub


Private Sub Form_Load()

  chk_Main_StayOnTop = IIf(m_Option_Get(c_Opt_Main_StayOnTop, "", c_Prm_Enabled) <> c_Val_No, vbChecked, vbUnchecked)
  chk_Main_AutoQuit = IIf(m_Option_Get(c_Opt_Main_AutoQuit, "", c_Prm_Enabled) <> c_Val_No, vbChecked, vbUnchecked)
  chk_Main_RestorePos = IIf(m_Option_Get(c_Opt_Main_RestorePos, "", c_Prm_Enabled) <> c_Val_No, vbChecked, vbUnchecked)
  chk_Mouse_RestorePos = IIf(m_Option_Get(c_Opt_Mouse_RestorePos, "", c_Prm_Enabled) <> c_Val_No, vbChecked, vbUnchecked)

End Sub
