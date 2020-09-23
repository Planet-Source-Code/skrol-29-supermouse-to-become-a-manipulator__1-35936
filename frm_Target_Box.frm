VERSION 5.00
Begin VB.Form frm_Target_Box 
   Caption         =   "Form1"
   ClientHeight    =   1635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1560
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   1560
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   195
      Left            =   960
      TabIndex        =   3
      Top             =   720
      Width           =   75
   End
   Begin VB.CommandButton btn_Ok 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   720
      Width           =   75
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   1080
      Top             =   0
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      X1              =   600
      X2              =   600
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   600
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   120
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      X1              =   120
      X2              =   600
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label lbl_LineV 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   120
      Width           =   15
   End
   Begin VB.Label lbl_LineH 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      Caption         =   "Label1"
      ForeColor       =   &H80000008&
      Height          =   15
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
End
Attribute VB_Name = "frm_Target_Box"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RegionHwnd As Long

Dim OffsetX As Long
Dim OffsetY As Long

Private Sub btn_Cancel_Click()
  Unload Me
End Sub

Private Sub btn_Ok_Click()
  Unload Me
End Sub

Private Sub Form_Load()

  'Calculate offsets
  m_Frm_GetEdges Me, OffsetX, OffsetY
  OffsetX = OffsetX + lbl_LineV.Left + (lbl_LineV.Width / 2)
  OffsetY = OffsetY + lbl_LineH.Top + (lbl_LineH.Height / 2)

  'Hide the mouse pointer
  m_Mouse_Show False

  'Make the form transparent
  btn_Ok.Width = 0
  btn_Ok.Height = 0
  btn_Cancel.Width = 0
  btn_Cancel.Height = 0
  RegionHwnd = m_Frm_ChangeEffect(Me, 2)
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  m_Mouse_Show True
  m_Frm_CloseEffect Me, RegionHwnd
End Sub

Private Sub lbl_LineH_Click()
  Unload Me
End Sub

Private Sub lbl_LineV_Click()
  Unload Me
End Sub

Private Sub Timer1_Timer()

  Dim x As Long
  Dim y As Long
  
  m_Mouse_GetCurrPosPixels x, y
  Me.Left = Me.ScaleX(x, vbPixels, vbTwips) - OffsetX
  Me.Top = Me.ScaleY(y, vbPixels, vbTwips) - OffsetY

End Sub
