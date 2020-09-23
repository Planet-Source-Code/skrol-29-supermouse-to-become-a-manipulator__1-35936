VERSION 5.00
Begin VB.Form frm_Target_Point 
   Caption         =   "Form1"
   ClientHeight    =   855
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   1725
   LinkTopic       =   "Form1"
   ScaleHeight     =   855
   ScaleWidth      =   1725
   Begin VB.CommandButton btn_Ok 
      Caption         =   "ok"
      Default         =   -1  'True
      Height          =   195
      Left            =   840
      TabIndex        =   2
      Top             =   480
      Width           =   195
   End
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "cancel"
      Height          =   195
      Left            =   1200
      TabIndex        =   1
      Top             =   480
      Width           =   195
   End
   Begin VB.Timer Timer1 
      Left            =   1080
      Top             =   0
   End
   Begin VB.Line lh2 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   480
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Line lv2 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   240
      Y1              =   240
      Y2              =   480
   End
   Begin VB.Line lv1 
      BorderColor     =   &H000000FF&
      X1              =   240
      X2              =   240
      Y1              =   0
      Y2              =   240
   End
   Begin VB.Line lh1 
      BorderColor     =   &H000000FF&
      X1              =   0
      X2              =   240
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lbl_locate 
      Appearance      =   0  'Flat
      BackColor       =   &H000000FF&
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   255
   End
End
Attribute VB_Name = "frm_Target_Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This form enables to locate a point and/or color

Const c_Timer = 50

Dim RegionHwnd As Long
Dim OffsetX As Long
Dim OffsetY As Long

'Center of the cross
Dim CenterX As Long
Dim CenterY As Long
Dim CenterW As Long
Dim CenterH As Long

Dim ActWX    As Long
Dim ActWY    As Long
Dim ActWHwnd As Long
Dim MedianeX As Long

Private Sub p_StopLocating(Cancel As Boolean)

  Timer1.Interval = 0

  If Cancel Then
  
    With g_Pos_Frm
      If g_Pos_EditXY Then
        .txt_X.Text = g_Pos_SaveX
        .txt_Y.Text = g_Pos_SaveY
      End If
      If g_Pos_EditColor Then
        .txt_Color.Text = g_Pos_SaveColor
      End If
    End With
    
    g_Pos_Ok = False
    
  Else
    
    g_Pos_Ok = True
  
  End If

  Unload Me

End Sub

Private Sub btn_Cancel_Click()
  p_StopLocating True
End Sub

Private Sub btn_Ok_Click()
  p_StopLocating False
End Sub

Private Sub Form_Load()

  Dim WinW As Long
  Dim WinH As Long

  g_Pos_Ok = False

  m_Win_StayOnTop Me.Hwnd, True
  
  'Save data
  With g_Pos_Frm
  
    If g_Pos_EditXY Then
      g_Pos_SaveX = .txt_X.Text
      g_Pos_SaveY = .txt_Y.Text
    End If
    
    If g_Pos_EditColor Then
      g_Pos_SaveColor = .txt_Color.Text
    End If
    
    ActWHwnd = .ActiveWinHwnd
    If (ActWHwnd <> 0) Then m_Win_GetPosition ActWHwnd, ActWX, ActWY
    
    If .Mediane Then
      If (ActWHwnd = 0) Then
        MedianeX = (Screen.Width / Screen.TwipsPerPixelX) / 2
      Else
        m_Win_GetSize ActWHwnd, WinW, WinH
        MedianeX = WinW / 2
      End If
    End If
  
  End With
    
  'Arrange controls
  '----------------
  
  'Hide buttuns but keeps them enabled
  btn_Ok.Width = 0
  btn_Ok.Height = 0
  btn_Cancel.Width = 0
  btn_Cancel.Height = 0

  'Calculate the center of the cross
  CenterX = lv1.X1
  CenterY = lh1.Y1
  
  'Calculate the size of the central pixels of the cross
  CenterW = Me.ScaleX(1, vbPixels, vbTwips)
  CenterH = Me.ScaleY(1, vbPixels, vbTwips)
  
  'Moves controls
  lv1.X1 = CenterX
  lv1.X2 = CenterX
  lv1.Y1 = 0
  lv1.Y2 = CenterY - CenterH / 2
  
  lv2.X1 = CenterX
  lv2.X2 = CenterX
  lv2.Y1 = CenterY + CenterH + CenterH / 2
  lv2.Y2 = 2 * CenterY

  lh1.X1 = 0
  lh1.X2 = CenterX - CenterW / 2
  lh1.Y1 = CenterY
  lh1.Y2 = CenterY
  
  lh2.X1 = CenterX + CenterW / 2
  lh2.X2 = 2 * CenterX
  lh2.Y1 = CenterY
  lh2.Y2 = CenterY

  'Now we move the label. We use a label because it is clickable.
  lbl_locate.Left = CenterX
  lbl_locate.Top = CenterY + CenterH / 2
  lbl_locate.Width = CenterW
  lbl_locate.Height = CenterH
  m_Frm_GetEdges Me, OffsetX, OffsetY
  OffsetX = OffsetX + lbl_locate.Left
  OffsetY = OffsetY + lbl_locate.Top

  'Make the form transparent
  m_Mouse_Show False
  RegionHwnd = m_Frm_ChangeEffect(Me, 2)
  
  'Posionate the mouse on the position
  If (Val(g_Pos_SaveX) <> 0) Or (Val(g_Pos_SaveY) <> 0) Then
    m_Mouse_Move Val(g_Pos_SaveX) + ActWX + MedianeX, Val(g_Pos_SaveY) + ActWY
  End If
  
  'Posionate the form once before to get visible
  'Timer1_Timer
  
  'Start the locating
  Timer1.Interval = c_Timer
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  Me.Visible = False
  m_Mouse_Show True
  m_Frm_CloseEffect Me, RegionHwnd
End Sub

Private Sub lbl_locate_Click()
  p_StopLocating False
End Sub

Private Sub Timer1_Timer()

  Dim X As Long
  Dim Y As Long
  Dim c As Long
  
  m_Mouse_GetCurrPosPixels X, Y
  
  'Place the current form
  Me.Left = Me.ScaleX(X, vbPixels, vbTwips) - OffsetX
  Me.Top = Me.ScaleY(Y, vbPixels, vbTwips) - OffsetY
  
  'Calculate target position
  Y = Y - 1
  With g_Pos_Frm
  
    'Position
    If g_Pos_EditXY Then
      .txt_X.Text = "" & (X - ActWX - MedianeX)
      .txt_Y.Text = "" & (Y - ActWY)
    End If
    
    'Color
    If g_Pos_EditColor Then
      c = m_Pixel_GetColor(X, Y)
      .txt_Color.Text = Hex$(c)
    End If
    
  End With

End Sub
