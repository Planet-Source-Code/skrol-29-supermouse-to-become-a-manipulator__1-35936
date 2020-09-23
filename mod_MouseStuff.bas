Attribute VB_Name = "mod_MouseStuff"
Option Explicit
Option Compare Text

Const c_tStep = 40 'millisec, is the period od each step for the line-move and the circle-move
Const c_Pi = 3.41592653589793

Public UserPosX As Long
Public UserPosY As Long
Public SmsPosX  As Long
Public SmsPosY  As Long
Public SmsBtnUp As Boolean

Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Declare Sub api_MouseEvent Lib "user32" Alias "mouse_event" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cbuttons As Long, ByVal dwExtraInfo As Long)

Private Type POINTAPI
    X As Long
    y As Long
End Type
Private Declare Function api_GetCursorPos Lib "user32" Alias "GetCursorPos" (lpPoint As POINTAPI) As Long

Private Const c_Mickey = 65535

Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

Public Sub m_Mouse_Show(Display As Boolean)
  ShowCursor IIf(Display, -1, 0)
End Sub

Public Function p_Mouse_XInMickey(XInPixels As Long) As Long
  Dim ScrInPixel As Long
  ScrInPixel = Screen.Width \ Screen.TwipsPerPixelX
  p_Mouse_XInMickey = Round(XInPixels * (c_Mickey / ScrInPixel))
End Function

Public Function p_Mouse_YInMickey(YInPixels As Long) As Long
  Dim ScrInPixel As Long
  ScrInPixel = Screen.Height \ Screen.TwipsPerPixelY
  p_Mouse_YInMickey = Round(YInPixels * (c_Mickey / ScrInPixel))
End Function

Public Sub m_Mouse_GetCurrPosPixels(ByRef XInPixels As Long, ByRef YInPixels As Long)

  Dim Pt As POINTAPI
  
  'The api returns current cursor position in Pixels
  api_GetCursorPos Pt
  
  XInPixels = Pt.X
  YInPixels = Pt.y
  
End Sub

Public Sub m_Mouse_Action(Move As String, XInPixels As Long, YInPixels As Long, Button1 As String, Button2 As String, Optional Speed As Long, Optional Radius As Long)
  
  Dim n    As Long
  Dim NMax As Single
  
  Dim X0   As Long
  Dim Y0   As Long
  Dim Xn   As Long
  Dim Yn   As Long
  
  'Save user postion and restore previous SupermouseScript position
  If g_RestorePos Then
    If (Not SmsBtnUp) Then
      m_Mouse_GetCurrPosPixels UserPosX, UserPosY
        m_Mouse_Move SmsPosX, SmsPosY
    End If
  End If
  
  'Move the cursor
  Select Case Move
  
  Case c_Move_Line
    
    'button activation before the move
    p_Mouse_Button Button1
    
    If Speed <= 0 Then
    
      m_Mouse_Move XInPixels, YInPixels
    
    Else 'Simulate motion on a line from the current point to the target point.
      
      'Get current position
      m_Mouse_GetCurrPosPixels X0, Y0
      
      'Calculate the number of iteration to perfom
      NMax = 0
      p_Move_Line X0, Y0, XInPixels, YInPixels, Speed, NMax, 0, 0, 0
      
      If NMax > 0 Then
        m_Wait c_tStep
        For n = 1 To (NMax - 1)
          p_Move_Line X0, Y0, XInPixels, YInPixels, Speed, NMax, n, Xn, Yn
          m_Mouse_Move Xn, Yn
          m_Wait c_tStep
        Next n
      End If
      
      'In any way, the last move is stricly on the target
      m_Mouse_Move XInPixels, YInPixels
    
    End If
    
    'button activation after the move
    p_Mouse_Button Button2
    
  Case c_Move_Circle
  
    'Simulate motion on a circle aroud the target point.
  
    'Calculate the number of iteration to perfom
    p_Move_Circle XInPixels, YInPixels, Radius, Speed, NMax, 0, 0, 0
    
    If NMax > 0 Then
      For n = 0 To NMax
        p_Move_Circle XInPixels, YInPixels, Radius, Speed, NMax, n, Xn, Yn
        m_Mouse_Move Xn, Yn
        If n = 0 Then
          p_Mouse_Button Button1
        End If
        m_Wait c_tStep
      Next n
    End If
      
    p_Mouse_Button Button2
    
  End Select
  
  'Save SupermouseScript postion and restore previous user position
  If g_RestorePos Then
    If (Not SmsBtnUp) Then
      m_Mouse_GetCurrPosPixels SmsPosX, SmsPosY
      m_Mouse_Move UserPosX, UserPosY
    End If
  End If
  
End Sub

Public Sub m_Mouse_Move(XInTwips As Long, YInTwips As Long)

  If (XInTwips > 0) And (YInTwips > 0) Then
    api_MouseEvent MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_MOVE, p_Mouse_XInMickey(XInTwips), p_Mouse_YInMickey(YInTwips), 0, 0
  End If

End Sub

Private Sub p_Mouse_Button(BtnStr As String)
'This procedure reads the BtnStr string and performs all buttons actions specified

  Dim X    As String
  Dim Flag As Long
    
  X = BtnStr
  
  Do Until X = vbNullString
    
    Flag = 0
    
    'Check if the first action is WAIT
    If InStr(X, c_Button_Wait) = 1 Then
      X = Mid$(X, 1 + Len(c_Button_Wait))
      m_Wait (500)
    End If
    
    'Check if the first action is DOWN
    If InStr(X, c_Button_Down) = 1 Then
      X = Mid$(X, 1 + Len(c_Button_Down))
      Flag = Flag Or MOUSEEVENTF_LEFTDOWN
    End If
    
    'Check if the first action is UP
    If InStr(X, c_Button_Up) = 1 Then
      X = Mid$(X, 1 + Len(c_Button_Up))
      Flag = Flag Or MOUSEEVENTF_LEFTUP
    End If
    
    If Flag <> 0 Then
      'Performs the button action
      api_MouseEvent Flag, 0, 0, 0, 0
    Else
      'Delete the first char
      X = Mid$(X, 2)
    End If
  
  Loop

End Sub

Private Sub p_Move_Line(ByRef X0 As Long, ByRef Y0 As Long, ByRef X1 As Long, ByRef Y1 As Long, ByRef Speed As Long, ByRef NMax As Single, ByRef n As Long, ByRef Xn As Long, ByRef Yn As Long)
'This procedure calculates NMax, Xn and Yn that are the Number of iteration and the coordonate of a point for a move on a line.

  If NMax = 0 Then
    'Calculates NMax
    NMax = Sqr((X1 - X0) ^ 2 + (Y1 - Y0) ^ 2) / (0.01 * Speed * c_tStep)
  Else
    'Calcultaes Xn and Yn
    Xn = Round(X0 + (X1 - X0) * n / NMax)
    Yn = Round(Y0 + (Y1 - Y0) * n / NMax)
  End If

End Sub

Private Sub p_Move_Circle(ByRef X0 As Long, ByRef Y0 As Long, ByRef Radius As Long, ByRef Speed As Long, ByRef NMax As Single, ByRef n As Long, ByRef Xn As Long, ByRef Yn As Long)
'This procedure calculates NMax, Xn and Yn that are the Number of iteration and the coordonate of a point for a move on a line.

  If NMax = 0 Then
    'Calculates NMax
    NMax = 2 * c_Pi * Radius / (0.01 * Speed * c_tStep)
  Else
    'Calculates Xn and Yn
    Xn = Round(X0 + Radius * Cos(2 * c_Pi * n / NMax))
    Yn = Round(Y0 + Radius * Sin(2 * c_Pi * n / NMax))
  End If

End Sub
