Attribute VB_Name = "mod_TransparentFrm"
Option Explicit

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal Hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Const RGN_OR = 2 'creates the union of combined regions
Private Const RGN_DIFF = 4 'creates the intersection of combined regions
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Public Sub m_Frm_CloseEffect(Frm As Form, RegionHwnd As Long)

  SetWindowRgn Frm.Hwnd, 0, False
  DeleteObject RegionHwnd

End Sub

Public Function m_Frm_ChangeEffect(Frm As Form, inEffect As Integer) As Long
'Apply effect and returns the Hwnd of the region wich has to be destroy before quit the form
'inEffect=0 => normal
'inEffect=1 => transparent with visible frame
'inEffect=2 => transparent all
    
    Dim w As Single, h As Single
    Dim Edge As Single
    Dim TopEdge As Single
    Dim X1, X2, Y1, Y2 As Long
    Dim i As Integer
    Dim r As Long
    Dim Outer As Long
    Dim Inner As Long
    Dim RegionHwnd As Long
    Dim CtrOk   As Boolean
    
    Frm.ScaleMode = vbPixels
    
       ' Put width/height in same denomination of scalewidth/scaleheight
    w = Frm.ScaleX(Frm.Width, vbTwips, vbPixels)
    h = Frm.ScaleY(Frm.Height, vbTwips, vbPixels)
    
    If inEffect = 0 Then
      RegionHwnd = CreateRectRgn(0, 0, w, h)
      SetWindowRgn Frm.Hwnd, RegionHwnd, True
      m_Frm_ChangeEffect = RegionHwnd
      Exit Function
    End If
    
    RegionHwnd = CreateRectRgn(0, 0, 0, 0)
    ' Frame edges measurement
    Edge = (w - Frm.ScaleWidth) / 2
    TopEdge = h - Edge - Frm.ScaleHeight
    
    ' Get frame
    If inEffect = 1 Then
      Outer = CreateRectRgn(0, 0, w, h)
      Inner = CreateRectRgn(Edge, TopEdge, w - Edge, h - Edge)
      CombineRgn RegionHwnd, Outer, Inner, RGN_DIFF
    End If
    
    ' Combine regions of controls on form
    For i = 0 To Frm.Controls.Count - 1
      With Frm.Controls(i)
        CtrOk = False
        On Error Resume Next
        CtrOk = (.X2 >= 0)
        On Error GoTo 0
        If CtrOk Then
          If .Visible = True Then
            X1 = Frm.ScaleX(.X1, Frm.ScaleMode, vbPixels) + Edge
            X2 = Frm.ScaleX(.X2, Frm.ScaleMode, vbPixels) + Edge
            Y1 = Frm.ScaleY(.Y1, Frm.ScaleMode, vbPixels) + TopEdge
            Y2 = Frm.ScaleY(.Y2, Frm.ScaleMode, vbPixels) + TopEdge
            If X1 = X2 Then X2 = X2 + .BorderWidth
            If Y1 = Y2 Then Y2 = Y2 + .BorderWidth
          Else
            CtrOk = False
          End If
        Else
          On Error Resume Next
          CtrOk = (.Width >= 0)
          On Error GoTo 0
          If CtrOk Then
            If .Visible = True Then
              X1 = Frm.ScaleX(.Left, Frm.ScaleMode, vbPixels) + Edge
              X2 = X1 + Frm.ScaleX(.Width, Frm.ScaleMode, vbPixels)
              Y1 = Frm.ScaleY(.Top, Frm.ScaleMode, vbPixels) + TopEdge
              Y2 = Y1 + Frm.ScaleY(.Height, Frm.ScaleMode, vbPixels)
            Else
              CtrOk = False
            End If
          End If
        End If
        If CtrOk Then
          r = CreateRectRgn(X1, Y1, X2, Y2)
          CombineRgn RegionHwnd, r, RegionHwnd, RGN_OR
        End If
      End With
    Next
    
    ' We allow toggle
    SetWindowRgn Frm.Hwnd, RegionHwnd, True
    
    m_Frm_ChangeEffect = RegionHwnd
    
End Function

Public Sub m_Frm_GetEdges(Frm As Form, ByRef EdgeInTwips As Long, TopEdgeInTwips As Long)

  Frm.ScaleMode = vbTwips
  EdgeInTwips = (Frm.Width - Frm.ScaleWidth) / 2
  TopEdgeInTwips = Frm.Height - Frm.ScaleHeight - EdgeInTwips

End Sub
