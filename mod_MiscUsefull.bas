Attribute VB_Name = "mod_MiscUsefull"
Option Explicit
Option Compare Text

'Returns the number of milli-seconds until Windows stared
Public Declare Function api_GetTickCount Lib "kernel32" Alias "GetTickCount" () As Long

Type t_OuvrirFichier
    lTailleStruct               As Long
    hwndPropriétaire            As Long
    hInstance                   As Long
    lpstrFiltre                 As String
    lpstrFiltrePersonnalisé     As Long
    nFiltrePersonMax            As Long
    nIndexFiltre                As Long
    lpstrFichier                As String
    nFichierMax                 As Long
    lpstrTitreFichier           As String
    nTitreFichierMax            As Long
    lpstrRépInitial             As String
    lpstrTitre                  As String
    indicateurs                 As Long
    nPartieFichier              As Integer
    nExtensionFichier           As Integer
    lpstrExtDéf                 As String
    lDonnéesClient              As Long
    lpfnCrochet                 As Long
    lpNomModèle                 As Long
End Type

'File dialog box
Declare Function api_GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As t_OuvrirFichier) As Boolean
Declare Function api_GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As t_OuvrirFichier) As Boolean

'Pixel
Declare Function api_CreateDC Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, ByVal lpDeviceName As String, ByVal lpOutput As String, lpInitData As Any) As Long
Declare Function api_DeleteDC Lib "gdi32" Alias "DeleteDC" (ByVal hdc As Long) As Long
Declare Function api_GetPixel Lib "gdi32" Alias "GetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

'Stay on top
Declare Function api_SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal Hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Public Const HWND_BOTTOM = 1
Public Const HWND_BROADCAST = &HFFFF&
Public Const HWND_DESKTOP = 0
Public Const HWND_NOTOPMOST = -2
Public Const HWND_TOP = 0
Public Const HWND_TOPMOST = -1

Public Const SWP_FRAMECHANGED = &H20
Public Const SWP_HIDEWINDOW = &H80
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_NOCOPYBITS = &H100
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOOWNERZORDER = &H200
Public Const SWP_NOREDRAW = &H8
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOZORDER = &H4
Public Const SWP_SHOWWINDOW = &H40
Public Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Public Const SWP_DRAWFRAME = SWP_FRAMECHANGED

Declare Function api_IsWindowVisible Lib "user32" Alias "IsWindowVisible" (ByVal Hwnd As Long) As Boolean

'----------------- Window stuff

Private Type RECT
  Left   As Long
  Top    As Long
  Right  As Long
  Bottom As Long
End Type

Private Type WINDOWINFO
    cbSize As Long
    rcWindow As RECT
    rcClient As RECT
    dwStyle As Long
    dwExStyle As Long
    dwWindowStatus As Long
    cxWindowBorders As Long
    cyWindowBorders As Long
    atomWindowType As Integer
    wCreatorVersion As Long
End Type

Private Declare Function api_GetActiveWindow Lib "user32" Alias "GetActiveWindow" () As Long
Private Declare Function api_GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal Hwnd As Long) As Long
Private Declare Function api_GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function api_GetWindowInfo Lib "user32" Alias "GetWindowInfo" (ByVal Hwnd As Long, Pwi As WINDOWINFO) As Long
Private Declare Function api_EnumWindows Lib "user32" Alias "EnumWindows" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long

'Window style (used by GetWindowInfo)
Public Const WS_OVERLAPPED = &H0         'overlapped window with a title bar and a border.
Public Const WS_POPUP = &H80000000
Public Const WS_CHILD = &H40000000
Public Const WS_MINIMIZE = &H20000000    'the window is currently minimized
Public Const WS_VISIBLE = &H10000000
Public Const WS_DISABLED = &H8000000
Public Const WS_CLIPSIBLINGS = &H4000000
Public Const WS_CLIPCHILDREN = &H2000000
Public Const WS_MAXIMIZE = &H1000000     'the window is currently maximized
Public Const WS_CAPTION = &HC00000       'window that has a title bar (automatically includes the WS_BORDER if no other border style is specified).
Public Const WS_BORDER = &H800000        'window with a thin-line border.
Public Const WS_DLGFRAME = &H400000      'window that has a border of a style typically used with dialog boxes. The border is not resizeable.
Public Const WS_VSCROLL = &H200000
Public Const WS_HSCROLL = &H100000
Public Const WS_SYSMENU = &H80000        'window with a window menu on its title bar. The WS_CAPTION style must also be specified. Allows WS_MAXIMIZEBOX and WS_MINIMIZEBOX styles.
Public Const WS_THICKFRAME = &H40000     'window with a sizing border.
Public Const WS_GROUP = &H20000
Public Const WS_TABSTOP = &H10000
Public Const WS_MINIMIZEBOX = &H20000    'window with a maximize button. The WS_SYSMENU style must also be specified.
Public Const WS_MAXIMIZEBOX = &H10000    'window with a minimize button. The WS_SYSMENU style must also be specified.

'Extended Window style (used by GetWindowInfo)
Public Const WS_EX_DLGMODALFRAME = &H1     'window with a double border. not resizeable.
Public Const WS_EX_NOPARENTNOTIFY = &H4
Public Const WS_EX_TOPMOST = &H8           'Window is on top of the normal application windows. Can be overlapped only by other top most windows.
Public Const WS_EX_ACCEPTFILES = &H10
Public Const WS_EX_TRANSPARENT = &H20
Public Const WS_EX_MDICHILD = &H40
Public Const WS_EX_TOOLWINDOW = &H80       'Window with a thin caption. Does not appear in the taskbar or in the Alt-Tab palette. WS_CAPTION also must be specified.
Public Const WS_EX_WINDOWEDGE = &H100      'window with a raised edge border.
Public Const WS_EX_CLIENTEDGE = &H200      'window with a sunken edge.
Public Const WS_EX_CONTEXTHELP = &H400
Public Const WS_EX_RIGHT = &H1000          '"right-aligned" window. see comments for the WS_EX_LEFTSCROLLBAR.
Public Const WS_EX_LEFT = &H0              'left-aligned texts and other controls. default.
Public Const WS_EX_RTLREADING = &H2000
Public Const WS_EX_LTRREADING = &H0        'right-to-left alignment of the window text. See comments for the WS_EX_LEFTSCROLLBAR.
Public Const WS_EX_LEFTSCROLLBAR = &H4000  'For language with right-to-left default flowing (Hebrew, Arabic and so on). Scrollbar appears on the left side of the window. Takes effect only if the system supports some of these languages.
Public Const WS_EX_RIGHTSCROLLBAR = &H0    '"right-aligned" window. see comments for the WS_EX_LEFTSCROLLBAR.
Public Const WS_EX_CONTROLPARENT = &H10000
Public Const WS_EX_STATICEDGE = &H20000    'window with a three-dimensional border.
Public Const WS_EX_APPWINDOW = &H40000     'window will be shown onto the taskbar when visible.
Public Const WS_EX_OVERLAPPEDWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_CLIENTEDGE)
Public Const WS_EX_PALETTEWINDOW = (WS_EX_WINDOWEDGE Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST)
                                           'Combines the WS_EX_WINDOWEDGE, WS_EX_TOOLWINDOW, and WS_EX_TOPMOST.

'Variables for the Window Sreach
Dim WinNbr    As Integer     'Number of window to search
Dim WinZHwnd  As Long        'Forbidden window Handle

Dim WinXNum   As Integer     'Current window number
Dim WinXHwnd  As Long        'Current window Handle
Dim WinXInfo  As WINDOWINFO  'Current window info

'---------------

Public Function m_CheckFlag(Parameters As Long, Flag As Long) As Boolean
  m_CheckFlag = ((Parameters And Flag) = Flag)
End Function

Public Function m_Win_GetTitle(Hwnd As Long) As String
'Returns the title for a specific window
  
  Dim Length  As Long
  Dim Caption As String

  Length = api_GetWindowTextLength(Hwnd)
  If Length = 0 Then
    Caption = ""
  Else
    Caption = Space$(Length + 1)
    Length = api_GetWindowText(Hwnd, Caption, Length + 1)
    Caption = Left$(Caption, Len(Caption) - 1)
  End If
  
  m_Win_GetTitle = Caption

End Function

Public Function m_Win_GetNextActive(Nbr As Integer, Optional AvoidHwnd As Long) As Long
'Returns the handle of the next visible window after the active window (using Windows Z order).
'The function can search for the next window, or more depending to the [Nbr] value.
'The function can ignore one specific Window given with the [AvoidHwnd] parameter
'  (this is usefull for searching the windows that is not the application itself)
  
  'Initilize variable for the search
  WinNbr = Nbr
  WinXNum = 0
  WinXHwnd = 0
  WinZHwnd = AvoidHwnd
  
  'm_Log "-----------  GetNextActive(Nbr=" & Nbr & ",AvoidHwnd=" & AvoidHwnd & ") -----------"
  
  'Scann all windows. The EnumWindows API calls the specified function as many times as there are visible windows
  api_EnumWindows AddressOf p_EnumWindowsProc, 0
  
  'Returns the result
  m_Win_GetNextActive = WinXHwnd 'Next visible window
  
End Function

Private Function p_EnumWindowsProc(ByVal Hwnd As Long, ByVal Param As Long) As Long
'This function is called by EnumWindows API function as many times as they are visible windows
  
  Dim Ok As Boolean
  
  'This code check if the currently scanned window is visible or not.
  'There is no further check is the searched window is found.
  'This is beacause the function call continue even if the searced wind is founded.
  Ok = False
  If (WinXHwnd = 0) And (WinZHwnd <> Hwnd) Then
    Ok = p_Win_Visible(Hwnd)
  End If
  
  'Count how many visible window have bee found
  If Ok Then
    If WinXNum >= WinNbr Then
      WinXHwnd = Hwnd
    Else
      WinXNum = WinXNum + 1
    End If
  End If
  
'  'developper stuff (to be deleted)
'  Dim X As String
'  Ok = p_Win_Visible(Hwnd)
'  X = ""
'  X = X & " " & IIf(WinZHwnd = Hwnd, "xxxxx", IIf(Ok, "+++++", "....."))
'  X = X & " " & Hwnd & "=" & m_Win_GetTitle(Hwnd)
'  X = X & IIf(api_IsWindowVisible(Hwnd), " (vis)", " (inv)")
'' X = X & IIf(m_CheckFlag(WinXInfo.dwStyle, WS_VISIBLE), " (vis2)", " (inv2)")
'  X = X & IIf(m_CheckFlag(WinXInfo.dwStyle, WS_MAXIMIZE), " (max)", "")
'  X = X & IIf(m_CheckFlag(WinXInfo.dwStyle, WS_MINIMIZE), " (min)", "")
'  X = X & IIf(m_CheckFlag(WinXInfo.dwExStyle, WS_EX_TOPMOST), " (topmost)", "")
'  X = X & " (top=" & WinXInfo.rcWindow.Top & ",bot=" & WinXInfo.rcWindow.Bottom & ",left=" & WinXInfo.rcWindow.Left & ",rigth=" & WinXInfo.rcWindow.Right & ") "
'  m_Log X
  
  'Returns allways 1
  p_EnumWindowsProc = 1
  
End Function

Private Function p_Win_Visible(Hwnd As Long) As Boolean

  Dim Visible As Boolean

  Visible = api_IsWindowVisible(Hwnd) 'this is equivalant as checking the WS_VISIBLE flag of WINDOWINFO
  If Visible Then
    api_GetWindowInfo Hwnd, WinXInfo
    With WinXInfo
      Visible = Visible And (Not m_CheckFlag(.dwStyle, WS_MINIMIZE))
      Visible = Visible And (.rcWindow.Bottom > .rcWindow.Top)
      Visible = Visible And (.rcWindow.Right > .rcWindow.Left)
      If Visible Then
        If m_CheckFlag(.dwExStyle, WS_EX_TOPMOST) Then 'TopMost is a window that stay on top, they appears first in the Z Order but they must not be coutned
          Visible = False
        End If
      End If
    End With
  End If

  p_Win_Visible = Visible

End Function

Public Sub m_Win_GetPosition(ByRef Hwnd As Long, ByRef XInPixels As Long, ByRef YInPixels As Long)
'Get the position of the window specified with his handle
  
  Dim WinInfo As WINDOWINFO
  
  api_GetWindowInfo Hwnd, WinInfo
  
  XInPixels = WinInfo.rcWindow.Left
  YInPixels = WinInfo.rcWindow.Top

'  Dim WinP As WINDOWPLACEMENT
'
'  api_GetWindowPlacement Hwnd, WinP
'
'  Select Case WinP.showCmd
'  Case SW_SHOWMAXIMIZED, SW_MAXIMIZE
'    XInPixels = WinP.ptMaxPosition.X
'    YInPixels = WinP.ptMaxPosition.Y
'  Case SW_SHOWMINIMIZED, SW_MINIMIZE
'    XInPixels = WinP.ptMinPosition.X
'    YInPixels = WinP.ptMinPosition.Y
'  Case SW_SHOWNORMAL, SW_RESTORE
'    XInPixels = WinP.rcNormalPosition.Left
'    YInPixels = WinP.rcNormalPosition.Top
'  Case Else
'    XInPixels = 0
'    YInPixels = 0
'  End Select
  
End Sub

Public Sub m_Win_GetSize(ByRef Hwnd As Long, ByRef WInPixels As Long, ByRef HInPixels As Long)
'Get the size of the window specified with his handle
  
  Dim WinInfo As WINDOWINFO
  
  api_GetWindowInfo Hwnd, WinInfo
  
  With WinInfo.rcWindow
    WInPixels = .Right - .Left
    HInPixels = .Bottom - .Top
  End With
  
End Sub

Public Sub m_Win_StayOnTop(Hwnd As Long, Stay As Boolean)

  If Stay Then
    api_SetWindowPos Hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  Else
    api_SetWindowPos Hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
  End If

End Sub

Public Sub m_Win_ShowActiveWinList(Nbr As Integer)

  Dim i As Integer
  Dim x As String
  Dim h As Long
  
  x = "Only non-TopMost windows and non-Minimized windows are listed." & vbNewLine
  
  For i = 0 To Nbr
    h = m_Win_GetNextActive(i, g_ForbiddenWin)
    x = x & vbNewLine & i & " : " & m_Win_GetTitle(h)
  Next i

  MsgBox x, vbInformation, "Active windows list"

End Sub

Public Sub m_Log(Line As String)

  Dim FicNum As Integer
  
  FicNum = FreeFile()
  Open App.Path & "\" & App.EXEName & ".log" For Append Access Write As #FicNum
  Print #FicNum, Line
  Close #FicNum

End Sub

Public Function m_Pixel_GetColor(TwipsX As Long, TwipsY As Long) As Long
'Get the actual color of the pixel on the screen
  
  Dim ScreenDC As Long
  
  ScreenDC = api_CreateDC("DISPLAY", "", "", 0&)
  m_Pixel_GetColor = api_GetPixel(ScreenDC, TwipsX, TwipsY)
  api_DeleteDC ScreenDC
  
End Function

Public Function m_ItemLst_Get(ItemLst As String, ItemName As String) As String
'Get the value of an item into a list that has the fallowing form :
'Name1=Value1;Name2=Value2;Name3=Value3;...
  
  Dim Found As Boolean 'True if the item has been found in the list
  Dim Pos   As Integer
  Dim Lst1  As String
  Dim Lst2  As String
  Dim x     As String

  'Search for the item name
  Found = False
  Pos = 0
  Do
    Pos = InStr(Pos + 1, ItemLst, ItemName)
    If Pos > 0 Then
      'Check if it begins ok
      Lst1 = Left$(ItemLst, Pos - 1)
      Lst1 = Trim$(Lst1)
      x = Right$(Lst1, Len(c_Sep))
      If (x = "") Or (x = c_Sep) Then
        'Check if it ends ok
        Lst2 = Mid$(ItemLst, Pos + Len(ItemName))
        Lst2 = Trim$(Lst2)
        x = Left$(Lst2, Len(c_Sym))
        If (x = "") Or (x = c_Sym) Then
          Found = True
        End If
      End If
    End If
  Loop Until Found Or (Pos = 0)
  
  If Found Then
    Lst2 = Mid$(Lst2, Len(c_Sym) + 1)
    Lst2 = Trim$(Lst2)
    Pos = InStr(Lst2, c_Sep)
    If Pos > 0 Then
      x = Left$(Lst2, Pos - 1)
      x = Trim$(x)
    Else
      x = Lst2
    End If
  Else
    x = ""
  End If
  
  m_ItemLst_Get = x
  
End Function

Public Function m_ItemLst_Set(ItemLst As String, ItemName As String, ItemValue As String) As String
'Set the value of an item into a list that has the fallowing form :
'Name1=Value1;Name2=Value2;Name3=Value3;...

  Dim Found As Boolean 'True if the item has been found in the list
  Dim Pos   As Integer
  Dim Lst1  As String
  Dim Lst2  As String
  Dim x     As String

  'Search for the item name
  Found = False
  Pos = 0
  Do
    Pos = InStr(Pos + 1, ItemLst, ItemName)
    If Pos > 0 Then
      'Check if it begins ok
      Lst1 = Left$(ItemLst, Pos - 1)
      Lst1 = Trim$(Lst1)
      x = Right$(Lst1, Len(c_Sep))
      If (x = "") Or (x = c_Sep) Then
        'Check if it ends ok
        Lst2 = Mid$(ItemLst, Pos + Len(ItemName))
        Lst2 = Trim$(Lst2)
        x = Left$(Lst2, Len(c_Sym))
        If (x = "") Or (x = c_Sym) Then
          Found = True
        End If
      End If
    End If
  Loop Until Found Or (Pos = 0)
  
  If Found Then
    Lst2 = Mid$(Lst2, Len(c_Sym) + 1)
    Lst2 = Trim$(Lst2)
    Pos = InStr(Lst2, c_Sep)
    If Pos > 0 Then
      Lst2 = Mid$(Lst2, Pos)
      Lst2 = Trim$(Lst2)
    Else
      Lst2 = ""
    End If
  Else
    Lst1 = Trim$(ItemLst)
    If Lst1 <> "" Then
      Lst1 = Lst1 & c_Sep
    End If
    Lst2 = ""
  End If
  
  ItemName = Replace(ItemName, c_Sep, "")
  'ItemValue = Replace(ItemValue, c_Sym, "")
  m_ItemLst_Set = Lst1 & ItemName & c_Sym & Replace(ItemValue, c_Sep, vbNullString) & Lst2

End Function

Public Sub m_Wait(NbrMilliSec As Long, Optional NoEvents As Boolean)

  Dim Ticks0 As Long
  
  Ticks0 = api_GetTickCount()
  Do Until api_GetTickCount() >= Ticks0 + NbrMilliSec
    If Not NoEvents Then
      DoEvents
    End If
  Loop

End Sub

Public Function m_File_Open(aTitre As String, Optional aFiltre As String, Optional aDoitExister As Boolean, Optional aFichier As String, Optional aRépertoire As String, Optional Hwnd As Long) As String

    Const OFN_AUTORISERMULTISELECT = &H200
    Const OFN_CREERATTENTE = &H2000
    Const OFN_EXPLORER = &H80000                ' Nouveau design commdlg.
    Const OFN_FICHIERDOITEXISTER = &H1000
    Const OFN_MASQUERLECTURESEULE = &H4
    Const OFN_PASCHANGERREP = &H8
    Const OFN_LIAISONSREFERENCES = &H100000
    Const OFN_PASBOUTONRESEAU = &H20000
    Const OFN_PASRETOURLECTURESEULE = &H8000
    Const OFN_PASVALIDER = &H100
    Const OFN_ECRASERINVITE = &H2
    Const OFN_CHEMINDOITEXISTER = &H800
    Const OFN_LECTURESEULE = &H1
    Const OFN_AFFICHERAIDE = &H10

    Dim vOF    As t_OuvrirFichier
    Dim Pos    As String
    Dim retour As Boolean
    
    With vOF
    
        .hwndPropriétaire = Hwnd
        
        .hInstance = 0
        .lpstrFiltrePersonnalisé = 0
        .nFiltrePersonMax = 0
        .lpfnCrochet = 0
        .lpNomModèle = 0
        .lDonnéesClient = 0
        
        'Confection de la chaîne filtre
        If aFiltre = "" Then
            aFiltre = "Tous (*.*)|*.*||"
        End If
        Pos = InStr(aFiltre, "|")
        Do Until Pos = 0
            aFiltre = Left$(aFiltre, Pos - 1) & vbNullChar & Mid$(aFiltre, Pos + 1)
            Pos = InStr(aFiltre, "|")
        Loop
        .lpstrFiltre = aFiltre
        
        .nIndexFiltre = 1
        .lpstrFichier = aFichier & String(512 - Len(aFichier), 0)
        .nFichierMax = 511
        .lpstrTitreFichier = String(512, 0)
        .nTitreFichierMax = 511
        .lpstrTitre = aTitre
        .lpstrRépInitial = aRépertoire
        .lpstrExtDéf = ""
        .indicateurs = OFN_MASQUERLECTURESEULE
        .lTailleStruct = Len(vOF)
        
        
        If aDoitExister Then
            retour = api_GetOpenFileName(vOF)
        Else
            retour = api_GetSaveFileName(vOF)
        End If
        
        If retour Then
        'On retire le chr(0) de fin de chaîne
            m_File_Open = Left$(.lpstrFichier, InStr(.lpstrFichier, vbNullChar) - 1)
        Else
            m_File_Open = ""
        End If
        
    End With

End Function

Public Function m_Txt_IncrementeCopie(Txt As String) As String
'Incremente the last integer value that is present at the end of the texte
  
  Dim Pos As Long
  Dim Ok  As Boolean
  Dim x   As String
  
  'Star from and end and search the last char that is a number
  Pos = Len(Txt)
  Ok = True
  Do Until (Pos = 0) Or (Ok = False)
    x = Mid$(RTrim$(Txt), Pos, 1)
    If (Asc(x) >= Asc("0")) And (Asc(x) <= Asc("9")) Then
      Pos = Pos - 1
    Else
      Ok = False
    End If
  Loop

  x = Mid$(Txt, Pos + 1)
  If x = "" Then
    x = " 2"
  Else
    x = "" & (Val(x) + 1)
  End If
  
  m_Txt_IncrementeCopie = Left$(Txt, Pos) & x
  
End Function

Public Function m_File_GetName(FileFullPath As String) As String
'Returns the folder path of the specified file.

  Dim i As Integer
  
  m_File_GetName = FileFullPath
  i = Len(FileFullPath)
  Do Until i = 0
    If Mid$(FileFullPath, i, 1) = "\" Then
      m_File_GetName = Mid$(FileFullPath, i + 1)
      i = 1
    Else
      i = i - 1
    End If
  Loop

End Function

Public Function m_File_Exists(File As String) As Boolean

  On Error Resume Next
  m_File_Exists = False
  m_File_Exists = (Dir$(File) <> vbNullString)

End Function

Public Function m_Frm_PositionCtr(Frm As Form, Optional Init As Boolean)
'Cette fonction permet de repositionner automatiquement les Ctr d'un formulaire.
'Ajouter "m_Frm_PositionCtr Me, True"  sur l'évènement Load   du formulaire.
'Ajouter "m_Frm_PositionCtr Me, False" sur l'évènement Resize du formulaire.
'Dans vos contrôles, ajoutez à la propriété Remarque (Tag en VBA)
' les mots clés "MargeG=","MargeD=","MargeH=","MargeB=" sans valeur affectée
' et séparés par des points-virgules. Chaque mot-clé indique une marge qui reste fixe.
' Pour que le contrôle s'étire sans se déplacer, utilisez 2 mots-clés.
' Exemple : Remarque = "MargeG=;MargeD="

  Const c_Initialiser = -2
  Const c_PasDeMarge = -1
  Const c_IdxMargeG = 1
  Const c_IdxMargeD = 2
  Const c_IdxMargeH = 3
  Const c_IdxMargeB = 4
  Const c_TxtMargeG = "MargeG="
  Const c_TxtMargeD = "MargeD="
  Const c_TxtMargeH = "MargeH="
  Const c_TxtMargeB = "MargeB="

  Dim Ctr  As Control
  Dim j    As Long
  Dim x    As String
  Dim z    As Long
  Dim PosM As Long
  Dim PosS As Long
  Dim NbrMrg As Integer 'Nombre de déf marge trouvée dans le contrôle
  
  Dim MargeVal(1 To 4) As Long
  Dim MargeTxt(1 To 4) As String

  'Initialisation du tableau
  MargeTxt(c_IdxMargeG) = c_TxtMargeG
  MargeTxt(c_IdxMargeD) = c_TxtMargeD
  MargeTxt(c_IdxMargeH) = c_TxtMargeH
  MargeTxt(c_IdxMargeB) = c_TxtMargeB

  For Each Ctr In Frm
  
    If Trim$(Ctr.Tag) <> vbNullString Then
    
      'On lit les infos de marges
      NbrMrg = 0
      For j = 1 To 4
        PosM = InStr(Ctr.Tag, MargeTxt(j))
        If PosM > 0 Then
          NbrMrg = NbrMrg + 1
          PosM = PosM + Len(MargeTxt(j))
          PosS = InStr(PosM, Ctr.Tag, ";") 'recherche du séparateur
          If PosS = 0 Then
            x = Mid$(Ctr.Tag, PosM)
          Else
            x = Mid$(Ctr.Tag, PosM, PosS - PosM)
          End If
          If Trim$(x) = vbNullString Then
            MargeVal(j) = c_Initialiser
          Else
            MargeVal(j) = Val(x)
          End If
        Else
          MargeVal(j) = c_PasDeMarge
        End If
      Next j
      
      'On a récolté les infos des marges, mainteant on les initalise ou on les applique
      If NbrMrg > 0 Then
        
        If Init Then
        
          'Initialisation des marges
          '-------------------------
          
          For j = 1 To 4
            If MargeVal(j) = c_Initialiser Then
              'Initialisation des déf de marge du CTR
              Select Case j
              Case c_IdxMargeG: MargeVal(j) = Ctr.Left
              Case c_IdxMargeD: MargeVal(j) = Frm.Width - Ctr.Left - Ctr.Width
              Case c_IdxMargeH: MargeVal(j) = Ctr.Top
              Case c_IdxMargeB: MargeVal(j) = Frm.Height - Ctr.Top - Ctr.Height
              End Select
              Ctr.Tag = Replace(Ctr.Tag, MargeTxt(j), MargeTxt(j) & MargeVal(j))
            End If
          Next j
          
        Else
          
          'Déplacement du Ctr
          '------------------
          
          'Déplacement horizontal
          If MargeVal(c_IdxMargeG) >= 0 Then
            If MargeVal(c_IdxMargeD) >= 0 Then
              'Étirement du contrôle
              z = Frm.Width - MargeVal(c_IdxMargeG) - MargeVal(c_IdxMargeD)
              If z > 0 Then Ctr.Width = z
            Else
              'Déplacement du contrôle
              z = MargeVal(c_IdxMargeG)
              If z > 0 Then Ctr.Left = z
            End If
          Else
            If MargeVal(c_IdxMargeD) >= 0 Then
              'Déplacement du contrôle
              z = Frm.Width - Ctr.Width - MargeVal(c_IdxMargeD)
              If z > 0 Then Ctr.Left = z
            End If
          End If
        
          'Déplacement vertical
          If MargeVal(c_IdxMargeH) >= 0 Then
            If MargeVal(c_IdxMargeB) >= 0 Then
              'Étirement du contrôle
              z = Frm.Height - MargeVal(c_IdxMargeH) - MargeVal(c_IdxMargeB)
              If z > 0 Then Ctr.Height = z
            Else
              'Déplacement du contrôle
              z = MargeVal(c_IdxMargeH)
              If z > 0 Then Ctr.Top = z
            End If
          Else
            If MargeVal(c_IdxMargeB) >= 0 Then
              'Déplacement du contrôle
              z = Frm.Height - Ctr.Height - MargeVal(c_IdxMargeB)
              If z > 0 Then Ctr.Top = z
            End If
          End If
        
        End If
      End If
      
    End If
    
  Next Ctr

End Function
