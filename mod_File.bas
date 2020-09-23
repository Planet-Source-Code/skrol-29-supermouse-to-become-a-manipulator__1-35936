Attribute VB_Name = "mod_File"
Option Explicit
Option Compare Text

Public Const c_Ext = ".sms"
Public Const c_Sep = ";"
Public Const c_Sym = "="
Public Const c_Filter = "Supermouse Script (*.sms)|*.sms||"

Public g_Dirty   As Boolean
Public g_File    As String
Public g_RunMode As Integer '0=manual,1=auto opening,2=auto ending

Const c_Opt_Pref = "Opt:"
Const c_Act_Pref = "Act:"

'Action are saved in a line with the fallowing syntax :
'act:type=xxx;name=xxx;...
'Options are saved in a line with the fallowing syntax :
'opt:type=xxx;name=xxx;...

Public Function m_SmsFile_Check() As Boolean
'Check if the current script have been modified.
'If it's so, the application ask for save it.

  Dim Resp As Integer
  Dim Ok   As Boolean
  Dim X    As String

  If g_Dirty Then
    Resp = MsgBox("The current script has been modified, do you want to save the changes ?", vbQuestion + vbDefaultButton3 + vbYesNoCancel)
    Select Case Resp
    Case vbYes
      If g_File = "" Then
        X = m_File_Open("Save As", c_Filter, False, "", App.Path, frm_Main.Hwnd)
        If X = "" Then
          Ok = False
        Else
          m_SmsFile_Save X
          Ok = True
        End If
      End If
    Case vbNo
      Ok = True
    Case vbCancel
      Ok = False
    End Select
  Else
    Ok = True
  End If

  m_SmsFile_Check = Ok

End Function

Public Function m_SmsFile_Save(File As String)
'Save information in the specified file

  Dim FileNum As Integer
  Dim i       As Long
  
  FileNum = FreeFile()
  Open File For Output As #FileNum
  
  'Save options
  For i = 0 To OptionNbr - 1
    Print #FileNum, c_Opt_Pref & OptionLst(i)
  Next i
  
  Print #FileNum, ""
  
  'Save actions
  For i = 0 To ActionNbr - 1
    Print #FileNum, c_Act_Pref & ActionLst(i).Prameters
  Next i
  
  Close #FileNum

End Function

Public Function m_SmsFile_Load(File As String)
'Load information from the specified file

  Dim FileNum As Integer
  Dim i       As Long
  Dim X       As String
  Dim ActType As String
  
  ActionNbr = 0
  
  FileNum = FreeFile()
  Open File For Input Access Read As #FileNum
  
  Do Until EOF(FileNum)
  
    Line Input #FileNum, X
    
    X = Trim$(X)
    If Left$(X, Len(c_Act_Pref)) = c_Act_Pref Then
      'Add the action
      X = Mid$(X, Len(c_Act_Pref) + 1)
      X = Trim$(X)
      m_ActLst_Insert ActionNbr, X
    ElseIf Left$(X, Len(c_Opt_Pref)) = c_Opt_Pref Then
      'Add the option
      X = Mid$(X, Len(c_Opt_Pref) + 1)
      X = Trim$(X)
      OptionNbr = OptionNbr + 1
      ReDim Preserve OptionLst(0 To (OptionNbr - 1))
      OptionLst(OptionNbr - 1) = X
    End If
    
  Loop
  
  Close #FileNum

End Function
