VERSION 5.00
Begin VB.Form frm_Help 
   Caption         =   "Help"
   ClientHeight    =   5265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   6360
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_Help 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2895
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "frm_Help"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

  Dim FileName As String
  Dim FileNum  As Integer
  Dim Line     As String
  
  FileName = Dir$(App.Path & "\" & App.EXEName & ".txt")
  If FileName = "" Then
    MsgBox "The " & App.EXEName & ".txt file is missing.", vbCritical
    Unload Me
  Else
    
    txt_Help = ""
    FileNum = FreeFile()
    Open (App.Path & "\" & FileName) For Input Access Read As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, Line
      txt_Help = txt_Help & Line & vbNewLine
    Loop
    Close FileNum
    
  End If

End Sub

Private Sub Form_Resize()

  txt_Help.Width = Me.Width - 100 '100 is for the border width
  txt_Help.Height = Me.Height - 500

End Sub
