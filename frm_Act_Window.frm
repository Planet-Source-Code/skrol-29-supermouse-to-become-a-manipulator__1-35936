VERSION 5.00
Begin VB.Form frm_Act_Window 
   Caption         =   "Form1"
   ClientHeight    =   3495
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4425
   LinkTopic       =   "Form1"
   ScaleHeight     =   3495
   ScaleWidth      =   4425
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox L_Action 
      Height          =   315
      ItemData        =   "frm_Act_Window.frx":0000
      Left            =   1560
      List            =   "frm_Act_Window.frx":0010
      TabIndex        =   12
      Top             =   2520
      Width           =   2775
   End
   Begin VB.CommandButton btn_OK 
      Caption         =   "Ok"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton btn_Cancel 
      Cancel          =   -1  'True
      Caption         =   "Annuler"
      Height          =   375
      Left            =   3240
      TabIndex        =   4
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CheckBox chk_Enabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Activé"
      Height          =   255
      Left            =   1920
      TabIndex        =   2
      Top             =   120
      Value           =   1  'Checked
      Width           =   975
   End
   Begin VB.TextBox txt_Name 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   2775
   End
   Begin VB.Frame frm_Window 
      Caption         =   "Fenêtre"
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   4215
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1920
         TabIndex        =   11
         Text            =   "Text3"
         Top             =   720
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2520
         TabIndex        =   10
         Text            =   "Text2"
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Text            =   "Text1"
         Top             =   1080
         Width           =   495
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Pixel"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Nommé"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Active"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Action :"
      Height          =   375
      Left            =   120
      TabIndex        =   13
      Top             =   2520
      Width           =   1335
   End
   Begin VB.Label lbl_Name 
      Caption         =   "Désignation"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "frm_Act_Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Label1_Click()

End Sub
