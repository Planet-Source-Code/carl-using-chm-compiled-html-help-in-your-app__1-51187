VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "HTML Help Demo"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   HelpContextID   =   10
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdHelpSearch 
      Caption         =   "This button shows the help file search option"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   4095
   End
   Begin VB.CommandButton cmdHelpContextID 
      Caption         =   "This button shows the item at HelpContextID 100"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1560
      Width           =   4095
   End
   Begin VB.CommandButton cmdHelpContents 
      Caption         =   "This button shows the help file contents"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   4095
   End
   Begin VB.Label lbl 
      Caption         =   $"HTMLHelp_Demo.frx":0000
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   0
      Top             =   120
      Width           =   3855
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdHelpContents_Click()
    HTMLHELP_Contents Me
End Sub

Private Sub cmdHelpContextID_Click()
    HTMLHELP_ContextID Me, 100
End Sub

Private Sub cmdHelpSearch_Click()
    HTMLHelp_Search Me
End Sub

Private Sub Form_Load()

    Dim s As String
    
    'Get the application path and verify it ends with "\"
    s = App.Path
    If (Right$(s, 1) <> "\") Then s = s + "\"
    
    'Set the programs help file path
    App.HelpFile = s + "HTMLHelp_Demo.chm"
    
End Sub
