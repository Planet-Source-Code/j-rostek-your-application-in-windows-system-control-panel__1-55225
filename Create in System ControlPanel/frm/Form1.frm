VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows-Standard
   Begin VB.CommandButton Command2 
      Caption         =   "Remove from System Control Panel"
      Height          =   795
      Left            =   120
      TabIndex        =   1
      Top             =   1500
      Width           =   4395
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Create in System Control Panel "
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   4395
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
CreateEntryToSystemPanel "{9d6D8ED6-116D-4D4E-B1C2-87098DB509BA}", "My first Program in Control Panel", "Cool. My Application in Control Panel", App.Path & "\" & App.EXEName & ".exe,0", App.Path & "\" & App.EXEName & ".exe -options"
End Sub

Private Sub Command2_Click()
modControlPanel.DeleteEntryFromSystemPanel "{9d6D8ED6-116D-4D4E-B1C2-87098DB509BA}"
End Sub
