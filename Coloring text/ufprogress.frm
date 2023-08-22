VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufprogress 
   Caption         =   "Running..."
   ClientHeight    =   1230
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   OleObjectBlob   =   "ufprogress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufprogress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LabelCopyright_Click()

End Sub

'PLACE IN YOUR USERFORM CODE
Private Sub UserForm_Initialize()
#If IsMac = False Then
    'hide the title bar if you're working on a windows machine. Otherwise, just display it as you normally would
    Me.Height = Me.Height - 10
    HideTitleBar.HideTitleBar Me
#End If
End Sub
