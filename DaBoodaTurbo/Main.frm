VERSION 5.00
Begin VB.Form Main 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DaBoodaTurbo Example"
   ClientHeight    =   9000
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   9000
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "Main.frx":0CCA
   MousePointer    =   99  'Custom
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Me.Show
DoEvents
'AlwaysOnTop Me, True
'I set all my games up under modules, just preference really... or bad habit..
'Run the setup
SetUp

'shut down program
Set Engine = Nothing
End

End Sub

Private Sub Form_Unload(Cancel As Integer)
'unload the engine
Set Engine = Nothing
End

End Sub
