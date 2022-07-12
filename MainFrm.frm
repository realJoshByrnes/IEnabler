VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "IEnabler"
   ClientHeight    =   0
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   0
   ControlBox      =   0   'False
   Icon            =   "MainFrm.frx":0000
   LinkTopic       =   "MainFrm"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   0
   ScaleWidth      =   0
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    On Error GoTo ErrorHandler
    Me.Hide
    Set iExplore = Controls.Add("Shell.Explorer.2", "webctl", Me) ' Create
    iExplore.object.Navigate2 "about:blank" ' Activate
    iExplore.object.Document.parentWindow.execScript "window.open('about:home');" ' Magic Happens (New Window)
    Unload Me ' Self destruct.
    Exit Sub

ErrorHandler:
    MsgBox "Could not create Internet Explorer instance" & vbCrLf & _
    "Have you installed it or enabled Internet Explorer Mode?", vbCritical, "Unable to start Internet Explorer"
    Unload Me
End Sub
