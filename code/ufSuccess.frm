VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufSuccess 
   Caption         =   "Task Completed!"
   ClientHeight    =   9580.001
   ClientLeft      =   -390
   ClientTop       =   -1540
   ClientWidth     =   9680.001
   OleObjectBlob   =   "ufSuccess.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufSuccess"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnOK_Click()
    Unload Me
End Sub
Private Sub btnUpgrade_Click()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink ("https://pythonandvba.com/go/whatsapp-pro-purchase")
    Unload Me
End Sub

Private Sub imgUpgrade_Click()
    On Error Resume Next
    ThisWorkbook.FollowHyperlink ("https://pythonandvba.com/go/whatsapp-pro-purchase")
    Unload Me
End Sub
Private Sub UserForm_Initialize()
    Me.Height = 500
    Me.Width = 430
    'Start Userform Centered inside Excel Screen (for dual monitors)
    Me.StartUpPosition = 0
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
    With Me.imgUpgrade
        .Left = 6
        .Top = 162
        .Width = 408
        .Height = 300
        .PictureSizeMode = fmPictureSizeModeClip
        .PictureAlignment = fmPictureAlignmentCenter
        .AutoSize = False
    End With
End Sub
