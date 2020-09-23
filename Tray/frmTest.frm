VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "System Tray"
   ClientHeight    =   2190
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   5145
   ClipControls    =   0   'False
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MouseIcon       =   "frmTest.frx":030A
   ScaleHeight     =   2190
   ScaleWidth      =   5145
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add_Tray / Hide_Me"
      Height          =   855
      Left            =   1200
      MouseIcon       =   "frmTest.frx":0614
      TabIndex        =   0
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "The events on tray icon will send to immediate window"
      Height          =   240
      Left            =   180
      TabIndex        =   1
      Top             =   1440
      Width           =   4815
   End
   Begin VB.Menu mnuForm 
      Caption         =   "Form"
      Begin VB.Menu mnuShow 
         Caption         =   "&Show"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "&Hide"
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'[Adding the tray]
Private Sub cmdAdd_Click()
    TrayAdd hwnd, Me.Icon, "System Tray", MouseMove
    mnuHide_Click
End Sub

'[Checking The mouse event]
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim cEvent As Single
cEvent = X / Screen.TwipsPerPixelX
Select Case cEvent
    Case MouseMove
        Debug.Print "MouseMove"
    Case LeftUp
        Debug.Print "Left Up"
    Case LeftDown
        Debug.Print "LeftDown"
    Case LeftDbClick
        Debug.Print "LeftDbClick"
    Case MiddleUp
        Debug.Print "MiddleUp"
    Case MiddleDown
        Debug.Print "MiddleDown"
    Case MiddleDbClick
        Debug.Print "MiddleDbClick"
    Case RightUp
        Debug.Print "RightUp": PopupMenu mnuForm
    Case RightDown
        Debug.Print "RightDown"
    Case RightDbClick
        Debug.Print "RightDbClick"
End Select
End Sub

Private Sub mnuHide_Click()
    If Not Me.WindowState = 1 Then WindowState = 1: Me.Hide
End Sub

Private Sub mnuShow_Click()
    If Me.WindowState = 1 Then WindowState = 0: Me.Show
    TrayDelete  '[Deleting Tray]
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim X As Integer
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "Please tell Me") = vbYes Then
        X = MsgBox("(  Please 'RATE' this code  ).Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "ThankYou")
    Else
        X = MsgBox("( Please give feedback ) to improve this code.Click 'Ok' to copy the site address  to your clipboard", vbInformation + vbOKCancel, "Please Give FeedBack")
    End If
    If X = vbOK Then Clipboard.SetText ("Not set")
End Sub

