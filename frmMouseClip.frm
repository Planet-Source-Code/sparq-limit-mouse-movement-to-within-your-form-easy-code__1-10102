VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "ClipMouseToWindow"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1155
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Any) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long

' confine (clip) the mouse's cursor to the client
' area of a given window or control.
' if the argument is omitted, any current clipping is canceled
Sub ClipMouseToWindow(Optional ByVal hWnd As Long)
    Dim rcTarg As RECT
                         
        If hWnd Then
            ' clip the mouse to the specified window
            ' get the window's client area
            GetClientRect hWnd, rcTarg
            ' convert to screen coordinates. Two steps:
            ' first, the upper-left corner
            ClientToScreen hWnd, rcTarg
            ' next, the bottom-right corner
            ClientToScreen hWnd, rcTarg.Right
            ' finally, we can clip the cursor
            ClipCursor rcTarg
        Else
            ' unclip the mouse if no argument has been passed
            ClipCursor ByVal 0&
        End If
End Sub

Private Sub Form_Load()
    ClipMouseToWindow (Me.hWnd)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ClipMouseToWindow
End Sub
