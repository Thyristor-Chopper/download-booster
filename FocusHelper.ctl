VERSION 5.00
Begin VB.UserControl FocusHelper 
   Appearance      =   0  '평면
   BackColor       =   &H80000005&
   BackStyle       =   0  '투명
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   ClipBehavior    =   0  '없음
   ClipControls    =   0   'False
   HasDC           =   0   'False
   PaletteMode     =   4  '없음
   ScaleHeight     =   90
   ScaleWidth      =   90
   Windowless      =   -1  'True
End
Attribute VB_Name = "FocusHelper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Event Click()

Private Sub UserControl_EnterFocus()
    RaiseEvent Click
End Sub
