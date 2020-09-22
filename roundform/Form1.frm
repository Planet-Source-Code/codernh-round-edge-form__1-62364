VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Round Edge Form"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Dim ret As Long
Private Sub Form_Paint()
    RoundBorder Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Public Sub RoundBorder(ByVal Frm As Form)
    Dim PriorScaleMode As Long
    PriorScaleMode = Frm.ScaleMode
    If PriorScaleMode <> vbPixels Then Frm.ScaleMode = vbPixels
    ret = CreateRoundRectRgn((Frm.ScaleWidth / 60), (Frm.ScaleHeight / 60), Frm.ScaleWidth - 4, Frm.ScaleHeight, (Frm.ScaleWidth / 20), (Frm.ScaleHeight / 20))
    SelectClipRgn Frm.hdc, ret
    GetClipRgn Frm.hdc, ret
    SetWindowRgn Frm.hWnd, ret, True
    If PriorScaleMode <> vbPixels Then Frm.ScaleMode = PriorScaleMode
End Sub


