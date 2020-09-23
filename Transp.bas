Attribute VB_Name = "Transp"
' Credit to Chris Yates
' I have borrowed this code from Chris Yates. Thanx Chris!
' and i will kindly ask all people to read the authors rules
' for letting you use it.....

' |
' |
' '----> This code was written by Chris Yates. I am freely distributing this code under one condition.
' If you use this code in your program please give me credit and e-mail me (cyates@neo.rr.com) and
' tell me about your program.  Thanks, and enjoy.

Public Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Sub ReleaseCapture Lib "user32" ()
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Const RGN_DIFF = 4
Public Const SC_CLICKMOVE = &HF012&
Public Const WM_SYSCOMMAND = &H112
Dim CurRgn, TempRgn As Long
Public Function AutoFormShape(bg As Form, transColor)
Dim X, Y As Integer
CurRgn = CreateRectRgn(0, 0, bg.ScaleWidth, bg.ScaleHeight)
While Y <= bg.ScaleHeight
    While X <= bg.ScaleWidth
        If GetPixel(bg.hdc, X, Y) = transColor Then
            TempRgn = CreateRectRgn(X, Y, X + 1, Y + 1)
            success = CombineRgn(CurRgn, CurRgn, TempRgn, RGN_DIFF)
            DeleteObject (TempRgn)
        End If
        X = X + 1
    Wend
        Y = Y + 1
        X = 0
Wend
success = SetWindowRgn(bg.hwnd, CurRgn, True)
DeleteObject (CurRgn)
End Function
Public Sub FormDrag(TheForm As Form)
    ReleaseCapture
    Call SendMessage(TheForm.hwnd, &HA1, 2, 0&)
End Sub
