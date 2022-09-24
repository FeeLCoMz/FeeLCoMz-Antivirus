Attribute VB_Name = "ModulFadeForm"
Option Explicit

'*************************************************************************
'* Function: FadeForm(WhatForm As Form)
'*
'*
'*************************************************************************
'* Description: This code fade the backgound of a form from black to blue.
'*
'*
'*************************************************************************
'* Parameters: Form
'*
'*************************************************************************
'* Notes:
'*
'*************************************************************************
'* Returns:
'*************************************************************************
Sub FadeForm(WhatForm As Form)
Dim I As Integer
Dim Y As Integer

    WhatForm.AutoRedraw = True
    WhatForm.DrawStyle = 6  ' Inside Solid
    WhatForm.DrawMode = 13  ' Copy Pen (Default)
    WhatForm.DrawWidth = 2
    WhatForm.ScaleMode = 3  ' Pixel (smallest unit of monitor or printer resolution).
    WhatForm.ScaleHeight = (256 * 2)
    For I = 0 To 255
        WhatForm.Line (0, Y)-(WhatForm.Width, Y + 2), RGB(0, 0, I), BF
        Y = Y + 2
    Next I
End Sub
