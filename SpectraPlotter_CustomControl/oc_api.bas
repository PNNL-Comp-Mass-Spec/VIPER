'
' This file contains constants used to assist in programming
' with the Olectra Chart controls.


' HugeValue is returned in some API calls when the control
' can't determine an appropriate value.
'
Public Const ocHugeValue As Double = 1E+308


' Windows messages, taken out of the WINAPI.TXT file
'
' Constants for dealing with mouse events
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MOUSELAST = &H209

' Flags set when one of the mouse events is triggered
Public Const MK_LBUTTON = &H1
Public Const MK_MBUTTON = &H10
Public Const MK_RBUTTON = &H2

' Keyboard events for when a key is pressed/released
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYUP = &H101

' Flags set when a keyboard event is triggered
Public Const MK_ALT = &H20
Public Const MK_CONTROL = &H8
Public Const MK_SHIFT = &H4

' The Virtual Key codes 
Public Const VK_ESCAPE = &H1B  'The <Esc> key (ASCII Character 27)
Public Const VK_SHIFT = &H10   'The <Shift> key
Public Const VK_CONTROL = &H11 'The <Ctrl> key
