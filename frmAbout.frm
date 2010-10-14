VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00C0C0C0&
   ClientHeight    =   5250
   ClientLeft      =   2355
   ClientTop       =   1950
   ClientWidth     =   7020
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   7020
   Begin VB.TextBox txtLicenseInfo 
      BackColor       =   &H8000000F&
      Height          =   3015
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "frmAbout.frx":030A
      Top             =   1680
      Width           =   6735
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   5760
      TabIndex        =   3
      Top             =   120
      Width           =   1095
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         ClipControls    =   0   'False
         Height          =   540
         Left            =   280
         Picture         =   "frmAbout.frx":031B
         ScaleHeight     =   337.12
         ScaleMode       =   0  'User
         ScaleWidth      =   337.12
         TabIndex        =   4
         Top             =   320
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   2640
      TabIndex        =   0
      Top             =   4800
      Width           =   1140
   End
   Begin VB.Label lblDate 
      BackStyle       =   0  'Transparent
      Caption         =   "Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   5
      Top             =   560
      Width           =   5325
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   120
      X2              =   6900
      Y1              =   1560
      Y2              =   1560
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Visual Inspection of Peak/Elution Relationships; previously known as Contemporary 2D Displays and LaV2DG"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   570
      Left            =   240
      TabIndex        =   1
      Top             =   900
      Width           =   5205
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "VIPER - Version 1.1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   5205
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   120
      X2              =   6900
      Y1              =   1560
      Y2              =   1560
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created 12/20/98 nt
Option Explicit

' Reg Key Security Options...
'''Const READ_CONTROL = &H20000
'''Const KEY_QUERY_VALUE = &H1
'''Const KEY_SET_VALUE = &H2
'''Const KEY_CREATE_SUB_KEY = &H4
'''Const KEY_ENUMERATE_SUB_KEYS = &H8
'''Const KEY_NOTIFY = &H10
'''Const KEY_CREATE_LINK = &H20
'''Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
'''                       KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
'''                       KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
                     
'' Reg Key ROOT Types...
''Const HKEY_LOCAL_MACHINE = &H80000002
''Const ERROR_SUCCESS = 0
''Const REG_SZ = 1                         ' Unicode nul terminated string
''Const REG_DWORD = 4                      ' 32-bit number
''
''Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
''Const gREGVALSYSINFOLOC = "MSINFO"
''Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
''Const gREGVALSYSINFO = "PATH"
'
''Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
''Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
''Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long

Private Const APP_DESCRIPTION As String = "Visual Inspection of Peak/Elution Relationships; previously known as Contemporary 2D Displays and LaV2DG"

Private mAdvancedMessageDisplayed As Boolean

Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
    PositionControls
    UpdateMessages
End Sub

Private Sub PositionControls()
    On Error GoTo PositionControlsErrorHandler

    Const MINIMUM_HEIGHT As Long = 3750

    Dim lngDesiredValue As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Height < MINIMUM_HEIGHT Then
        Me.Height = MINIMUM_HEIGHT
    End If
    
    lngDesiredValue = Me.Height - txtLicenseInfo.Top - 1200
    If lngDesiredValue < 1000 Then
        lngDesiredValue = 1000
    End If
    
    txtLicenseInfo.Height = lngDesiredValue
    
    lngDesiredValue = Me.width - txtLicenseInfo.Left - 240
    If lngDesiredValue < 1000 Then lngDesiredValue = 1000
    txtLicenseInfo.width = lngDesiredValue
    
    lngDesiredValue = txtLicenseInfo.Left + txtLicenseInfo.width / 2 - cmdOK.width / 2
    If lngDesiredValue < 120 Then lngDesiredValue = 120
    
    cmdOK.Left = lngDesiredValue
    cmdOK.Top = txtLicenseInfo.Top + txtLicenseInfo.Height + 120

    Exit Sub

PositionControlsErrorHandler:
    Debug.Assert False

End Sub

Private Sub UpdateMessages()
    Dim strMessage As String

    Me.Caption = "About " & App.Title
    
    strMessage = App.Title & " - Version " & GetProgramVersion()
    
    lblTitle.Caption = strMessage
    
    lblDate.Caption = APP_BUILD_DATE
    
    lblDescription.Caption = APP_DESCRIPTION
    
    strMessage = ""
    strMessage = strMessage & "Program written by Matthew Monroe, Nikola Tolic, Deep Jaitly, Kyle Littlefield, and Jason McCann for the Department of Energy (PNNL, Richland, WA) in 2000-2006" & vbCrLf & vbCrLf
    
    strMessage = strMessage & "E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com" & vbCrLf
    strMessage = strMessage & "Website: http://omics.pnl.gov/ or http://www.sysbio.org/resources/staff/" & vbCrLf & vbCrLf
    
    If APP_BUILD_DISABLE_LCMSWARP Then
        strMessage = strMessage & "Linear NET Version" & vbCrLf & vbCrLf
    ElseIf APP_BUILD_DISABLE_MTS Then
        strMessage = strMessage & "PNNL Internal Use Version (MTS Disabled)" & vbCrLf & vbCrLf
    Else
        strMessage = strMessage & "PNNL Internal Use Version" & vbCrLf & vbCrLf
    End If
    
    strMessage = strMessage & "For information on some of VIPER's algorithms and on the AMT Tag approach, please see: " & vbCrLf
    
    strMessage = strMessage & """Advances in Proteomics Data Analysis and Display Using an Accurate Mass and Time Tag Approach,"" J.D. Zimmer, M.E. Monroe, W.J. Qian, and R.D. Smith. Mass Spectrometry Reviews, 25, 450-482 (2006)." & vbCrLf & vbCrLf
    
    strMessage = strMessage & "Licensed under the Apache License, Version 2.0; you may not use this file except in compliance with the License.  "
    strMessage = strMessage & "You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0" & vbCrLf & vbCrLf
    
    strMessage = strMessage & "All publications that result from the use of this software should include "
    strMessage = strMessage & "the following acknowledgment statement: " & vbCrLf
    strMessage = strMessage & "Portions of this research were supported by the U.S. Department of Energy "
    strMessage = strMessage & "Office of Biological and Environmental Research Genomes:GtL Program, the NIH "
    strMessage = strMessage & "National Center for Research Resources (Grant RR018522), and the National "
    strMessage = strMessage & "Institute of Allergy and Infectious Diseases (NIH/DHHS through interagency "
    strMessage = strMessage & "agreement Y1-AI-4894-01).  PNNL is operated by Battelle Memorial Institute "
    strMessage = strMessage & "for the U.S. Department of Energy under contract DE-AC05-76RL0 1830." & vbCrLf & vbCrLf
    
    strMessage = strMessage & "Notice: This computer software was prepared by Battelle Memorial Institute, "
    strMessage = strMessage & "hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the "
    strMessage = strMessage & "Department of Energy (DOE).  All rights in the computer software are reserved "
    strMessage = strMessage & "by DOE on behalf of the United States Government and the Contractor as "
    strMessage = strMessage & "provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY "
    strMessage = strMessage & "WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS "
    strMessage = strMessage & "SOFTWARE.  This notice including this sentence must appear on any copies of "
    strMessage = strMessage & "this computer software." & vbCrLf
        
    txtLicenseInfo.Text = strMessage
End Sub
    
Private Sub Form_Resize()
    PositionControls
End Sub

Private Sub lblDescription_Click()
    mAdvancedMessageDisplayed = Not mAdvancedMessageDisplayed
    If mAdvancedMessageDisplayed Then
        If APP_BUILD_DISABLE_LCMSWARP Then
            lblDescription.Caption = "Note: Advanced Warping features are disabled in this release of VIPER"
        End If
    Else
        lblDescription.Caption = APP_DESCRIPTION
    End If
    
End Sub

' Unused Function (May 2003)
'''Public Sub StartSysInfo()
'''    On Error GoTo SysInfoErr
'''
'''    Dim SysInfoPath As String
'''
'''    ' Try To Get System Info Program Path\Name From Registry...
'''    If GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFO, gREGVALSYSINFO, SysInfoPath) Then
'''    ' Try To Get System Info Program Path Only From Registry...
'''    ElseIf GetKeyValue(HKEY_LOCAL_MACHINE, gREGKEYSYSINFOLOC, gREGVALSYSINFOLOC, SysInfoPath) Then
'''        ' Validate Existance Of Known 32 Bit File Version
'''        If (Dir(SysInfoPath & "\MSINFO32.EXE") <> "") Then
'''            SysInfoPath = SysInfoPath & "\MSINFO32.EXE"
'''
'''        ' Error - File Can Not Be Found...
'''        Else
'''            GoTo SysInfoErr
'''        End If
'''    ' Error - Registry Entry Can Not Be Found...
'''    Else
'''        GoTo SysInfoErr
'''    End If
'''
'''    Call Shell(SysInfoPath, vbNormalFocus)
'''
'''    Exit Sub
'''SysInfoErr:
'''    MsgBox "System Information Is Unavailable At This Time", vbOKOnly
'''End Sub
'''
' Unused Function (May 2003)
'''Public Function GetKeyValue(KeyRoot As Long, KeyName As String, SubKeyRef As String, ByRef KeyVal As String) As Boolean
'''    Dim i As Long                                           ' Loop Counter
'''    Dim rc As Long                                          ' Return Code
'''    Dim hKey As Long                                        ' Handle To An Open Registry Key
'''    Dim KeyValType As Long                                  ' Data Type Of A Registry Key
'''    Dim tmpVal As String                                    ' Tempory Storage For A Registry Key Value
'''    Dim KeyValSize As Long                                  ' Size Of Registry Key Variable
'''    '------------------------------------------------------------
'''    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
'''    '------------------------------------------------------------
'''    rc = RegOpenKeyEx(KeyRoot, KeyName, 0, KEY_ALL_ACCESS, hKey) ' Open Registry Key
'''
'''    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Error...
'''
'''    tmpVal = String$(1024, 0)                             ' Allocate Variable Space
'''    KeyValSize = 1024                                       ' Mark Variable Size
'''
'''    '------------------------------------------------------------
'''    ' Retrieve Registry Key Value...
'''    '------------------------------------------------------------
'''    rc = RegQueryValueEx(hKey, SubKeyRef, 0, _
'''                         KeyValType, tmpVal, KeyValSize)    ' Get/Create Key Value
'''
'''    If (rc <> ERROR_SUCCESS) Then GoTo GetKeyError          ' Handle Errors
'''
'''    If (Asc(Mid(tmpVal, KeyValSize, 1)) = 0) Then           ' Win95 Adds Null Terminated String...
'''        tmpVal = Left(tmpVal, KeyValSize - 1)               ' Null Found, Extract From String
'''    Else                                                    ' WinNT Does NOT Null Terminate String...
'''        tmpVal = Left(tmpVal, KeyValSize)                   ' Null Not Found, Extract String Only
'''    End If
'''    '------------------------------------------------------------
'''    ' Determine Key Value Type For Conversion...
'''    '------------------------------------------------------------
'''    Select Case KeyValType                                  ' Search Data Types...
'''    Case REG_SZ                                             ' String Registry Key Data Type
'''        KeyVal = tmpVal                                     ' Copy String Value
'''    Case REG_DWORD                                          ' Double Word Registry Key Data Type
'''        For i = Len(tmpVal) To 1 Step -1                    ' Convert Each Bit
'''            KeyVal = KeyVal + Hex(Asc(Mid(tmpVal, i, 1)))   ' Build Value Char. By Char.
'''        Next
'''        KeyVal = Format$("&h" + KeyVal)                     ' Convert Double Word To String
'''    End Select
'''
'''    GetKeyValue = True                                      ' Return Success
'''    rc = RegCloseKey(hKey)                                  ' Close Registry Key
'''    Exit Function                                           ' Exit
'''
'''GetKeyError:      ' Cleanup After An Error Has Occured...
'''    KeyVal = ""                                             ' Set Return Val To Empty String
'''    GetKeyValue = False                                     ' Return Failure
'''    rc = RegCloseKey(hKey)                                  ' Close Registry Key
'''End Function
Private Sub txtLicenseInfo_Change()

End Sub
