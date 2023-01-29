VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VB LessPass"
   ClientHeight    =   3828
   ClientLeft      =   36
   ClientTop       =   384
   ClientWidth     =   4884
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   7.8
      Charset         =   204
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3828
   ScaleWidth      =   4884
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCopy 
      Caption         =   "Copy"
      Default         =   -1  'True
      Height          =   432
      Left            =   3948
      TabIndex        =   16
      Top             =   3192
      Width           =   600
   End
   Begin VB.TextBox txtOutput 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   10.2
         Charset         =   204
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   336
      TabIndex        =   15
      Text            =   "output"
      Top             =   3192
      Width           =   3540
   End
   Begin VB.TextBox txtCounter 
      Height          =   288
      Left            =   3948
      TabIndex        =   14
      Text            =   "1"
      Top             =   2436
      Width           =   600
   End
   Begin VB.TextBox txtLength 
      Height          =   288
      Left            =   1596
      TabIndex        =   12
      Text            =   "16"
      Top             =   2436
      Width           =   600
   End
   Begin VB.CheckBox chkSymbols 
      Caption         =   "%!@"
      Height          =   264
      Left            =   3612
      TabIndex        =   10
      Top             =   1848
      Value           =   1  'Checked
      Width           =   684
   End
   Begin VB.CheckBox chkDigits 
      Caption         =   "0-9"
      Height          =   264
      Left            =   2940
      TabIndex        =   9
      Top             =   1848
      Value           =   1  'Checked
      Width           =   684
   End
   Begin VB.CheckBox chkUppercase 
      Caption         =   "A-Z"
      Height          =   264
      Left            =   2268
      TabIndex        =   8
      Top             =   1848
      Value           =   1  'Checked
      Width           =   684
   End
   Begin VB.CheckBox chkLowercase 
      Caption         =   "a-z"
      Height          =   264
      Left            =   1596
      TabIndex        =   7
      Top             =   1848
      Value           =   1  'Checked
      Width           =   684
   End
   Begin VB.TextBox txtMasterPassword 
      Height          =   288
      IMEMode         =   3  'DISABLE
      Left            =   1596
      PasswordChar    =   "*"
      TabIndex        =   5
      Text            =   "password"
      Top             =   1260
      Width           =   2952
   End
   Begin VB.TextBox txtLogin 
      Height          =   288
      Left            =   1596
      TabIndex        =   3
      Text            =   "contact@example.org"
      Top             =   756
      Width           =   2952
   End
   Begin VB.TextBox txtSite 
      Height          =   288
      Left            =   1596
      TabIndex        =   1
      Text            =   "example.org"
      Top             =   252
      Width           =   2952
   End
   Begin VB.Label labLabel 
      Alignment       =   1  'Right Justify
      Caption         =   "Counter:"
      Height          =   264
      Index           =   5
      Left            =   2688
      TabIndex        =   13
      Top             =   2436
      Width           =   1188
   End
   Begin VB.Label labLabel 
      Caption         =   "Length:"
      Height          =   264
      Index           =   4
      Left            =   336
      TabIndex        =   11
      Top             =   2436
      Width           =   1272
   End
   Begin VB.Label labLabel 
      Caption         =   "Options:"
      Height          =   264
      Index           =   3
      Left            =   336
      TabIndex        =   6
      Top             =   1848
      Width           =   1272
   End
   Begin VB.Label labLabel 
      Caption         =   "Master Pass:"
      Height          =   264
      Index           =   2
      Left            =   336
      TabIndex        =   4
      Top             =   1260
      Width           =   1272
   End
   Begin VB.Label labLabel 
      Caption         =   "Login:"
      Height          =   264
      Index           =   1
      Left            =   336
      TabIndex        =   2
      Top             =   756
      Width           =   1272
   End
   Begin VB.Label labLabel 
      Caption         =   "Site:"
      Height          =   264
      Index           =   0
      Left            =   336
      TabIndex        =   0
      Top             =   252
      Width           =   1272
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
'
' VB LessPass Desktop Tool (c) 2023 by wqweto@gmail.com
'
' Based on https://github.com/lesspass/lesspass by Guillaume Vincent
'
' This project is licensed under the terms of the MIT license
' See the LICENSE file in the project root for more information
'
'=========================================================================
Option Explicit
DefObj A-Z

'=========================================================================
' API
'=========================================================================

'--- for SystemParametersInfo
Private Const SPI_GETICONTITLELOGFONT       As Long = 31
Private Const FW_NORMAL                     As Long = 400
Private Const LOGPIXELSY                    As Long = 90

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (lpString As Any) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long

Private Type LOGFONT
    lfHeight            As Long
    lfWidth             As Long
    lfEscapement        As Long
    lfOrientation       As Long
    lfWeight            As Long
    lfItalic            As Byte
    lfUnderline         As Byte
    lfStrikeOut         As Byte
    lfCharSet           As Byte
    lfOutPrecision      As Byte
    lfClipPrecision     As Byte
    lfQuality           As Byte
    lfPitchAndFamily    As Byte
    lfFaceName(1 To 32) As Byte
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const DEF_SITE                  As String = "example.org"
Private Const DEF_LOGIN                 As String = "user@example.org"
Private Const DEF_LENGTH                As Long = 16
Private Const DEF_COUNTER               As Long = 1

'=========================================================================
' Methods
'=========================================================================

Private Sub pvGenerate()
    Dim uProfile        As LessPasswordProfile
    
    On Error GoTo EH
    If LenB(txtSite.Text) = 0 Or LenB(txtLogin.Text) = 0 Or LenB(txtMasterPassword.Text) = 0 Or Znl(Val(txtLength.Text), DEF_LENGTH) <= 4 Then
        txtOutput.Text = vbNullString
        GoTo QH
    End If
    Caption = App.Title
    uProfile = DefPasswordProfile
    uProfile.Lowercase = (chkLowercase.Value = vbChecked)
    uProfile.Uppercase = (chkUppercase.Value = vbChecked)
    uProfile.Digits = (chkDigits.Value = vbChecked)
    uProfile.Symbols = (chkSymbols.Value = vbChecked)
    uProfile.Length = Znl(Val(txtLength.Text), DEF_LENGTH)
    uProfile.Counter = Znl(Val(txtCounter.Text), DEF_COUNTER)
    txtOutput.Text = GeneratePassword(txtSite.Text, txtLogin.Text, txtMasterPassword.Text, uProfile)
QH:
    cmdCopy.Enabled = (LenB(txtOutput.Text) <> 0)
    Exit Sub
EH:
    Debug.Print "Critical error:" & Err.Description
    Caption = "Critical error:" & Err.Description
    txtOutput.Text = vbNullString
    GoTo QH
End Sub

Private Function Znl(vValue As Variant, Optional IfEmptyLong As Variant = Null) As Variant
    Dim lValue As Long
    
    On Error Resume Next
    lValue = CLng(vValue)
    Znl = IIf(lValue = 0, IfEmptyLong, lValue)
End Function

Private Property Get SystemIconFont() As StdFont
    Dim uFont           As LOGFONT
    Dim sBuffer         As String
    Dim hTempDC         As Long
    
    Call SystemParametersInfo(SPI_GETICONTITLELOGFONT, LenB(uFont), uFont, 0)
    Set SystemIconFont = New StdFont
    With SystemIconFont
        sBuffer = Space$(lstrlenA(uFont.lfFaceName(1)))
        CopyMemory ByVal sBuffer, uFont.lfFaceName(1), Len(sBuffer)
        .Name = sBuffer
        .Bold = (uFont.lfWeight >= FW_NORMAL)
        .Charset = uFont.lfCharSet
        .Italic = (uFont.lfItalic <> 0)
        .Strikethrough = (uFont.lfStrikeOut <> 0)
        .Underline = (uFont.lfUnderline <> 0)
        .Weight = uFont.lfWeight
        hTempDC = GetDC(0)
        .Size = -(uFont.lfHeight * 72) / GetDeviceCaps(hTempDC, LOGPIXELSY)
        Call ReleaseDC(0, hTempDC)
    End With
End Property

'=========================================================================
' Control events
'=========================================================================

Private Sub Form_Load()
    Dim oCtl            As Object
    
    On Error GoTo EH
    '--- fix UI
    Set Font = SystemIconFont
    For Each oCtl In Controls
        Set oCtl.Font = Font
    Next
    Set txtOutput.Font = SystemIconFont
    txtOutput.FontBold = True
    txtOutput.FontSize = txtOutput.FontSize + 2
    cmdCopy.Height = txtOutput.Height
    '--- load settings
    txtSite.Text = GetSetting(App.ProductName, "Common", "Site", DEF_SITE)
    txtLogin.Text = GetSetting(App.ProductName, "Common", "Login", DEF_LOGIN)
    txtMasterPassword.Text = GetSetting(App.ProductName, "Common", "MasterPassword", vbNullString)
    chkLowercase.Value = GetSetting(App.ProductName, "Common", "Lowercase", vbChecked)
    chkUppercase.Value = GetSetting(App.ProductName, "Common", "Uppercase", vbChecked)
    chkDigits.Value = GetSetting(App.ProductName, "Common", "Digits", vbChecked)
    chkSymbols.Value = GetSetting(App.ProductName, "Common", "Symbols", vbChecked)
    txtLength.Text = GetSetting(App.ProductName, "Common", "Length", DEF_LENGTH)
    txtCounter.Text = GetSetting(App.ProductName, "Common", "Counter", DEF_COUNTER)
    pvGenerate
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error GoTo EH
    SaveSetting App.ProductName, "Common", "Site", txtSite.Text
    SaveSetting App.ProductName, "Common", "Login", txtLogin.Text
    SaveSetting App.ProductName, "Common", "MasterPassword", txtMasterPassword.Text
    SaveSetting App.ProductName, "Common", "Lowercase", chkLowercase.Value
    SaveSetting App.ProductName, "Common", "Uppercase", chkUppercase.Value
    SaveSetting App.ProductName, "Common", "Digits", chkDigits.Value
    SaveSetting App.ProductName, "Common", "Symbols", chkSymbols.Value
    SaveSetting App.ProductName, "Common", "Length", Znl(Val(txtLength.Text), DEF_LENGTH)
    SaveSetting App.ProductName, "Common", "Counter", Znl(Val(txtCounter.Text), DEF_COUNTER)
    Exit Sub
EH:
    Debug.Print "Critical error: " & Err.Description
End Sub

Private Sub cmdCopy_Click()
    On Error GoTo EH
    Clipboard.Clear
    Clipboard.SetText txtOutput.Text
    Exit Sub
EH:
    MsgBox Err.Description, vbCritical
End Sub

Private Sub txtSite_Change()
    pvGenerate
End Sub

Private Sub txtLogin_Change()
    pvGenerate
End Sub

Private Sub txtMasterPassword_Change()
    pvGenerate
End Sub

Private Sub chkLowercase_Click()
    pvGenerate
End Sub

Private Sub chkUppercase_Click()
    pvGenerate
End Sub

Private Sub chkDigits_Click()
    pvGenerate
End Sub

Private Sub chkSymbols_Click()
    pvGenerate
End Sub

Private Sub txtLength_Change()
    pvGenerate
End Sub

Private Sub txtCounter_Change()
    pvGenerate
End Sub
