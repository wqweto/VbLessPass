Attribute VB_Name = "mdPassword"
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

#Const HasPtrSafe = (VBA7 <> 0)

'=========================================================================
' API
'=========================================================================

#If HasPtrSafe Then
'--- bcrypt
Private Declare PtrSafe Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal hPrf As LongPtr, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As currency, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#Else
Private Enum LongPtr
    [_]
End Enum
'--- bcrypt
Private Declare Function BCryptOpenAlgorithmProvider Lib "bcrypt" (phAlgorithm As LongPtr, ByVal pszAlgId As LongPtr, ByVal pszImplementation As LongPtr, ByVal dwFlags As Long) As Long
Private Declare Function BCryptCloseAlgorithmProvider Lib "bcrypt" (ByVal hAlgorithm As LongPtr, ByVal dwFlags As Long) As Long
Private Declare Function BCryptDeriveKeyPBKDF2 Lib "bcrypt" (ByVal hPrf As LongPtr, pbPassword As Any, ByVal cbPassword As Long, pbSalt As Any, ByVal cbSalt As Long, ByVal cIterations As Currency, pbDerivedKey As Any, ByVal cbDerivedKey As Long, ByVal dwFlags As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As LongPtr, ByVal cchWideChar As Long, lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As LongPtr, ByVal lpUsedDefaultChar As LongPtr) As Long
#End If

Public Type LessPasswordProfile
    Lowercase           As Boolean
    Uppercase           As Boolean
    Digits              As Boolean
    Symbols             As Boolean
    Length              As Long
    Counter             As Long
    Exclude             As String
    Hash                As String
    Iterations          As Long
    KeySize             As Long
End Type

'=========================================================================
' Functions
'=========================================================================

Public Function DefPasswordProfile() As LessPasswordProfile
    With DefPasswordProfile
        .Lowercase = True
        .Uppercase = True
        .Digits = True
        .Symbols = True
        .Length = 16
        .Counter = 1
        .Hash = "SHA256"
        .Iterations = 100000
        .KeySize = 32
    End With
End Function

Public Function GeneratePassword(sSite As String, sLogin As String, sMasterPassword As String, uProfile As LessPasswordProfile) As String
    GeneratePassword = pvRenderPassword(uProfile, pvCalcEntropy(sSite, sLogin, sMasterPassword, uProfile))
End Function

'= private ===============================================================

Private Function pvCalcEntropy(sSite As String, sLogin As String, sMasterPassword As String, uProfile As LessPasswordProfile) As Byte()
    Const BCRYPT_ALG_HANDLE_HMAC_FLAG   As Long = 8
    Dim hShaAlg         As Long
    Dim baPass()        As Byte
    Dim baSalt()        As Byte
    Dim baRetVal()      As Byte
    Dim hResult         As Long
    Dim sApiName        As String
    
    hResult = BCryptOpenAlgorithmProvider(hShaAlg, StrPtr(uProfile.Hash), StrPtr("Microsoft Primitive Provider"), BCRYPT_ALG_HANDLE_HMAC_FLAG)
    If hResult < 0 Then
        sApiName = "BCryptOpenAlgorithmProvider"
        GoTo QH
    End If
    baPass = ToUtf8Array(sMasterPassword)
    baSalt = ToUtf8Array(sSite & sLogin & LCase$(Hex$(uProfile.Counter)))
    ReDim baRetVal(0 To uProfile.KeySize - 1) As Byte
    hResult = BCryptDeriveKeyPBKDF2(hShaAlg, baPass(0), UBound(baPass) + 1, baSalt(0), UBound(baSalt) + 1, uProfile.Iterations / 10000@, baRetVal(0), UBound(baRetVal) + 1, 0)
    If hResult < 0 Then
        sApiName = "BCryptDeriveKeyPBKDF2"
        GoTo QH
    End If
    pvCalcEntropy = baRetVal
QH:
    If hShaAlg <> 0 Then
        Call BCryptCloseAlgorithmProvider(hShaAlg, 0)
    End If
    If LenB(sApiName) <> 0 Then
        Err.Raise hResult, sApiName, "&H" & Hex$(hResult)
    End If
End Function

Private Function pvRenderPassword(uProfile As LessPasswordProfile, baEntropy() As Byte) As String
    Const STR_LOWERCASE As String = "abcdefghijklmnopqrstuvwxyz"
    Const STR_UPPERCASE As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Const STR_DIGITS    As String = "0123456789"
    Const STR_SYMBOLS   As String = "!""#$%&'()*+,-./:;<=>?@[\]^_`{|}~"
    Dim vRules          As Variant
    Dim lIdx            As Long
    Dim vElem           As Variant
    Dim sSetOfChars     As String
    Dim sCharsToAdd     As String
    Dim lRemainder      As Long
    
    '--- get configured rules
    If uProfile.Lowercase Then
        sSetOfChars = sSetOfChars & Chr$(1) & STR_LOWERCASE
    End If
    If uProfile.Uppercase Then
        sSetOfChars = sSetOfChars & Chr$(1) & STR_UPPERCASE
    End If
    If uProfile.Digits Then
        sSetOfChars = sSetOfChars & Chr$(1) & STR_DIGITS
    End If
    If uProfile.Symbols Then
        sSetOfChars = sSetOfChars & Chr$(1) & STR_SYMBOLS
    End If
    vRules = Split(Mid$(sSetOfChars, 2), Chr$(1))
    sSetOfChars = Replace(sSetOfChars, Chr$(1), vbNullString)
    '--- generate initial password
    sSetOfChars = pvRemoveExcludedChars(sSetOfChars, uProfile.Exclude)
    pvRenderPassword = pvConsumeEntropy(baEntropy, sSetOfChars, uProfile.Length - UBound(vRules) - 1)
    '--- get one character per rule
    For Each vElem In vRules
        sSetOfChars = pvRemoveExcludedChars(CStr(vElem), uProfile.Exclude)
        sCharsToAdd = sCharsToAdd & pvConsumeEntropy(baEntropy, sSetOfChars, 1)
    Next
    '--- insert strings pseudo-randomly
    For lIdx = 1 To Len(sCharsToAdd)
        lRemainder = pvDivMod(baEntropy, Len(pvRenderPassword))
        pvRenderPassword = Left$(pvRenderPassword, lRemainder) & Mid$(sCharsToAdd, lIdx, 1) & Mid$(pvRenderPassword, lRemainder + 1)
    Next
End Function

Private Function pvRemoveExcludedChars(sChars As String, sExclude As String) As String
    Dim lIdx            As Long
    
    pvRemoveExcludedChars = sChars
    For lIdx = 1 To Len(sExclude)
        pvRemoveExcludedChars = Replace(pvRemoveExcludedChars, Mid$(sExclude, lIdx, 1), vbNullString)
    Next
End Function

Private Function pvConsumeEntropy(baEntropy() As Byte, sChars As String, ByVal lSize As Long) As String
    Dim lRemainder       As Long
    
    Do While Len(pvConsumeEntropy) < lSize
        lRemainder = pvDivMod(baEntropy, Len(sChars))
        pvConsumeEntropy = pvConsumeEntropy & Mid$(sChars, lRemainder + 1, 1)
    Loop
End Function

Private Function pvDivMod(baDivident() As Byte, ByVal lDivisor As Long) As Long
    Dim lIdx            As Long
    
    For lIdx = 0 To UBound(baDivident)
        pvDivMod = pvDivMod * 256 + baDivident(lIdx)
        baDivident(lIdx) = pvDivMod \ lDivisor
        pvDivMod = pvDivMod Mod lDivisor
    Next
End Function

'= shared ================================================================

Private Function ToUtf8Array(sText As String) As Byte()
    Const CP_UTF8       As Long = 65001
    Dim baRetVal()      As Byte
    Dim lSize           As Long
    
    lSize = WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), ByVal 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baRetVal(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(CP_UTF8, 0, StrPtr(sText), Len(sText), baRetVal(0), lSize, 0, 0)
    Else
        baRetVal = vbNullString
    End If
    ToUtf8Array = baRetVal
End Function
