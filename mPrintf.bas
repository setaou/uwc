Attribute VB_Name = "mPrintf"
'//**************************************************************************
'// ----------------- Module -----------------
'// Name        : c++ sprintf/fprintf functions for VB
'// Version     :
'// Author      : Benoit Frigon
'// Created on  : 22-JUL-2002
'// Last update : 24-JUL-2002
'// File        : m_Utils.bas
'// Desc.       :
'//**************************************************************************
'// All rights reserved@Logiciels M.T.L enr. NEQ# 22-48153829(Québec)
'//**************************************************************************
'//
'//==========================================================================
'// Usage :
'// =======
'//
'// sPrintf usage :
'//
'//     %[flags][field width].[precision][conversion character]
'//
'//     ** Flags, field width and precision are optional **
'//
'//     sample #1 : sprintf("number : %10.2f", 23.34899934)
'//
'//                 will output > "Number :      23.35"
'//
'//     sample #2 : sprintf("My name is %s. i am %d years old.", "Benoit", 21)
'//
'//                 will output > "My name is Benoit. i am 21 years old."
'//
'//
'// fprintf usage :
'//
'//     Same as sPrintf except that format string will be printed in a file
'//
'//     sample :  open "c:\output.txt" for output as #20
'//
'//               call fprintf(20, "This is a test -> %d", 34.2)
'//
'//               close #20
'//
'//**************************************************************************
'// Reference :
'// ===========
'//
'// Supported conversion character :
'// ----------------------------------------------------------------
'// %d %i   : Signed decimal notation.
'// %f      : Double-precision floating-point.
'// %g      : Compact double-precision floating-point, using %f or %e.
'// %G      : Same as %g, except it uses %E instead of %e.
'// %e      : Double-precision floating-point, exponential notation using e.
'// %E      : Double-precision floating-point, exponential notation using E.
'// %c      : Single character.
'// %s      : Character string.
'// %o      : Octal notation
'// %x      : Hexadecimal notation, using abcdef
'// %X      : Hexadecimal notation, using ABCDEF
'// %u      : Unsigned decimal notation.
'// %p      : Pointer address
'// %n      : Dump what was printed so far to a variable.
'// %%      : literal % (No argument consumed)
'//
'//
'// Supported flags
'// ----------------------------------------------------------------
'// -       : Left-justify
'// +       : Display a sign (+ or -)
'// 0       : Pad with leading zeroes.
'// *       : Field width specified by an argument.
'// Space   : Pad with spaces.
'//
'//
'// Supported escape characters
'// ----------------------------------------------------------------
'// \r      : carriage return
'// \n      : new line
'// \t      : tab character
'// \\      : backslash character
'//           ** Other escape characters are not supported **
'//
'//**************************************************************************
'// Note :
'// ======
'//
'// Don't expect these functions to be 100% perfect. The original c++
'// sprintf/fprintf function is complex to reproduce.
'//
'// If you find any bugs, e-mail me at mtlsoftware@idz.net
'//
'// Before you e-mail me to ask me if you can use my code in
'// your programs, yes you can, it's free.
'//
'//**************************************************************************
Option Explicit
Option Base 0




'//**************************************************************************
'// Constants
'//**************************************************************************
'//--- Convertion type ---
Private Const VALTYPE_STRING = 1
Private Const VALTYPE_DECIMALSIGNED = 2
Private Const VALTYPE_FLOAT = 3
Private Const VALTYPE_OCTAL = 4
Private Const VALTYPE_HEX = 5
Private Const VALTYPE_CHAR = 6
Private Const VALTYPE_DECIMALUNSIGNED = 7
Private Const VALTYPE_EXPONENTIAL = 8
Private Const VALTYPE_COMPACT_EXPONENTIAL = 9
Private Const VALTYPE_POINTER = 10
Private Const VALTYPE_BUFFERDUMP = 20

'//**************************************************************************
'// sprintf : Formatted print into a string.
'// -------------------
'// Inputs > - lpFormat : Formating string (see instructions above)
'//          - Arguments : Array of arguments
'//
'// Output > Return formated string
'//**************************************************************************
Public Function sprintf(lpFormat As String, ParamArray Arguments() As Variant) As String
    If (UBound(Arguments) <> -1) Then
        
        '//***** Determine if parameter array is passed directly or from another function ****
        Dim isIndirectArray As Boolean
        isIndirectArray = (VarType(Arguments(0)) = (vbArray + vbVariant))
    Else
        isIndirectArray = False
    End If
    
    
    Dim ArgumentIndex As Long
    If (isIndirectArray) Then
        ArgumentIndex = LBound(Arguments(0))
        
        Dim ArgumentsIndirect() As Variant
        Let ArgumentsIndirect = Arguments(0)
    Else
        ArgumentIndex = LBound(Arguments)
    End If
    
    Dim bOnMarker As Boolean
    Dim i As Long
    For i = 1 To Len(lpFormat)
        Dim sChar As String
        sChar = Mid(lpFormat, i, 1)
        
        Dim BeginPos As Long
        If (sChar = "%") And (Not bOnMarker) Then
            
            '//**** Check if this ****
            If (Not Mid(lpFormat, i + 1, 1) = "%") And (i < Len(lpFormat)) Then
                bOnMarker = True
                
                BeginPos = i
                
                i = i + 1
                sChar = Mid(lpFormat, i, 1)
            Else
                i = i + 1
            End If
        End If
        
        Dim bNoArguments As Boolean
        If (bOnMarker) Then
            If (InStr("gGxXeEpnfcdiosu%", sChar)) Then
                bOnMarker = False
                
                Dim sFormat As String
                sFormat = Mid(lpFormat, BeginPos, (i - BeginPos + 1))
                
                Dim ValueType As Long
                Dim bUpper As Boolean
                Dim Width As Long
                Dim Precision As Long
                Dim sPadding As String
                Dim LeftJustify As Boolean
                Dim Signed As Boolean
                Dim bNoUnsignificant As Boolean
                
                '//**** Check if the convertion format is valid ****
                Dim bInvalidFormat As Boolean
                If (IsValidFormat(sFormat, ValueType, bUpper, Width, Precision, sPadding, LeftJustify, Signed, bNoUnsignificant)) Then
                    
                    If (Width = -2) Then    '//**** The width is set by an argument instead of format string ****
                        If (ArgumentIndex <= UBound(Arguments)) Then
                            Dim vArgument As Variant
                            vArgument = Arguments(ArgumentIndex)
                        
                            If (ConvertArgument(vArgument, vbLong)) Then
                                Width = vArgument
                            Else
                                Width = -1
                            End If
                            
                            ArgumentIndex = ArgumentIndex + 1
                            bInvalidFormat = False
                        Else
                            bNoArguments = True
                            bInvalidFormat = True
                        End If
                    End If
                    
                    '//**** If the format is valid, read the next argument ****
                    If (isIndirectArray) Then
                        If (ArgumentIndex <= UBound(ArgumentsIndirect)) Then
                            vArgument = ArgumentsIndirect(ArgumentIndex)
                        Else
                            bNoArguments = True
                            bInvalidFormat = True
                        End If
                    Else
                        If (ArgumentIndex <= UBound(Arguments)) Then
                            vArgument = Arguments(ArgumentIndex)
                        Else
                            bNoArguments = True
                            bInvalidFormat = True
                        End If
                    End If
                    
                    If (Not bNoArguments) Then
                        ArgumentIndex = ArgumentIndex + 1
                        bInvalidFormat = False
                    End If
                        
                    '//**** If an argument was found, format it ****
                    If (Not bNoArguments) Then
                        Dim sValue As String
                        sValue = ""
                        
                        '//**** Set default precision value for this type of argument ****
                        If (Precision = -1) Then
                            Select Case ValueType
                                Case VALTYPE_STRING
                                    Precision = -1  '//**** No precision ****
                                
                                Case VALTYPE_COMPACT_EXPONENTIAL
                                    Precision = 5
                                
                                Case VALTYPE_FLOAT, VALTYPE_EXPONENTIAL
                                    Precision = 6
                                
                                Case VALTYPE_DECIMALSIGNED, VALTYPE_DECIMALUNSIGNED
                                    Precision = 1
                                    
                                Case Else
                                    Precision = 0
                            End Select
                        End If
                        
                        Select Case ValueType
                            Case VALTYPE_FLOAT
                                '//**** If argument is a string, convert to numeric ****
                                If (ConvertArgument(vArgument, vbDouble)) Then
                                    sValue = FormatNumber(vArgument, sPadding, Width, Precision, LeftJustify, Signed, ValueType, bNoUnsignificant)
                                Else
                                    sValue = "0"
                                End If
                            
                            '//**** %c : Single char ****
                            Case VALTYPE_CHAR
                                If (VarType(vArgument) = vbString) Then
                                    sValue = Left(CStr(vArgument), 1)
                                Else
                                    If (ConvertArgument(vArgument, vbByte)) Then
                                        sValue = Chr(vArgument)
                                    Else
                                        sValue = ""
                                    End If
                                End If
                            
                            '//**** %s : String ****
                            Case VALTYPE_STRING
                                If (ConvertArgument(vArgument, vbString)) Then
                                    sValue = vArgument
                                    
                                    If (Precision < Len(sValue) And (Precision > -1)) Then
                                        sValue = Left(sValue, Precision)
                                    End If
                                    
                                    If (Width > Len(sValue)) Then
                                        If (LeftJustify) Then
                                            sValue = sValue & String((Width - Len(sValue)), sPadding)
                                        Else
                                            sValue = String((Width - Len(sValue)), sPadding) & sValue
                                        End If
                                    End If
                                    
                                Else
                                    sValue = ""
                                End If
                                
                            '//**** %x, %X : Hex notation ****
                            Case VALTYPE_HEX
                                If (VarType(vArgument) = vbString) Then
                                    
                                    '//**** If the value is a string, return the binary structure of the string ****
                                    sValue = BinaryToHEX(CStr(vArgument), (Width > 0))
                                    
                                Else
                                    If (ConvertArgument(vArgument, vbDouble)) Then
                                        sValue = FormatNumber(vArgument, sPadding, Width, Precision, LeftJustify, False, VALTYPE_HEX, bNoUnsignificant)
                                    End If
                                End If
                                
                                If (bUpper) Then sValue = UCase(sValue) Else sValue = LCase(sValue)
                            
                            '//**** %o : Octal notation ****
                            Case VALTYPE_OCTAL
                                If (ConvertArgument(vArgument, vbLong)) Then
                                    sValue = FormatNumber(vArgument, sPadding, Width, Precision, LeftJustify, False, VALTYPE_OCTAL, bNoUnsignificant)
                                End If
                        
                            '//**** %d : Decimal notation (Signed) ****
                            Case VALTYPE_DECIMALSIGNED
                                If (ConvertArgument(vArgument, vbLong)) Then
                                    sValue = FormatNumber(vArgument, sPadding, Width, Precision, LeftJustify, Signed, VALTYPE_DECIMALSIGNED, bNoUnsignificant)
                                Else
                                    sValue = "0"
                                End If
                                
                            '//**** %u : Decimal notation (Unsigned) ****
                            Case VALTYPE_DECIMALUNSIGNED
                                If (ConvertArgument(vArgument, vbLong)) Then
                                    Dim dValue As Double
                                    dValue = CDbl(vArgument)
                                    
                                    If (dValue < 0) Then
                                        dValue = (dValue + (4# * 1024# * 1024# * 1024#))
                                    End If
                                    
                                    sValue = FormatNumber(dValue, sPadding, Width, Precision, LeftJustify, Signed, VALTYPE_DECIMALUNSIGNED, bNoUnsignificant)
                                Else
                                    sValue = "0"
                                End If

                            '//**** %e, %E : Exponential notation ****
                            Case VALTYPE_EXPONENTIAL
                                If ConvertArgument(vArgument, vbDouble) Then
                                    sValue = FormatNumber(vArgument, sPadding, Width, Precision, LeftJustify, Signed, VALTYPE_EXPONENTIAL, bNoUnsignificant)
                                Else
                                    sValue = "0"
                                End If
                                
                                If (bUpper) Then sValue = UCase(sValue) Else sValue = LCase(sValue)
                            
                            '//**** %g, %G : Compact exponential notation ****
                            Case VALTYPE_COMPACT_EXPONENTIAL
                                If ConvertArgument(vArgument, vbDouble) Then
                                    sValue = FormatNumber(vArgument, sPadding, Width, Precision, LeftJustify, Signed, VALTYPE_COMPACT_EXPONENTIAL, bNoUnsignificant)
                                Else
                                    sValue = "0"
                                End If
                                
                                If (bUpper) Then sValue = UCase(sValue) Else sValue = LCase(sValue)
                            
                            '//**** %p : Pointer ****
                            Case VALTYPE_POINTER
                                If (ConvertArgument(vArgument, vbLong)) Then
                                    sValue = FormatNumber(vArgument, " ", Width, Precision, LeftJustify, False, VALTYPE_POINTER, False)
                                Else
                                    sValue = "0000:0000"
                                End If
                                
                            Case VALTYPE_BUFFERDUMP
                                sValue = ""
                                
                                If (VarType(vArgument) = vbString) Then
                                    Let Arguments(ArgumentIndex - 1) = sprintf
                                End If
                        End Select
                    End If
                Else
                    bInvalidFormat = True
                    
                    i = i - 1
                    sFormat = Left(sFormat, Len(sFormat) - 1)
                End If
                
                '//**** If the format valid, write the formated text to the output, _
                if not, write format string to the output ****
                If (Not bInvalidFormat) Then
                    sprintf = sprintf & sValue
                Else
                    sprintf = sprintf & sFormat
                End If
            End If
            
            
        Else
            If (sChar = "\") Then
                Dim sType As String
                sType = Mid(lpFormat, i + 1, 1)
                
                Dim bNotValid As Boolean
                bNotValid = False
                Select Case sType
                    Case "r"    '// Carriage Return
                        sprintf = sprintf & vbCr
                    
                    Case "n"    '// New line
                        sprintf = sprintf & vbLf
                    
                    Case "t"    '// Tab
                        sprintf = sprintf & Chr(9)
                    
                    Case "f", "b"
                        '//**** These escape charcters cannot be used in vb ****
                        
                    Case "\"    '// Slash char
                        sprintf = sprintf & "\"
                    
                    Case "%"
                        sprintf = sprintf & "%"
                    
                    Case Else
                        bNotValid = True
                End Select
            
                If (bNotValid) Then
                    sprintf = sprintf & sChar
                Else
                    i = i + 1
                End If
            Else
            
                sprintf = sprintf & sChar
            End If
        End If
    Next
End Function


Private Function IsValidFormat(sFormatString As String, ValueType As Long, bUpper As Boolean, Width As Long, Precision As Long, sPadding As String, LeftJustify As Boolean, Signed As Boolean, bNoUnsignificant As Boolean) As Boolean
    Dim sFormat As String
    sFormat = Mid(sFormatString, 2)
    
    '//**** Enumerate flags (0, +, -, " ", *, #) ****
    sPadding = Chr(32)
    LeftJustify = False
    Signed = False
    Width = -1
    bNoUnsignificant = False
    
    Dim i As Long
    For i = 1 To Len(sFormatString)
        If (InStr("0-+*#" & Chr(32), Mid(sFormat, i, 1))) Then
        
            Select Case Mid(sFormat, i, 1)
                Case "0"    '// Pad with zeros instead of spaces
                    sPadding = "0"
                
                Case "-"    '// Left justify
                    LeftJustify = True
                
                Case "+"    '// Always show sign
                    Signed = True
                    
                Case "*"    '// Field width specified by an argument
                    Width = -2
                
                Case "#"
                    bNoUnsignificant = True
                
            End Select
        Else
            Exit For
        End If
    Next
    
    '//**** Trim the format string where the flags ends ****
    sFormat = Mid(sFormat, i)
    
    '//**** Extract width and precision value ****
    Dim sParams As String
    sParams = Left(sFormat, Len(sFormat) - 1)
    sFormat = Right(sFormat, 1)
    
    Dim vParamValue As Variant
    If (Len(sParams) > 0) Then
        '//**** Find the precision separator, if it is present, it means _
        we have a width and precision value, if not, we have only a width value ****
        Dim nPos As Long
        nPos = InStr(sParams, ".")
        
        If (nPos <> 0) Then
            If (Width <> -2) Then
                vParamValue = Left(sParams, nPos - 1)
                If (ConvertArgument(vParamValue, vbLong)) Then
                    Width = Val(vParamValue)
                Else
                    Width = 0
                End If
            End If
            
            vParamValue = Mid(sParams, nPos + 1)
            If (ConvertArgument(vParamValue, vbLong)) Then
                Precision = Val(vParamValue)
            Else
                Precision = -1
            End If
        
        Else
            If (Width <> -2) Then
                vParamValue = sParams
                If (ConvertArgument(vParamValue, vbLong)) Then
                    Width = Val(vParamValue)
                Else
                    Width = 0
                End If
            End If
            Precision = -1
        End If
    Else
        If (Width <> -2) Then Width = 0
        Precision = -1
    End If
    
    '//**** Check the conversion character ****
    Select Case sFormat
        Case "c"    '// Single char
            ValueType = VALTYPE_CHAR
        
        Case "o"    '// Octal notation
            ValueType = VALTYPE_OCTAL
        
        Case "s"    '// String - Lower case
            ValueType = VALTYPE_STRING
        
        Case "x"    '// Hexadecimal notation - Lower case
            ValueType = VALTYPE_HEX
            bUpper = False
            
        Case "X"    '// Hexadecimal notation - Upper case
            ValueType = VALTYPE_HEX
            bUpper = True
        
        Case "f"    '// Fixed-point notation
            ValueType = VALTYPE_FLOAT
        
        Case "d", "i"    '// Decimal notation (Signed)
            ValueType = VALTYPE_DECIMALSIGNED
        
        Case "u"    '// Decimal notation (Unsigned)
            ValueType = VALTYPE_DECIMALUNSIGNED
        
        Case "e"    '// Exponential notation - Lower case "E"
            ValueType = VALTYPE_EXPONENTIAL
            bUpper = False
        
        Case "E"    '// Exponential notation - Upper case "E"
            ValueType = VALTYPE_EXPONENTIAL
            bUpper = True
        
        Case "g"    '// Compact exponential notation - Lower case "E"
            ValueType = VALTYPE_COMPACT_EXPONENTIAL
            bUpper = False
        
        Case "G"    '// Compact exponential notation - Upper case "E"
            ValueType = VALTYPE_COMPACT_EXPONENTIAL
            bUpper = True
        
        Case "p"    '// Address pointer
            ValueType = VALTYPE_POINTER
        
        Case "n"    '// Nothing printed, store the current output to the pointed buffer ****
            ValueType = VALTYPE_BUFFERDUMP
        
        Case Else   '// Unkown vonversion caracter
            IsValidFormat = False
            Exit Function
    End Select
    
    IsValidFormat = True
End Function
Private Function FormatNumber(Number As Variant, sPadding As String, FieldWidth As Long, Precision As Long, LeftJustify As Boolean, Signed As Boolean, ValueType As Long, bNoUnsignificant As Boolean) As String
    
    
    '//**** Check the number sign ****
    Dim sSign As String
    If (Signed) Then
        sSign = IIf(Number >= 0, "+", "-")
    Else
        If (Number < 0) Then sSign = "-"
    End If
    
    Dim sFormatChar As String
    If (bNoUnsignificant) Then
        sFormatChar = "#"
    Else
        sFormatChar = "0"
    End If
    
    
    Dim sOutput As String
    Select Case ValueType
        Case VALTYPE_POINTER
            sOutput = Left(Hex(Number), 8)
            sOutput = String(8 - Len(sOutput), "0") & sOutput
            sOutput = Left(sOutput, 4) & ":" & Mid(sOutput, 5)
            
        
        Case VALTYPE_HEX, VALTYPE_DECIMALSIGNED, VALTYPE_DECIMALUNSIGNED, VALTYPE_OCTAL
            '//--- Hexadecimal notation ---
            If (ValueType = VALTYPE_HEX) Then
                If (Precision > 520) Then Precision = 520
                sOutput = Hex(Number)
                
            '//--- Ocatal notation ---
            ElseIf (ValueType = VALTYPE_OCTAL) Then
                If (Precision > 520) Then Precision = 520
                sOutput = Oct(Number)
            
            '//--- Decimal notation ---
            Else
                If (Precision > 22) Then Precision = 22
                sOutput = CStr(Round(Number, 0))
            End If
            
            If (Precision > Len(sOutput)) Then
                sOutput = String(Precision - Len(sOutput), "0") & sOutput
            End If
            
            
            If (ValueType = VALTYPE_HEX And (Number <> 0) And (bNoUnsignificant)) Then
                sOutput = "0x" & sOutput
            End If
            
            sSign = ""
        
        Case VALTYPE_EXPONENTIAL, VALTYPE_COMPACT_EXPONENTIAL
            If (Precision > 500) Then Precision = 500
            
            '//**** Precision must be at least 1 ****
            If (Precision = 0) Then Precision = 1
            
            Dim bExponential As Boolean
            bExponential = True
            
            '//**** Define format string ****
            Dim FormatString As String
            If (ValueType = VALTYPE_COMPACT_EXPONENTIAL) Then
                FormatString = "0." & String(Precision - 1, sFormatChar) & "e+000"
            Else
                FormatString = "0." & String(Precision, sFormatChar) & "e+000"
            End If
            
            '//**** Format the number using exponential notation ****
            sOutput = Format(Abs(Number), FormatString)
            
            If (ValueType = VALTYPE_COMPACT_EXPONENTIAL) Then
                Dim lPower As Long
                lPower = Val(Right(sOutput, 3))
                
                Dim NumeralLen As Long
                NumeralLen = lPower
                
                If (Abs(Number) < 1) Then
                    
                    If (lPower <= 4) Then
                        
                        sOutput = Abs(Number)
                        bExponential = False
                    End If
                Else
                    If (NumeralLen < (Precision)) Then
                        Dim DecimalLen As Long
                        DecimalLen = (Len(Number) - 1) - NumeralLen
                        
                        '//**** Define format string ****
                        If (NumeralLen > Precision) Or (NumeralLen = Precision) Then
                            FormatString = "0"
                        Else
                            FormatString = "0." & String(Precision - NumeralLen, sFormatChar)
                        End If
                        
                        sOutput = Format(Abs(Number), FormatString)
                        bExponential = False
                    End If
                End If
                
            End If
            
            '//**** Replace the coma by a point ****
            If (InStrRev(sOutput, ",")) Then
                Mid(sOutput, InStr(sOutput, ","), 1) = "."
            End If
            
            If (Len(sOutput) < FieldWidth - Len(sSign)) And (sPadding = "0") Then
                sOutput = String(FieldWidth - Len(sOutput) - Len(sSign), sPadding) & sOutput
            End If
        
        Case Else
            If (Precision > 22) Then Precision = 22
            
            '//**** Define format string ****
            If (sPadding = "0") And (FieldWidth > (Precision + 1) And Not (LeftJustify)) Then
                FormatString = String(FieldWidth - Precision - 1 - Len(sSign), "0")
            Else
                FormatString = sFormatChar
            End If
            
            '//**** Add precision formating ****
            If (Precision > 0) Then
                FormatString = "0." & String(Precision, sFormatChar)
            End If
            
            '//**** Format the output ****
            sOutput = Format(Abs(Number), FormatString)
            
            
            '//**** Replace the coma by a point ****
            If (InStrRev(sOutput, ",")) Then
                Mid(sOutput, InStr(sOutput, ","), 1) = "."
            End If
    End Select
    
    If (Right(sOutput, 1) = ".") Then
        sOutput = Left(sOutput, Len(sOutput) - 1)
    End If
    
    '//**** Add sign and padding ****
    If (Len(sOutput) + Len(sSign) < FieldWidth) Then
        If (LeftJustify) Then
            FormatNumber = sSign & sOutput & Space(FieldWidth - Len(sOutput) - Len(sSign))
        Else
            FormatNumber = Space(FieldWidth - Len(sOutput) - Len(sSign)) & sSign & sOutput
        End If
    Else
        FormatNumber = sSign & sOutput
    End If
End Function
Public Function ConvertArgument(vArgument As Variant, ArgType As Long) As Boolean
    On Error Resume Next
    
    If ((ArgType <> vbString) And (VarType(vArgument) = vbString)) Then
        '//**** Replace the point by a coma ****
        If (InStrRev(vArgument, ".")) Then
            Mid(vArgument, InStr(vArgument, "."), 1) = ","
        End If
    End If
    
    Select Case ArgType
        Case vbLong: vArgument = CLng(vArgument)
        Case vbByte: vArgument = CByte(vArgument)
        Case vbString: vArgument = CStr(vArgument)
        Case vbBoolean: vArgument = CBool(vArgument)
        Case vbSingle: vArgument = CSng(vArgument)
        Case vbDouble: vArgument = CDbl(vArgument)
        Case vbDecimal: vArgument = CDec(vArgument)
        Case vbInteger: vArgument = CInt(vArgument)
        Case vbCurrency: vArgument = CCur(vArgument)
        Case Else
            ConvertArgument = False
            Exit Function
    End Select
    
    ConvertArgument = (Err = 0)
    On Error GoTo 0
End Function
Public Function BinaryToHEX(SInput As String, Optional AddSpaces As Boolean) As String
    Dim x As Long
    For x = 1 To Len(SInput)
        
        Dim sByteHex As String
        sByteHex = Hex(Asc(Mid(SInput, x, 1)))
        sByteHex = String(2 - Len(sByteHex), "0") & sByteHex
        
        BinaryToHEX = BinaryToHEX & sByteHex & IIf(AddSpaces, " ", "")
    Next
End Function
