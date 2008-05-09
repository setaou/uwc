VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWBC_File"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'    Ultimate Webshots Converter 1
'    Copyright (C) 2007  Hervé "Setaou" BRY <uwc at apinc dot org>
'
'    This program is free software; you can redistribute it and/or modify
'    it under the terms of the GNU General Public License as published by
'    the Free Software Foundation; either version 2 of the License, or
'    (at your option) any later version.
'
'    This program is distributed in the hope that it will be useful,
'    but WITHOUT ANY WARRANTY; without even the implied warranty of
'    MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'    GNU General Public License for more details.
'
'    You should have received a copy of the GNU General Public License
'    along with this program; if not, write to the Free Software
'    Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA


Option Explicit

Private Type WBC_FileHeader
    WBC_File_Marker As Long
    First_PB_Offset As Long
    z_Unknown As Long
    File_Title As String * 89
End Type

Private Type WBC_Index_Entry
    PB_Offset As Long
    PB_Size As Long
    z_Unknown As Long
    PB_Addition_Date As Long
    z_Unknown2 As String * 24
End Type

Private Const WBC_File_Marker As Long = &H95FA16AB
Private Const PB_Marker As Long = &HF071CDE2

Private Const Code0 As String = "WWBB0000"
Private Const Key0 As Long = &HA4
Private Const Code1 As String = "WWBB1111"
Private Const Key1 As Long = &HF2

Private Pictures() As cWBC_File_Picture
Private FileTitle As String
Private Valid_PB_Number As Integer

Private Type WBC_PB_Header
    PB_Marker As Long
    Header_Size As Long
    PB_Size As Long
    Original_Filename As String * 256
    Image_Title As String * 128
    Image_Description As String * 256
    Image_Credits As String * 256
    Original_File_Extension As String * 8
    JPG_File_Size As Long
    BMP_File_Size As Long
    z_Unknown As String * 140
    Daily_Date As Long
    Addition_Date As Long
    Fit_To_Screen As Long
    Picture_ID As String * 128
    Child_Category As String * 96
    Root_Category As String * 788
End Type
Private Function ReduceStr(str As String)
    ReduceStr = Mid(str, 1, InStr(1, str, Chr(0), vbBinaryCompare) - 1)
End Function


Private Function Decode(Fichier() As Byte, Codage As String) As Byte()
    Dim cle As Long, i As Integer
    
    Select Case Codage
        Case Code0: cle = Key0
        Case Code1: cle = Key1
    End Select
    
    For i = 1 To 100
        Fichier(i) = (Fichier(i + 100) Xor (Not Fichier(i))) Xor cle
    Next i
    
    Decode = Fichier
End Function



Public Function OpenFile(ByVal File As String)
    Dim FileHeader As WBC_FileHeader
    Dim PB_Number As Long
    Dim Index_PB() As WBC_Index_Entry
    Dim PB_Header As WBC_PB_Header
    Dim Codage As String * 8
    Dim BMPBytes() As Byte, JPGBytes() As Byte

    Dim i As Integer, f As Integer
    
    'Effacement des anciens parametres
    Valid_PB_Number = 0
    ReDim Pictures(1 To 1)
    
    f = FreeFile
    Open File For Binary As #f
    
    'Header du fichier
    Get #f, 1, FileHeader
    
    If FileHeader.WBC_File_Marker = WBC_File_Marker Then 'Si le format de fichier est bon
        
        'Titre du fichier
        FileTitle = ReduceStr(FileHeader.File_Title)
        
        'Table d'Index
        Get #f, 2197, PB_Number
        If PB_Number > 0 Then ReDim Index_PB(1 To PB_Number)
        Get #f, , Index_PB
        
        'PBs
        For i = 1 To PB_Number
            Get #f, Index_PB(i).PB_Offset + 1, PB_Header
            With PB_Header
            
                If (.PB_Marker = PB_Marker) And (.PB_Size + Index_PB(i).PB_Offset <= LOF(f)) Then  'Si le header du PB semble valide et si il est entier
                    
                    Valid_PB_Number = Valid_PB_Number + 1
                    
                    Seek #f, Index_PB(i).PB_Offset + .Header_Size + 1
                    Get #f, , Codage
                    
                    If Codage = Code0 Or Codage = Code1 Then
                        ReDim JPGBytes(1 To .JPG_File_Size - 8)
                        Get #f, , JPGBytes
                        JPGBytes = Decode(JPGBytes, Codage)
                    Else
                        Seek #f, Index_PB(i).PB_Offset + .Header_Size + 1
                        ReDim JPGBytes(1 To .JPG_File_Size)
                        Get #f, , JPGBytes
                    End If
                    
                    If .BMP_File_Size > 0 Then
                        ReDim BMPBytes(1 To .BMP_File_Size)
                        Get #f, , BMPBytes
                    End If
                    
                    ReDim Preserve Pictures(1 To Valid_PB_Number)
                    Set Pictures(i) = New cWBC_File_Picture
                    Pictures(i).Constructor BMPBytes, JPGBytes, _
                        PB_Header.Original_Filename, PB_Header.Image_Title, _
                        PB_Header.Image_Description, PB_Header.Image_Credits, _
                        PB_Header.Fit_To_Screen, PB_Header.Picture_ID, _
                        PB_Header.Child_Category, PB_Header.Root_Category, PB_Header.Daily_Date
                    
                Else 'Si le Header du PB est invalide
                
                    'PB IGNORE
                    
                End If
                
            End With
        Next i
        
        OpenFile = 1
        
    Else 'Si le format de fichier est incorrect
    
        OpenFile = -1
        
    End If
    
    Close #f

End Function

Public Property Get Picture(Number As Integer) As cWBC_File_Picture
    Set Picture = Pictures(Number)
End Property

Public Property Get PictureCount() As Integer
    PictureCount = Valid_PB_Number
End Property



Public Property Get File_Title() As String
    File_Title = FileTitle
End Property

