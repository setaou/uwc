VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWBZ_File"
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


Private Const Code0 As String = "WWBB0000"
Private Const Key0 As Long = &HA4
Private Const Code1 As String = "WWBB1111"
Private Const Key1 As Long = &HF2

Private JPG_Bytes() As Byte
Private Description_Bytes() As Byte
Private Description_String As String

Private Type WBZ_File_Header
    WBZ_File_Marker As Long
    First_File_Offset As Long
End Type

Private Type WBZ_Included_File_Header
    Included_File_Marker As Long
    Header_Size As Long
    zUnknown As Long
    File_Name As String * 256
    File_Size As Long
    zUnknown1 As String * 264
End Type

Private Const WBZ_File_Marker As Long = &H6791AB43
Private Const WBZ_Included_File_Marker As Long = &H1082CDE1

Public Property Get PictureID() As String
    PictureID = ExtractInfo(Description_String, "id")
End Property

Public Property Get PictureTitle() As String
    PictureTitle = ExtractInfo(Description_String, "title")
End Property

Public Property Get PictureCredit() As String
    PictureCredit = ExtractInfo(Description_String, "credit")
End Property

Public Property Get PictureFileName() As String
    PictureFileName = ExtractInfo(Description_String, "filename")
End Property

Public Property Get PictureSource() As String
    PictureSource = ExtractInfo(Description_String, "source")
End Property

Public Property Get PictureRootCategory() As String
    PictureRootCategory = ExtractInfo(Description_String, "catTitle")
End Property

Public Property Get PictureChildCategory() As String
    PictureChildCategory = ExtractInfo(Description_String, "albumTitle")
End Property

Public Property Get PicturePosition() As Long
    PicturePosition = CLng(Val(ExtractInfo(Description_String, "position")))
End Property

Public Property Get PictureDailyDate() As String
    PictureDailyDate = ExtractInfo(Description_String, "dailyDateYYYYMMDD")
    PictureDailyDate = Mid(PictureDailyDate, 1, 4) & "-" & Mid(PictureDailyDate, 5, 2) & "-" & Mid(PictureDailyDate, 7, 2)
End Property

Public Function OpenFile(File As String) As Integer
    Dim f As Integer
    Dim FileHeader As WBZ_File_Header
    Dim IncludedFileHeader As WBZ_Included_File_Header
    Dim Codage As String * 8
    Dim i As Integer
    
    f = FreeFile
    Open File For Binary As #f
    
    Get #f, , FileHeader
    
    If FileHeader.WBZ_File_Marker = WBZ_File_Marker Then 'Type de fichier correct
    
        Seek #f, FileHeader.First_File_Offset + 1
       
        'Premier fichier inclus : Image JPEG
        Get #f, , IncludedFileHeader
        If IncludedFileHeader.Included_File_Marker = WBZ_Included_File_Marker Then
            Get #f, , Codage
            If Codage = Code0 Or Codage = Code1 Then
                ReDim JPG_Bytes(1 To IncludedFileHeader.File_Size - 8)
                Get #f, , JPG_Bytes
                JPG_Bytes = Decode(JPG_Bytes, Codage)
            Else
                Seek #f, FileHeader.First_File_Offset + IncludedFileHeader.Header_Size + 1
                ReDim JPG_Bytes(1 To IncludedFileHeader.File_Size)
                Get #f, , JPG_Bytes
            End If
        End If
        
        '2e fichier inclus : Description
        Get #f, , IncludedFileHeader
        If IncludedFileHeader.Included_File_Marker = WBZ_Included_File_Marker Then
            ReDim Description_Bytes(1 To IncludedFileHeader.File_Size)
            Get #f, , Description_Bytes
            For i = 1 To UBound(Description_Bytes)
                Description_String = Description_String + Chr(Description_Bytes(i))
            Next i
        End If
        
        OpenFile = 1
    
    Else 'Mauvais type de fichier
    
        OpenFile = -1
        
    End If
    
    Close #f
End Function
Private Sub ExportToFile(File_Path As String, Bytes() As Byte)
    Dim f As Integer
    
    f = FreeFile
    Open File_Path For Binary As #f
        Put #f, 1, Bytes
    Close #f
End Sub
Public Sub Save_JPG_Picture(ByVal File_Path As String)
    ExportToFile File_Path, JPG_Bytes()
End Sub
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

Public Property Get Description() As String
    Description = Description_String
End Property



Private Function ExtractInfo(From As String, Field As String) As String
    Dim start_position As Integer, end_position As Integer
    start_position = InStr(1, From, Field + "=")
    
    If start_position > 0 Then
        end_position = InStr(start_position, From, vbCrLf) - 1
        ExtractInfo = Mid(From, start_position + Len(Field) + 1, end_position - start_position - Len(Field))
    End If
End Function
