VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWBC_File_Picture"
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

Private BMP_Bytes() As Byte
Private JPG_Bytes() As Byte

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

Private Header As WBC_PB_Header
Private Sub ExportToFile(File_Path As String, Bytes() As Byte)
    Dim f As Long
    
    f = FreeFile
    Open File_Path For Binary As #f
    Put #f, 1, Bytes
    Close #f
End Sub

Public Sub Save_JPG_Picture(ByVal File_Path As String)
    ExportToFile File_Path, JPG_Bytes()
End Sub

Public Sub Save_BMP_Thumbnail(ByVal File_Path As String)
    ExportToFile File_Path, BMP_Bytes()
End Sub


Public Sub Constructor(ByRef BMPBytes() As Byte, ByRef JPGBytes() As Byte, _
    Original_Filename As String, Image_Title As String, Image_Description As String, _
    Image_Credits As String, Fit_To_Screen As Long, Picture_ID As String, _
    Child_Category As String, Root_Category As String, Daily_Date As Long)
    
    BMP_Bytes = BMPBytes
    JPG_Bytes = JPGBytes
    
    Header.Original_Filename = Original_Filename
    Header.Image_Title = Image_Title
    Header.Image_Description = Image_Description
    Header.Image_Credits = Image_Credits
    Header.Fit_To_Screen = Fit_To_Screen
    Header.Picture_ID = Right(Picture_ID, Len(Picture_ID) - 1)
    Header.Child_Category = Child_Category
    Header.Root_Category = Root_Category
    Header.Daily_Date = Daily_Date
End Sub


Public Property Get Original_Filename() As String
    Original_Filename = ReduceStr(Header.Original_Filename)
End Property
Public Property Get Image_Title() As String
    Image_Title = ReduceStr(Header.Image_Title)
End Property

Private Function ReduceStr(str As String)
    ReduceStr = Mid(str, 1, InStr(1, str, Chr(0), vbBinaryCompare) - 1)
End Function



Public Property Get Image_Description() As String
    Image_Description = ReduceStr(Header.Image_Description)
End Property
Public Property Get Image_Credits() As String
    Image_Credits = ReduceStr(Header.Image_Credits)
End Property

Public Property Get Picture_ID() As String
    Picture_ID = ReduceStr(Header.Picture_ID)
End Property
Public Property Get Child_Category() As String
    Child_Category = ReduceStr(Header.Child_Category)
End Property
Public Property Get Root_Category() As String
    Root_Category = ReduceStr(Header.Root_Category)
End Property
Public Property Get Fit_To_Screen() As Integer
    Fit_To_Screen = CInt(Header.Fit_To_Screen)
End Property

Public Property Get Daily_Date() As String
    Daily_Date = CStr(Header.Daily_Date)
    Daily_Date = Mid(Daily_Date, 1, 4) & "-" & Mid(Daily_Date, 5, 2) & "-" & Mid(Daily_Date, 7, 2)
End Property


