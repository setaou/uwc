VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWB1_File"
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

Private Const Code0 As String = "WWBB0000"
Private Const Key0 As Long = &HA4
Private Const Code1 As String = "WWBB1111"
Private Const Key1 As Long = &HF2

Private Const NoCode As Long = &HE0FFD8FF 'Header JPG non codé
Private Const Code As Long = &H42425757 '"WWBB" en hexa

Private JPG_Bytes() As Byte

Public Function OpenFile(ByVal File As String) As Integer
    Dim f As Integer
    Dim Marker As Long
    Dim Codage As String * 8
    
    f = FreeFile
    
    Open File For Binary As #f
    
    Get #f, 1, Marker
    
    Select Case Marker
        Case Code 'Commence par WWBB => Codé
        
            ReDim JPG_Bytes(1 To LOF(f) - 8)
            Get #f, 1, Codage
            Get #f, , JPG_Bytes
            JPG_Bytes = Decode(JPG_Bytes, Codage)
            
            OpenFile = 1
            
        Case NoCode 'Commence par un header JPEG en clair
        
            ReDim JPG_Bytes(1 To LOF(f))
            Get #f, 1, JPG_Bytes
            
            OpenFile = 1
            
        Case Else 'ERREUR : Mauvais format de fichier
        
            OpenFile = -1
            
    End Select
    
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
