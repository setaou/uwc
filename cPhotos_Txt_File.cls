VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPhotos_Txt_File"
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

Private Type tPicture
    ID As String
    Position As Long
    title As String
    Caption As String
    Credit As String
    Url_Photo As String
    Url_Thumb As String
    Source As String
End Type

Private Pictures() As tPicture

Private Type tAlbum
    ID As String
    title As String
    Photo_Count As Integer
    In_Use As Boolean
    Description As String
    Topic_Name As String
    Create_Date As Long
    Edit_Date As Long
End Type

Private Album As tAlbum

Private CurrentIndex As Integer
Private CurrentFile As String

Public Function OpenFilePhotos(ByVal File As String) As Integer
    Dim f As Integer, line As String
    Dim field_start As Integer, field_end As Integer, index_start As Integer, index_end As Integer, _
        data_start As Integer, data_end As Integer
    Dim Field As String, index As Integer, data As String
    
    'Ne ré-ouvre pas un fichier deja ouvert
    If File = CurrentFile Then Exit Function
    
    f = FreeFile
    
    Open File For Input As #f
    
        Line Input #f, line
        
        If (Mid(line, 1, InStr(1, line, "||") - 1) = "count") And (CInt(Mid(line, InStr(1, line, "||") + 2, 4)) > 0) Then 'Si le fichier est valide
            
            'Redimentionne le tableau en le vidant
            ReDim Pictures(0 To CInt(Mid(line, InStr(1, line, "||") + 2, 4)))
            
            Do 'Boucle de parcours du fichier
                Line Input #f, line
                
                'Si la ligne n'est pas vide
                If Len(line) > 0 And InStr(line, "||") Then
                    
                    'Coordonnées des différents champs contenus dans la ligne
                    index_end = InStr(line, "||") - 1
                    index_start = InStrRev(line, "_", index_end) + 1
                    field_end = index_start - 2
                    field_start = 1
                    data_start = index_end + 3
                    data_end = Len(line)
                    
                    'Extraction des valeurs
                    Field = Mid(line, field_start, field_end - field_start + 1)
                    index = CInt(Mid(line, index_start, index_end - index_start + 1))
                    data = Mid(line, data_start, data_end - data_start + 1)
                    
                    'Répartition des valeurs dans le tableau
                    With Pictures(index)
                        Select Case Field
                            Case "ID"
                                .ID = data
                            Case "POSITION"
                                .Position = CLng(data)
                            Case "TITLE"
                                .title = data
                            Case "CAPTION"
                                .Caption = data
                            Case "CREDIT"
                                .Credit = data
                            Case "URL_PHOTO"
                                .Url_Photo = LCase(data)
                            Case "URL_THUMB"
                                .Url_Thumb = LCase(data)
                            Case "SOURCE"
                                .Source = data
                        End Select
                    End With
                End If
            Loop Until EOF(f)
            
            CurrentFile = File
            OpenFilePhotos = 1
        
        Else 'Si le fichier est invalide
            
            OpenFilePhotos = -1
        
        End If
        
    Close #f
End Function

Public Function OpenFileAlbum(ByVal File As String) As Integer
    Dim f As Integer, line As String
    Dim field_start As Integer, field_end As Integer, index_start As Integer, index_end As Integer, _
        data_start As Integer, data_end As Integer
    Dim Field As String, index As Integer, data As String
    
    'Ne ré-ouvre pas un fichier deja ouvert
    If File = CurrentFile Then Exit Function
    
    f = FreeFile
    
    Open File For Input As #f
    
        Line Input #f, line
        
        If Mid(line, 1, InStr(1, line, "||") - 1) = "count" Then 'Si le fichier est valide
            
            Do 'Boucle de parcours du fichier
                Line Input #f, line
                
                'Si la ligne n'est pas vide
                If Len(line) > 0 And InStr(line, "||") Then
                    
                    'Coordonnées des différents champs contenus dans la ligne
                    index_end = InStr(line, "||") - 1
                    index_start = InStrRev(line, "_", index_end) + 1
                    field_end = index_start - 2
                    field_start = 1
                    data_start = index_end + 3
                    data_end = Len(line)
                    
                    'Extraction des valeurs
                    Field = Mid(line, field_start, field_end - field_start + 1)
                    data = Mid(line, data_start, data_end - data_start + 1)
                    
                    'Répartition des valeurs dans le tableau
                    With Album
                        Select Case Field
                            Case "ID"
                                .ID = CLng(data)
                            Case "TITLE"
                                .title = data
                            Case "PHOTO_COUNT"
                                .Photo_Count = CInt(data)
                            Case "IN_USE"
                                .In_Use = IIf(Field = "1", True, False)
                            Case "DESCRIPTION"
                                .Description = data
                            Case "TOPIC_NAME"
                                .Topic_Name = data
                            Case "CREATE_DATE"
                                .Create_Date = CLng(data)
                            Case "EDIT_DATE"
                                .Edit_Date = CLng(data)
                        End Select
                    End With
                End If
            Loop Until EOF(f)
            
            OpenFileAlbum = 1
        
        Else 'Si le fichier est invalide
            
            OpenFileAlbum = -1
        
        End If
        
    Close #f
End Function


Public Function SeekPicture(FileName As String) As Boolean
    Dim i As Integer
    
    'Si aucun fichier chargé, on quitte
    If Len(CurrentFile) = 0 Then Exit Function
    
    'Formattage du nom de fichier
    FileName = LCase(FileName)
    
    For i = 0 To UBound(Pictures) 'Parcours du tableau
    
        'S'arrete a la premiere occurence valide
        If Pictures(i).Url_Photo = FileName Then
            CurrentIndex = i
            Exit For
        End If
        
        'Si aucun enregistrement n'est trouvé, CurrentIndex passe à -1
        CurrentIndex = -1
        
    Next i
End Function


Public Property Get AlbumTitle() As String
    AlbumTitle = Album.title
End Property
Public Property Get AlbumDescription() As String
    AlbumDescription = Album.Description
End Property
Public Property Get AlbumTopicName() As String
    AlbumTopicName = Album.Topic_Name
End Property
Public Property Get AlbumInUse() As Boolean
    AlbumInUse = Album.In_Use
End Property
Public Property Get AlbumPhotoCount() As Integer
    AlbumPhotoCount = Album.Photo_Count
End Property
Public Property Get AlbumID() As String
    AlbumID = Album.ID
End Property

Private Sub Class_Initialize()
    CurrentIndex = -1
End Sub

Public Property Get PictureID() As String
    If CurrentIndex > -1 Then
        PictureID = Pictures(CurrentIndex).ID
    Else
        PictureID = 0
    End If
End Property

Public Property Get PicturePosition() As Long
    If CurrentIndex > -1 Then
        PicturePosition = Pictures(CurrentIndex).Position
    Else
        PicturePosition = -1
    End If
End Property

Public Property Get PictureTitle() As String
    If CurrentIndex > -1 Then
        PictureTitle = Pictures(CurrentIndex).title
    Else
        PictureTitle = ""
    End If
End Property

Public Property Get PictureCaption() As String
    If CurrentIndex > -1 Then
        PictureCaption = Pictures(CurrentIndex).Caption
    Else
        PictureCaption = ""
    End If
End Property

Public Property Get PictureCredit() As String
    If CurrentIndex > -1 Then
        PictureCredit = Pictures(CurrentIndex).Credit
    Else
        PictureCredit = ""
    End If
End Property

Public Property Get PictureURL_Photo() As String
    If CurrentIndex > -1 Then
        PictureURL_Photo = Pictures(CurrentIndex).Url_Photo
    Else
        PictureURL_Photo = ""
    End If
End Property

Public Property Get PictureURL_Thumb() As String
    If CurrentIndex > -1 Then
        PictureURL_Thumb = Pictures(CurrentIndex).Url_Thumb
    Else
        PictureURL_Thumb = ""
    End If
End Property

Public Property Get GetCurrentFile() As String
    GetCurrentFile = CurrentFile
End Property

Public Property Get PictureSource() As String
    If CurrentIndex > -1 Then
        PictureSource = Pictures(CurrentIndex).Source
    Else
        PictureSource = -1
    End If
End Property

Public Function AutoOpen(Directory As String) As Integer
    If Not (Right(Directory, 1) = "\") Then Directory = Directory + "\"
    
    'RAZ des données album
    Dim RazAlbum As tAlbum
    Album = RazAlbum
    
    If FileExists(Directory + "photos.txt") Then
        AutoOpen = OpenFilePhotos(Directory + "photos.txt")
    
        If FileExists(Directory + "album.txt") Then OpenFileAlbum Directory + "album.txt"
    Else
        AutoOpen = -1
    End If
End Function
