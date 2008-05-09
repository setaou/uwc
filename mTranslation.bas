Attribute VB_Name = "mTranslation"
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

Const NB_TRANSLATIONS = 512

Private Type tMetaData
    ID As String
    Value As String
End Type

Private Translations() As String
Public TransMetadata() As tMetaData

Public TransLoaded As Boolean
Public Function LoadTranslations(Optional ByVal File As String = "français.lng") As Boolean
    Dim f As Long, Ligne As String
    Dim ID As String, Value As String

    File = App.path & "\" & File
    
    'Efface les anciennes traductions et métadonnées
    ClearTranslation
    
    'Si le fichier existe
    If FileExists(File) Then
    
        'Ouverture
        f = FreeFile
        Open File For Input As #f
        
        'Pour chaque ligne du fichier
        Do
            Line Input #f, Ligne
            
            'Si la ligne n'est pas vide, n'est pas un commentaire et est valide
            If ((Len(Ligne) > 0) And (Left(Ligne, 2) <> "//") And (InStr(1, Ligne, "=") > 0)) Then
                
                'Sépare l'id et la valeur
                ID = Left(Ligne, InStr(1, Ligne, "=") - 2)
                Value = Right(Ligne, Len(Ligne) - InStr(1, Ligne, "=") - 1)
                
                'Si l'entrée est numérique
                If IsNumeric(ID) Then
                    'Ajoute la traduction
                    Translations(Val(ID)) = Value
                'Si l'entrée est textuelle
                Else
                    'Ajoute aux métadonnées
                    ReDim Preserve TransMetadata(0 To UBound(TransMetadata) + 1)
                    With TransMetadata(UBound(TransMetadata))
                        .ID = ID
                        .Value = Value
                    End With
                End If

            End If
        Loop Until EOF(f)
        
        Close #f
        
        SetGlobalTranslations
        
        'Chargement OK
        ReDim Preserve TransMetadata(0 To UBound(TransMetadata) + 1)
        With TransMetadata(UBound(TransMetadata))
            .ID = "file"
            .Value = GetFileName(File)
        End With
        TransLoaded = True
        
        LoadTranslations = True
        
    Else
    
        MsgBox "Unable to load the language file. Please Reinstall UWC !", vbCritical, "Cannot load language file"
        
        LoadTranslations = False
        
        End
        
    End If
    
End Function

Public Function GetTranslation(ID As Integer, Optional Text As String = "Missing translation") As String
    If Len(Translations(ID)) Then
        GetTranslation = Translations(ID)
    Else
        GetTranslation = sprintf("#%i# %s", CStr(ID), Text)
    End If
End Function

Public Sub SetGlobalTranslations()
    WebshotsExtensions = GetTranslation(254) & " (*.wb?)|*.wbc;*.wbz;*.wb1;*.wb0;*.wbd;*.wbp|" & GetTranslation(255) & " (*.wbc,*.wbp)|*.wbc;*.wbp|" & GetTranslation(256) & " (*.wb1,*.wb0,*.wbd)|*.wb1;*.wb0;*.wbd|" & GetTranslation(257) & " (*.wbz)|*.wbz"
End Sub

Public Function GetTransMetadata(ID As String) As String
    Dim i As Integer
    
    For i = 1 To UBound(TransMetadata)
        If TransMetadata(i).ID = ID Then GetTransMetadata = TransMetadata(i).Value
    Next i
End Function

Public Sub ClearTranslation()
    'Efface les anciennes traductions et métadonnées
    ReDim Translations(1 To NB_TRANSLATIONS)
    ReDim TransMetadata(0 To 0)
    TransLoaded = False
End Sub
