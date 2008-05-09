Attribute VB_Name = "mStart"
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
Public Sub Main()
    Dim Fichier As String, OpenMain As Boolean
    
    OpenMain = True
    
    'Initialisation de la localisation
    'Si aucune langue n'a été sauvegardée ou si le fichier n'existe pas
    If (GetSetting("UWC", "Language", "Current", "NotSet") = "NotSet") _
       Or FileExists(App.path + "\" + GetSetting("UWC", "Language", "Current", "NotSet")) = False Then
       
        ChLang.Show vbModal
    End If
    
    'Chargement de la localisation
    LoadTranslations GetSetting("UWC", "Language", "Current")
    
    'Gestion des paramètres en ligne de commande
    If Len(Command) Then
        Select Case CmdLineParam(Command, 1)
            Case "-logfile"
                OpenLogFile True
                
            Case "-logfilekeep"
                OpenLogFile False
                
            Case Else
                Fichier = CmdLineParam(Command, 1)
                
                If FileExists(Fichier) Then  'Le fichier existe
                    
                    OpenFile Fichier
                    OpenMain = False
                    
                Else 'fichier inexistant
                
                    MsgBox sprintf(GetTranslation(13, "Le fichier que vous tentez d'ouvrir n'existe pas ! \r\n\n Fichier : %s"), Fichier), vbCritical + vbOKOnly, GetTranslation(14, "Erreur : Fichier inexistant !")
                    End
                
                End If
                
        End Select
    End If

    If OpenMain Then MainForm.Show
End Sub
