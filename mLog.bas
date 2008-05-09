Attribute VB_Name = "mLog"
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

Public LogFileNumber As Integer

'Ouverture du fichier de log
Public Function OpenLogFile(Optional ClearFile As Boolean = True)
    On Error GoTo erreur:

    LogFileNumber = FreeFile
    If ClearFile Then
        Open App.path + "\log.txt" For Output As LogFileNumber
    Else
        Open App.path + "\log.txt" For Append As LogFileNumber
    End If
    
    Exit Function

'Gestion d'erreur
erreur:
    LogFileNumber = 0
End Function


'Fermeture du fichier de log
Public Function CloseLogFile()
    If LogFileNumber <> 0 Then Close LogFileNumber
End Function

