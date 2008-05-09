VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cEvalExpr"
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

Private Type tVar
    Nom As String
    Valeur As Variant
End Type

Private Variables() As tVar
Private Sub AddVar(ByVal VarName As String, ByVal Value As Variant)
    ReDim Preserve Variables(0 To UBound(Variables) + 1)
    
    With Variables(UBound(Variables))
        .Nom = VarName
        .Valeur = Value
    End With
End Sub

Public Function SetVar(ByVal VarName As String, ByVal Value As Variant) As Integer
    Dim i As Integer
    
    If Mid(VarName, 1, 1) = "%" Then VarName = Right(VarName, Len(VarName) - 1)
    
    For i = 1 To UBound(Variables)
        If Variables(i).Nom = VarName Then
            Variables(i).Valeur = Value
            SetVar = 1
            Exit Function
        End If
    Next i
    
    AddVar VarName, Value
End Function

Private Sub Class_Initialize()
    ReDim Variables(0)
End Sub



Public Sub ClearVars()
    ReDim Variables(0)
End Sub




Public Function GetVar(ByVal VarName As String) As Variant
    Dim i As Integer
    
    If Mid(VarName, 1, 1) = "%" Then VarName = Right(VarName, Len(VarName) - 1)
    
    For i = 1 To UBound(Variables)
        If Variables(i).Nom = VarName Then
            GetVar = Variables(i).Valeur
            Exit Function
        End If
    Next i
    
    GetVar = ""
End Function

Public Function Eval(ByVal Expr As String, Optional NotDefError As Boolean = False) As String
    'Si NotDefError est activé, la fonction revoie une variable vide si au moins
    'une des variables de l'expression n'est pas définie ou est vide.
    
    Dim i As Integer, Car As String, Niveau As Integer, Tampon As String, Sortie As String
    Dim Mode As Integer '0 = normal, 1 = variable
    Dim NonDef As Boolean
    
    '------ Juste pour le debug ------
    'Static DebugLevel As Integer
    'DebugLevel = DebugLevel + 1
    
    'Debug.Print Space(DebugLevel) + "> Etude de : " + Expr
    '---------------------------------
    
    Niveau = 0
    
    For i = 1 To Len(Expr)
        Car = Mid(Expr, i, 1)
        
        If Niveau > 0 Then Tampon = Tampon + Car
        
        If Mode = 0 Then 'Pas de mode particulier
        
            Select Case Car
                Case "("
                    Niveau = Niveau + 1
                Case "|"
                    If (Not NonDef) And Niveau = 0 Then Exit For
                    If NonDef And Niveau = 0 Then
                        Sortie = ""
                        NonDef = False
                    End If
                Case ")"
                    Niveau = Niveau - 1
                    If Niveau = 0 Then
                        Tampon = Eval(Left(Tampon, Len(Tampon) - 1), NotDefError)
                        If Len(Tampon) Then
                            Sortie = Sortie + Tampon
                        Else
                            NonDef = True
                        End If
                        
                        Tampon = ""
                    End If
                Case "%"
                    If Niveau = 0 Then
                        Mode = 1
                        Tampon = ""
                    End If
                Case Else
                    If Niveau = 0 Then Sortie = Sortie + Car
            End Select
            
        ElseIf Mode = 1 Then 'Détection d'un nom de variable et résolution
            Dim AC As Integer
            AC = Asc(Car)
            
            'Si c'est un caractère autorisé dans les noms de variables (- . 0-9 A-Z _ a-z)
            If AC = 45 Or AC = 46 Or (AC > 47 And AC < 58) Or (AC > 64 And AC < 91) Or AC = 95 Or (AC > 96 And AC < 122) Then
                
                Tampon = Tampon + Car
                
            End If
            'Si ce n'est pas un caractère autorisé ou si c'est la fin de l'expression
            If (Not (AC = 45 Or AC = 46 Or (AC > 47 And AC < 58) Or (AC > 64 And AC < 91) Or AC = 95 Or (AC > 96 And AC < 122))) Or i = Len(Expr) Then
            
                Dim ValVar As String
                ValVar = CStr(GetVar(Tampon))
                
                If Len(ValVar) = 0 And NotDefError Then
                    NonDef = True
                    'Debug.Print Space(DebugLevel) + "> variable %" + Tampon + " non définie !"
                Else
                    Sortie = Sortie + ValVar
                    'Debug.Print Space(DebugLevel) + "> variable %" + Tampon + " : " + ValVar
                End If
                
                Tampon = ""
                Mode = 0
                
                'Si ce n'est pas la fin de l'expression, recule d'un pas pour permettre le traitement
                'du caractère en cours
                If Not (i = Len(Expr)) Then i = i - 1
                
            End If
        End If
    Next i
    
    If NonDef Then
        Eval = ""
        'Debug.Print Space(DebugLevel) + "> Sortie : (pas de sortie) " + Eval
    Else
        Eval = Sortie
        'Debug.Print Space(DebugLevel) + "> Sortie : " + Eval
    End If

    '------ Juste pour le debug ------
    'DebugLevel = DebugLevel - 1
    '---------------------------------
End Function
