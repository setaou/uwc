Attribute VB_Name = "Module_IDL_Operations"
Option Explicit

'# Dossiers spéciaux de Windows
Public Function GetPIDLFromFolderID(hWnd As Long, nFolder As Long) As Long
    Dim pidl As Long
    If SHGetSpecialFolderLocation(hWnd, nFolder, pidl) >= 0 Then
        GetPIDLFromFolderID = pidl
    End If
End Function

Public Function GetFileInfo(ByVal pszPath As Variant, uFlags As Long, sfi As SHFILEINFO) As Long
    If (VarType(pszPath) = vbString) Then
        GetFileInfo = SHGetFileInfo(CStr(pszPath), 0, sfi, Len(sfi), uFlags)
    Else
        GetFileInfo = SHGetFileInfo(CLng(pszPath), 0, sfi, Len(sfi), uFlags Or SHGFI_PIDL)
    End If
End Function

'# L' IDL du bureau a une taille de 0...
Public Function IsDesktopPIDL(pidl As Long) As Boolean
    If pidl Then IsDesktopPIDL = (GetItemIDSize(pidl) = 0)
End Function

Public Function GetFileTypeNamePIDL(pidl As Long) As String
    Dim sfi As SHFILEINFO
    If SHGetFileInfo(pidl, 0, sfi, Len(sfi), SHGFI_PIDL Or SHGFI_TYPENAME) Then
        GetFileTypeNamePIDL = GetStrFromBufferA(sfi.szTypeName)
    End If
End Function

'# Renvoie l'index de l'icone (petite ou grande) du fichier specifié , dans l'ImageList Systeme
Public Function GetFileIconIndexPIDL(ByVal pszPath As Variant, uFlags As Long) As Long
    Dim sfi As SHFILEINFO
    If GetFileInfo(pszPath, SHGFI_SYSICONINDEX Or uFlags, sfi) Then
        GetFileIconIndexPIDL = sfi.iIcon
    End If
End Function

'# Renvoie la taille du premier element de la liste
'# API Compatible W2K existante.......
Public Function GetItemIDSize(ByVal pidl As Long) As Integer
    '# SHITEMID.cb => 2 bits , les deux premiers de la structure
    If pidl Then Call MoveMemory(GetItemIDSize, ByVal pidl, 2)
End Function

'# Renvoie le nombre d'element d'un PIDL
Public Function GetItemIDCount(ByVal pidl As Long) As Integer
  Dim NbItems As Integer
  
  '# si la taille renvoyée est 0, on arrete de compter
    Do While GetItemIDSize(pidl)
        pidl = GetNextItemID(pidl)
        NbItems = NbItems + 1
    Loop
    GetItemIDCount = NbItems
End Function

'# Decale le pointeur : renvoie le prochain item !
'# API Compatible W2K existante.......
Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer
  
  cb = GetItemIDSize(pidl)
  If cb Then GetNextItemID = pidl + cb
End Function

'# Renvoie la taille totale d'un IDL , avec tous ses sous elements compris , en bits
Public Function GetPIDLSize(ByVal pidl As Long) As Integer
    Dim cb As Integer
    '# j'ai pus constater a es depends que des erreurs survenaient (overflow) quand le PIDL donné était invalide....
    On Error GoTo Fin
  
    If pidl Then
        While pidl '# Tant que le terminateur n'est pas atteint....
            cb = cb + GetItemIDSize(pidl)
            pidl = GetNextItemID(pidl)
        Wend
        GetPIDLSize = cb + 2 '# Terminateur '0'
    End If
Fin:
End Function

'# Renvoie un element de la liste
Public Function GetItemID(ByVal pidl As Long, ByVal nItem As Integer) As Long
    Dim nCount As Integer
    Dim i As Integer
    Dim cb As Integer
    Dim pidlNew As Long
    
    '# On calcule le nombre d'elements dans la liste
    nCount = GetItemIDCount(pidl)
    If (nItem > nCount) Or (nItem = GIID_LAST) Then nItem = nCount
  
    '# On souhaite recuperer un element particulier de la liste.....
    For i = 1 To nItem - 1: pidl = GetNextItemID(pidl): Next
    
    cb = GetItemIDSize(pidl)
    
    '# On Alloue un nouvel emplacement pour stocker l'Item (plus 2 bits, pour le caractere nul de fin)
    pidlNew = isMalloc.Alloc(cb + 2)
    If pidlNew Then
        '# Copie les infos du pid (temporaire !) dans l'emplacement precedemment alloué, et renvoie le pointeur vers celui-ci
        MoveMemory ByVal pidlNew, ByVal pidl, cb
        MoveMemory ByVal pidlNew + cb, 0, 2
        
        GetItemID = pidlNew
    End If
End Function

'# Crée un PIDL, le met a 0 , a vous de penser a desallouer la memoire......
Public Function CreatePIDL(cb As Long) As Long
    Dim pidl As Long
    
    '# Alloue un nouvel espace mémoire
    pidl = isMalloc.Alloc(cb)
    If pidl Then
        FillMemory ByVal pidl, cb, 0 '# on met que des 0....
        CreatePIDL = pidl '# On renvoie l'adresse de l'ITEM : le PIDL
    End If
End Function

'# Fait une copie d'un Item (A vous de liberer la memoire ici allouée)
'# API Compatible W2K existante.......
Public Function CopyPIDL(pidl As Long) As Long
  Dim cb As Long
  Dim pidlNew As Long
  
  cb = GetPIDLSize(pidl)
  If cb Then
    pidlNew = CreatePIDL(cb)
    MoveMemory ByVal pidlNew, ByVal pidl, cb
    CopyPIDL = pidlNew
  End If
End Function

'# Libere l'emplacement memoire reservé pour un PIDL
'# API Compatible W2K existante.......
Public Sub FreePIDL(pidl As Long)
    On Error GoTo Out
    If pidl Then isMalloc.Free ByVal pidl
Out:
    If Err And (pidl <> 0) Then
        Call CoTaskMemFree(pidl)
    End If
    pidl = 0
End Sub

'# Copie et renvoie le le PIDL parent du PIDL passé en parametre
Public Function GetPIDLParent(pidl As Long, Optional fReturnDesktop As Boolean = False, Optional fFreeOldPidl As Boolean = False) As Long
    Dim nCount As Integer
    Dim pidl1 As Long
    Dim i As Integer
    Dim cb As Integer
    Dim pidlNew As Long
    
    nCount = GetItemIDCount(pidl)
    If (nCount = 0) And (fReturnDesktop = False) Then Exit Function
  
    '# On calcule la taille de la liste
    pidl1 = pidl
    For i = 1 To nCount - 1
        cb = cb + GetItemIDSize(pidl1)
        pidl1 = GetNextItemID(pidl1)
    Next i
  
    '# On alloue l'espace necessaire (+2 bits pour le 0 terminateur)
    pidlNew = isMalloc.Alloc(cb + 2)
  
    If pidlNew Then
        '# On copie la liste dans l'espace nouvellement alloué
        MoveMemory ByVal pidlNew, ByVal pidl, cb
        FillMemory ByVal pidlNew + cb, 2, 0
    
        '# On a demander a liberer l'espace mémoire.
        If fFreeOldPidl Then Call FreePIDL(pidl)
        GetPIDLParent = pidlNew
    End If
End Function

'# Concatene deux PIDS
'# API Compatible W2K existante.......
Public Function CombinePIDLs(pidl1 As Long, pidl2 As Long, Optional fFreePidl1 As Boolean = False, Optional fFreePidl2 As Boolean = False) As Long
    Dim cb1 As Integer
    Dim cb2 As Integer
    Dim pidlNew As Long

    If pidl1 Then
        cb1 = GetPIDLSize(pidl1)
        If cb1 Then cb1 = cb1 - 2
    End If
  
    If pidl2 Then
        cb2 = GetPIDLSize(pidl2)
        If cb2 Then cb2 = cb2 - 2
    End If

    pidlNew = CreatePIDL(cb1 + cb2 + 2)
    If (pidlNew) Then
        If cb1 Then MoveMemory ByVal pidlNew, ByVal pidl1, cb1
        If cb2 Then MoveMemory ByVal pidlNew + cb1, ByVal pidl2, cb2
          
        FillMemory ByVal pidlNew + cb1 + cb2, 2, 0
          
        If (pidl1 And fFreePidl1) Then isMalloc.Free ByVal pidl1
        If (pidl2 And fFreePidl2) Then isMalloc.Free ByVal pidl2
    End If
    CombinePIDLs = pidlNew
End Function

Public Function GetPathFromPIDL(pidl As Long) As String
    Dim sPath As String * MAX_PATH
    If SHGetPathFromIDList(pidl, sPath) Then
        GetPathFromPIDL = GetStrFromBufferA(sPath)
    End If
End Function

Public Function GetPidlFromPath(Path As String) As Long
    Dim pidlRoot As Long
 
    Dim Buf As String * MAX_PATH
 
    Call MultiByteToWideChar(CP_ACP, MB_PRECOMPOSED, Path, -1, Buf, MAX_PATH)
    Call isfDesktop.ParseDisplayName(ByVal 0&, ByVal 0&, Buf, 0&, pidlRoot, 0&)
 
    GetPidlFromPath = pidlRoot
    
End Function
