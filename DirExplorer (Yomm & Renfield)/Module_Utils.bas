Attribute VB_Name = "Module_Utils"
Option Explicit

'# Recuperes un type GUID d'une chaîne GUID
Public Function GetGUID(CLSID As String) As GUID
    Call CLSIDFromString(StrPtr(CLSID), GetGUID)
End Function

'# Recuperes une interface avec IMalloc
Public Function isMalloc() As IMalloc
  Static im As IMalloc
  If (im Is Nothing) Then Call SHGetMalloc(im)
  Set isMalloc = im
End Function

Public Function LOWORD(dwValue As Long) As Integer
    MoveMemory LOWORD, dwValue, 2
End Function

'# Permet de SousClasser facilement les controles
Public Sub Subclass(hWnd As Long, NewProc As Long)
    Dim oldProc As Long

    If GetProp(hWnd, "OldProc") Then Exit Sub

    oldProc = SetWindowLong(hWnd, GWL_WNDPROC, NewProc)
    Call SetProp(hWnd, "OldProc", oldProc)
End Sub

Public Sub UnSubclass(hWnd As Long)
    Dim oldProc As Long

    oldProc = GetProp(hWnd, "OldProc")
    If oldProc Then
        Call SetWindowLong(hWnd, GWL_WNDPROC, oldProc)
        Call RemoveProp(hWnd, "OldProc")
    End If
End Sub

'# Recuperes le nom d'un type StrRet (Obtenu via GetDisplayNameOf)
'# http://msdn.microsoft.com/library/default.asp?url=/library/en-us/shellcc/platform/shell/reference/shlwapi/string/strrettobuf.asp
Public Function GetStrRet(lpStr As STRRET, pidlRel As Long) As String
    '# grrr point d'entree non trouve , MSDN m'aurait mentit ?? tant pis, je vais le coder moi même !!!!
            '    Dim Buf As String * MAX_PATH
            '    Call StrRetToBuf(lpStr, pidlRel, Buf, MAX_PATH)
            '    GetStrRet = GetStrFromBufferA(Buf)

    Dim lpsz As Long
    Dim uOffset As Long
    Select Case (lpStr.uType)
        '# Type Unicode
        '# le premier bit donne le pointeur vers la chaine : on doit allouer un espace mémoire, et le liberer
        Case STRRET_WSTR
            MoveMemory lpsz, lpStr.CStr(0), 4
            GetStrRet = GetStrFromPtrW(lpsz)
            Call CoTaskMemFree(lpsz)
        '# Type Ansi
        '# le premier bit donne l'offset du nom en memoire, d'après le PID
        Case STRRET_OFFSET
            MoveMemory uOffset, lpStr.CStr(0), 4
            GetStrRet = GetStrFromPtrA(pidlRel + uOffset)
        '# Type String
        '# pointeur vers la chaine, allouée
        Case STRRET_CSTR
            GetStrRet = GetStrFromPtrA(VarPtr(lpStr.CStr(0)))
    End Select
End Function

'# Renvoie le handle de l'imagelist du systeme
'# Utiliser SHGFI_SMALLICON ou SHGFI_LARGEICON pour IconSize
Public Function GetSystemImageList(uFlags As Long) As Long
    Dim sfi As SHFILEINFO
    GetSystemImageList = GetFileInfo("C:\", SHGFI_SYSICONINDEX Or uFlags, sfi)
End Function

Public Function FARPROC(pfn As Long) As Long
    FARPROC = pfn
End Function

Public Function BrowseDialog(hWnd As Long, sPrompt As String, ulFlags As BF_Flags, Optional pidlRoot As Long = 0, Optional pidlPreSel As Long = 0) As Long
  Dim bi As BROWSEINFO
  
  With bi
    .hwndOwner = hWnd
    .pidlRoot = pidlRoot '# PID de la  racine de l'arbre visualise
    .lpszTitle = sPrompt
    .ulFlags = ulFlags
    .lParam = pidlPreSel '# cette valeur sera passee a la fonction de rappel
    .lpfn = FARPROC(AddressOf BrowseCallbackProc)
  End With
  
  BrowseDialog = SHBrowseForFolder(bi)
End Function

Public Function BrowseCallbackProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal lParam As Long, ByVal lpData As Long) As Long
    Select Case uMsg
        Case BFFM_INITIALIZED
            '# On selectionne le repertoire prevu
            Call SendMessage(hWnd, BFFM_SETSELECTIONA, ByVal False, ByVal lpData)
    End Select
End Function

'# Cette fonction renvoie un chaine de caractere, retirant les caracteres nuls suivant celle-ci...
Public Function GetStrFromBufferA(sz As String) As String
    GetStrFromBufferA = Split(sz, vbNullChar)(0)
End Function

'# Renvoie une chaine ANSII depuis un pointeur de ANSII
Public Function GetStrFromPtrA(lpszA As Long) As String
  Dim sRtn As String
  sRtn = String$(lstrlenA(ByVal lpszA), 0)
  Call lstrcpyA(ByVal sRtn, ByVal lpszA)
  GetStrFromPtrA = sRtn
End Function

'# Renvoie une chaine ANSII depuis un pointeur de UNICODE
Public Function GetStrFromPtrW(lpszW As Long) As String
  Dim sRtn As String
  '# 2 bits / char
  sRtn = String$(lstrlenW(ByVal lpszW) * 2, 0)
  Call WideCharToMultiByte(CP_ACP, 0, ByVal lpszW, -1, ByVal sRtn, Len(sRtn), 0, 0)
  GetStrFromPtrW = GetStrFromBufferA(sRtn)
End Function

'# Recuperes le handle du proprio du controle
Public Function GetTopLevelParent(hWnd As Long) As Long
    Dim hwndParent As Long
    Dim hwndTmp As Long
  
    hwndParent = hWnd
    Do
        hwndTmp = GetParent(hwndParent)
        If hwndTmp Then hwndParent = hwndTmp
    Loop While hwndTmp

    GetTopLevelParent = hwndParent
End Function

'# Ces deux fonctions permettent de passer des icones de base 0 (imagelist) aux indices des icones d'etat.
Public Function INDEXTOSTATEIMAGEMASK(iIndex As Long) As Long
    INDEXTOSTATEIMAGEMASK = iIndex * (2 ^ 12)
End Function

Public Function STATEIMAGEMASKTOINDEX(iState As Long) As Long
    STATEIMAGEMASKTOINDEX = iState / (2 ^ 12)
End Function


