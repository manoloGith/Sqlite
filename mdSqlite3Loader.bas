Attribute VB_Name = "mdSqlite3Loader"
'=========================================================================
' mdSqlite3Loader - locates and loads sqlite3.dll (SQLite3 Multiple
' Ciphers) by full path BEFORE the first Declare call, and explains why
' it failed when it cannot be loaded.
'
' The Declares in sqlite3win32stubs use Lib "sqlite3.dll". Windows binds
' them to an already loaded module with that base name, so pre-loading the
' right file here makes them work even when the DLL is not on the normal
' search path (typical in the twinBASIC IDE, whose process lives in the
' IDE folder, not in the project folder).
'
' Folders searched, in order (first sqlite3.dll found wins):
'   1. the path passed to Sqlite3EnsureLoaded (file or folder)
'   2. App.Path\x64 (Win64) or App.Path\x86 (Win32)
'   3. App.Path
'   4. current directory \x64 or \x86, then current directory
' Keeping each bitness in its own subfolder lets both DLLs (both named
' sqlite3.dll) live next to the project/exe.
'=========================================================================
Option Explicit

Private Const LOAD_WITH_ALTERED_SEARCH_PATH As Long = &H8
Private Const INVALID_FILE_ATTRIBUTES       As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY      As Long = &H10
Private Const MAX_PATH_W                    As Long = 32768

Private Declare PtrSafe Function LoadLibraryExW Lib "kernel32" (ByVal lpLibFileName As LongPtr, ByVal hFile As LongPtr, ByVal dwFlags As Long) As LongPtr
Private Declare PtrSafe Function FreeLibrary Lib "kernel32" (ByVal hLibModule As LongPtr) As Long
Private Declare PtrSafe Function GetModuleHandleW Lib "kernel32" (ByVal lpModuleName As LongPtr) As LongPtr
Private Declare PtrSafe Function GetModuleFileNameW Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpFilename As LongPtr, ByVal nSize As Long) As Long
Private Declare PtrSafe Function GetProcAddress Lib "kernel32" (ByVal hModule As LongPtr, ByVal lpProcName As String) As LongPtr
Private Declare PtrSafe Function GetFileAttributesW Lib "kernel32" (ByVal lpFileName As LongPtr) As Long

Private m_hModule                   As LongPtr
Private m_sLoadError                As String

'--- True when sqlite3.dll (SQLite3MC, matching bitness) is loaded; safe to call repeatedly
Public Function Sqlite3EnsureLoaded(Optional ByVal DllPath As String) As Boolean
    Dim cFolders        As Collection
    Dim vFolder         As Variant
    Dim sFile           As String
    Dim sTried          As String
    Dim hModule         As LongPtr

    If m_hModule <> 0 Then
        Sqlite3EnsureLoaded = True
        Exit Function
    End If
    m_sLoadError = vbNullString
    '--- already in the process (loaded by someone else or found by Windows)
    hModule = GetModuleHandleW(StrPtr("sqlite3.dll"))
    If hModule <> 0 Then
        Sqlite3EnsureLoaded = pvAccept(hModule, pvModulePath(hModule))
        Exit Function
    End If
    Set cFolders = New Collection
    If Len(DllPath) > 0 Then
        If pvIsFolder(DllPath) Then
            cFolders.Add DllPath
        Else
            sFile = DllPath
        End If
    End If
    cFolders.Add pvAppPath() & pvArchFolder()
    cFolders.Add pvAppPath()
    cFolders.Add CurDir$ & pvArchFolder()
    cFolders.Add CurDir$
    If Len(sFile) = 0 Then
        For Each vFolder In cFolders
            If Len(vFolder) > 0 Then
                If pvFileExists(pvAddSlash(vFolder) & "sqlite3.dll") Then
                    sFile = pvAddSlash(vFolder) & "sqlite3.dll"
                    Exit For
                End If
                sTried = sTried & vbCrLf & "  " & pvAddSlash(vFolder) & "sqlite3.dll"
            End If
        Next
    End If
    If Len(sFile) = 0 Then
        '--- last resort: the standard Windows DLL search order
        hModule = LoadLibraryExW(StrPtr("sqlite3.dll"), 0, 0)
        If hModule = 0 Then
            m_sLoadError = "No se encuentra sqlite3.dll (" & pvBitness() & "). Rutas probadas:" & sTried
            Exit Function
        End If
        Sqlite3EnsureLoaded = pvAccept(hModule, pvModulePath(hModule))
        Exit Function
    End If
    hModule = LoadLibraryExW(StrPtr(sFile), 0, LOAD_WITH_ALTERED_SEARCH_PATH)
    If hModule = 0 Then
        m_sLoadError = "No se puede cargar " & sFile & ": " & pvDescribeError(Err.LastDllError)
        Exit Function
    End If
    Sqlite3EnsureLoaded = pvAccept(hModule, sFile)
    If Not Sqlite3EnsureLoaded Then
        '--- do not leave a wrong sqlite3.dll for the Declares to bind to
        Call FreeLibrary(hModule)
    End If
End Function

'--- reason of the last failed Sqlite3EnsureLoaded (empty when loaded)
Public Property Get Sqlite3LoadError() As String
    Sqlite3LoadError = m_sLoadError
End Property

'--- full path of the loaded sqlite3.dll (empty when not loaded)
Public Property Get Sqlite3LoadedPath() As String
    If m_hModule <> 0 Then
        Sqlite3LoadedPath = pvModulePath(m_hModule)
    End If
End Property

'--- private helpers ------------------------------------------------------

Private Function pvAccept(ByVal hModule As LongPtr, sFile As String) As Boolean
    '--- a plain SQLite build loads fine but lacks the cipher API
    If GetProcAddress(hModule, "sqlite3_libversion") = 0 Then
        m_sLoadError = sFile & " no es una DLL de SQLite"
        Exit Function
    End If
    If GetProcAddress(hModule, "sqlite3mc_version") = 0 Then
        m_sLoadError = sFile & " es SQLite sin cifrado; se necesita la DLL de SQLite3 Multiple Ciphers renombrada a sqlite3.dll"
        Exit Function
    End If
    m_hModule = hModule
    pvAccept = True
End Function

Private Function pvDescribeError(ByVal lErr As Long) As String
    Select Case lErr
    Case 193
        pvDescribeError = "error 193, la DLL es de otra arquitectura. Este proceso es de " & pvBitness() & _
            "; use sqlite3mc_" & IIf(pvIs64(), "x64", "x86") & ".dll renombrada a sqlite3.dll"
    Case 126
        pvDescribeError = "error 126, falta una DLL de la que depende (p.ej. el runtime de Visual C++: " & _
            "instale Microsoft Visual C++ Redistributable " & IIf(pvIs64(), "x64", "x86") & ")"
    Case 5
        pvDescribeError = "error 5, acceso denegado (fichero bloqueado por Windows o antivirus; " & _
            "si se descargo de Internet: Propiedades > Desbloquear)"
    Case Else
        pvDescribeError = "error de Windows " & lErr
    End Select
End Function

Private Function pvModulePath(ByVal hModule As LongPtr) As String
    Dim sBuf            As String
    Dim lLen            As Long

    sBuf = String$(MAX_PATH_W, 0)
    lLen = GetModuleFileNameW(hModule, StrPtr(sBuf), MAX_PATH_W)
    pvModulePath = Left$(sBuf, lLen)
End Function

Private Function pvAppPath() As String
    On Error Resume Next
    pvAppPath = App.Path
End Function

Private Function pvIs64() As Boolean
#If Win64 Then
    pvIs64 = True
#End If
End Function

Private Function pvBitness() As String
    pvBitness = IIf(pvIs64(), "64 bits", "32 bits")
End Function

Private Function pvArchFolder() As String
    pvArchFolder = IIf(pvIs64(), "\x64", "\x86")
End Function

Private Function pvAddSlash(ByVal sFolder As String) As String
    If Right$(sFolder, 1) <> "\" Then
        sFolder = sFolder & "\"
    End If
    pvAddSlash = sFolder
End Function

Private Function pvFileExists(sFile As String) As Boolean
    Dim lAttr           As Long

    lAttr = GetFileAttributesW(StrPtr(sFile))
    pvFileExists = (lAttr <> INVALID_FILE_ATTRIBUTES) And ((lAttr And FILE_ATTRIBUTE_DIRECTORY) = 0)
End Function

Private Function pvIsFolder(sPath As String) As Boolean
    Dim lAttr           As Long

    lAttr = GetFileAttributesW(StrPtr(sPath))
    pvIsFolder = (lAttr <> INVALID_FILE_ATTRIBUTES) And ((lAttr And FILE_ATTRIBUTE_DIRECTORY) <> 0)
End Function
