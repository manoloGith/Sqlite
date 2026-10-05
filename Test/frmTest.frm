VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Prueba SQLite3 Multiple Ciphers (twinBASIC)"
   ClientHeight    =   7800
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   10200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   8640
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "&Ejecutar pruebas"
      Default         =   -1  'True
      Height          =   375
      Left            =   6960
      TabIndex        =   3
      Top             =   120
      Width           =   1575
   End
   Begin VB.TextBox txtKey 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   ""
      Top             =   120
      Width           =   2655
   End
   Begin VB.CheckBox chkKeepFiles 
      Caption         =   "Conservar ficheros"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
   Begin VB.TextBox txtLog 
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   600
      Width           =   9975
   End
   Begin VB.Label lblKey 
      Caption         =   "Clave:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   180
      Width           =   1215
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=========================================================================
' frmTest - smoke test of the SQLite wrapper on SQLite3 Multiple Ciphers
' (sqlite3.dll). Build the project once as Win32 and once as Win64 and run
' "Ejecutar pruebas" in both: every line must read OK (or SKIP for a cipher
' not compiled into the DLL).
'=========================================================================
Option Explicit

Private Const BIG_INT64             As String = "9007199254740993"   ' 2^53 + 1, not representable as Double

Private m_lPassed                   As Long
Private m_lFailed                   As Long
Private m_lSkipped                  As Long
Private m_sDbFile                   As String

'--- events ---------------------------------------------------------------

Private Sub Form_Load()
    m_sDbFile = pvTempFolder() & "sqlite3mc_test.db"
    '--- non-ASCII on purpose: the key is passed to SQLite as UTF-8
    txtKey.Text = "Clave-Secreta-" & ChrW$(&HF1) & ChrW$(&H20AC)
#If Win64 Then
    Me.Caption = Me.Caption & " - 64 bits"
#Else
    Me.Caption = Me.Caption & " - 32 bits"
#End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then
        Exit Sub
    End If
    On Error Resume Next
    txtLog.Move 120, 600, Me.ScaleWidth - 240, Me.ScaleHeight - 720
End Sub

Private Sub cmdClear_Click()
    txtLog.Text = vbNullString
End Sub

Private Sub cmdRun_Click()
    Dim sKey            As String

    sKey = txtKey.Text
    If Len(sKey) = 0 Then
        MsgBox "Introduzca una clave", vbExclamation
        txtKey.SetFocus
        Exit Sub
    End If
    m_lPassed = 0
    m_lFailed = 0
    m_lSkipped = 0
    cmdRun.Enabled = False
    Screen.MousePointer = vbHourglass
    txtLog.Text = vbNullString
    pvLog String$(70, "=")
    pvLog "Inicio: " & Format$(Now, "yyyy-mm-dd hh:nn:ss") & "   BD: " & m_sDbFile
    pvLog String$(70, "=")

    '--- load sqlite3.dll by full path first; without it nothing else can run
    If Not Sqlite3EnsureLoaded() Then
        pvCheck False, "Cargar sqlite3.dll"
        pvLog vbNullString
        pvLog Sqlite3LoadError
        pvLog vbNullString
        pvLog "App.Path = " & App.Path
        pvLog "CurDir   = " & CurDir$
        Screen.MousePointer = vbDefault
        cmdRun.Enabled = True
        Exit Sub
    End If
    pvCheck True, "sqlite3.dll cargada: " & Sqlite3LoadedPath

    pvRunTest "Entorno y versiones", 1, sKey
    pvRunTest "Crear BD cifrada y tipos de datos", 2, sKey
    pvRunTest "Leer con la clave correcta", 3, sKey
    pvRunTest "Rechazar clave incorrecta / sin clave", 4, sKey
    pvRunTest "Transacciones y savepoints", 5, sKey
    pvRunTest "Funciones, agregados y colaciones VB", 6, sKey
    pvRunTest "Comandos preparados", 7, sKey
    pvRunTest "ReKey (cambiar clave y quitar cifrado)", 8, sKey
    pvRunTest "Todos los cifrados de SQLite3MC", 9, sKey
    pvRunTest "Copia de BD (backup API)", 10, sKey

    If chkKeepFiles.Value = vbUnchecked Then
        pvDeleteDb m_sDbFile
    End If
    pvLog String$(70, "=")
    pvLog "RESULTADO: " & m_lPassed & " OK, " & m_lFailed & " FALLOS, " & m_lSkipped & " omitidas"
    pvLog String$(70, "=")
    Screen.MousePointer = vbDefault
    cmdRun.Enabled = True
End Sub

'--- test driver ----------------------------------------------------------

Private Sub pvRunTest(sTitle As String, ByVal lTest As Long, sKey As String)
    pvLog vbNullString
    pvLog "--- " & lTest & ". " & sTitle
    On Error GoTo EH
    Select Case lTest
    Case 1: pvTestEnvironment
    Case 2: pvTestCreate sKey
    Case 3: pvTestRead sKey
    Case 4: pvTestBadKey sKey
    Case 5: pvTestTransactions sKey
    Case 6: pvTestUdf sKey
    Case 7: pvTestCommands sKey
    Case 8: pvTestReKey sKey
    Case 9: pvTestCiphers sKey
    Case 10: pvTestCopy sKey
    End Select
    Exit Sub
EH:
    pvCheck False, "Error inesperado " & Err.Number & " (&H" & Hex$(Err.Number) & ") en " & Err.Source & ": " & Err.Description
End Sub

'--- tests ----------------------------------------------------------------

Private Sub pvTestEnvironment()
    Dim oCnn            As cConnection
    Dim lIdx            As Long
    Dim lCount          As Long
    Dim sNames          As String
    Dim lpDummy         As LongPtr

#If Win64 Then
    pvLog "Plataforma: Win64 (LongPtr = " & LenB(lpDummy) & " bytes)"
    pvCheck LenB(lpDummy) = 8, "LongPtr de 8 bytes"
#Else
    pvLog "Plataforma: Win32 (LongPtr = " & LenB(lpDummy) & " bytes)"
    pvCheck LenB(lpDummy) = 4, "LongPtr de 4 bytes"
#End If
    Set oCnn = New cConnection
    pvCheck oCnn.CreateNewDB(":memory:"), "Abrir BD en memoria"
    pvLog "SQLite:  " & oCnn.Version
    pvLog "SQLite3MC: " & oCnn.CipherLibVersion
    pvCheck Len(oCnn.CipherLibVersion) > 0, "sqlite3.dll es SQLite3 Multiple Ciphers"
    lCount = stub_sqlite3mc_cipher_count()
    For lIdx = 1 To lCount
        sNames = sNames & IIf(Len(sNames) > 0, ", ", vbNullString) & FromUtf8Ptr(stub_sqlite3mc_cipher_name(lIdx))
    Next
    pvLog "Cifrados disponibles (" & lCount & "): " & sNames
    pvCheck lCount > 0, "Hay cifrados registrados"
    pvCheck oCnn.DBHdl <> 0, "DBHdl es un LongPtr no nulo"
End Sub

Private Sub pvTestCreate(sKey As String)
    Dim oCnn            As cConnection
    Dim vId             As Variant

    pvDeleteDb m_sDbFile
    Set oCnn = New cConnection
    pvCheck oCnn.CreateNewDB(m_sDbFile, sKey), "CreateNewDB con clave"
    oCnn.Execute "CREATE TABLE T (ID INTEGER PRIMARY KEY, Txt TEXT, Big INTEGER, Dbl REAL, Bin BLOB, Fecha TEXT, Nada TEXT)"
    oCnn.ExecCmd "INSERT INTO T (Txt, Big, Dbl, Bin, Fecha, Nada) VALUES (?, ?, ?, ?, ?, ?)", _
        pvUnicodeText(), CLngLng(BIG_INT64), 3.14159265358979, pvTestBlob(), pvTestDate(), Null
    vId = oCnn.LastInsertAutoID
    pvCheck vId = 1, "LastInsertAutoID = 1 (obtenido " & vId & ")"
    pvCheck VarType(vId) = vbLongLong, "LastInsertAutoID es LongLong (VarType " & VarType(vId) & ")"
    pvCheck oCnn.AffectedRows = 1, "AffectedRows = 1"
    Set oCnn = Nothing
    pvCheck Not pvFileIsPlainSqlite(m_sDbFile), "El fichero esta cifrado (sin cabecera 'SQLite format 3')"
End Sub

Private Sub pvTestRead(sKey As String)
    Dim oCnn            As cConnection
    Dim oRs             As cRecordset
    Dim baBlob()        As Byte

    Set oCnn = New cConnection
    pvCheck oCnn.OpenDB(m_sDbFile, sKey), "OpenDB con la clave correcta"
    Set oRs = oCnn.GetRs("SELECT * FROM T WHERE ID = ?", 1)
    pvCheck oRs.RecordCount = 1, "Una fila leida"
    pvCheck oRs.Fields("Txt").Value = pvUnicodeText(), "Texto Unicode intacto (UTF-8)"
    pvCheck CStr(oRs.Fields("Big").Value) = BIG_INT64, "int64 > 2^53 intacto (" & CStr(oRs.Fields("Big").Value) & ")"
    pvCheck Abs(oRs.Fields("Dbl").Value - 3.14159265358979) < 0.000000000001, "Double intacto"
    baBlob = oRs.Fields("Bin").Value
    pvCheck pvSameBytes(baBlob, pvTestBlob()), "Blob de 256 bytes intacto"
    pvCheck CDate(oRs.Fields("Fecha").Value) = pvTestDate(), "Fecha intacta (" & oRs.Fields("Fecha").Value & ")"
    pvCheck IsNull(oRs.Fields("Nada").Value) Or IsEmpty(oRs.Fields("Nada").Value), "NULL leido como Null/Empty"
    pvCheck oCnn.CheckIntegrity = "ok", "PRAGMA integrity_check = ok"
End Sub

Private Sub pvTestBadKey(sKey As String)
    Dim oCnn            As cConnection

    Set oCnn = New cConnection
    pvCheck Not oCnn.OpenDB(m_sDbFile, sKey & "X"), "OpenDB con clave incorrecta devuelve False"
    pvCheck oCnn.OpenFailedBadKey, "OpenFailedBadKey = True (codigo " & oCnn.OpenErrorCode & ": " & oCnn.OpenErrorMessage & ")"
    Set oCnn = New cConnection
    pvCheck Not oCnn.OpenDB(m_sDbFile), "OpenDB sin clave devuelve False"
    pvCheck oCnn.OpenErrorCode = SQLITE_NOTADB, "Codigo SQLITE_NOTADB (26)"
End Sub

Private Sub pvTestTransactions(sKey As String)
    Dim oCnn            As cConnection

    Set oCnn = pvOpen(sKey)
    oCnn.Execute "CREATE TABLE IF NOT EXISTS Tr (N INTEGER)"
    oCnn.Execute "DELETE FROM Tr"
    oCnn.BeginTrans
    oCnn.ExecCmd "INSERT INTO Tr VALUES (?)", 1
    oCnn.RollbackTrans
    pvCheck pvCount(oCnn, "Tr") = 0, "RollbackTrans deshace la insercion"
    oCnn.BeginTrans
    oCnn.ExecCmd "INSERT INTO Tr VALUES (?)", 1
    oCnn.BeginTrans "SP1"
    oCnn.ExecCmd "INSERT INTO Tr VALUES (?)", 2
    oCnn.RollbackTrans "SP1"
    oCnn.CommitTrans
    pvCheck pvCount(oCnn, "Tr") = 1, "Savepoint anidado: rollback parcial + commit"
    pvCheck oCnn.TransactionStackCounter = 0, "Pila de transacciones vacia"
End Sub

Private Sub pvTestUdf(sKey As String)
    Dim oCnn            As cConnection
    Dim oFunc           As cTestFunc
    Dim oAgg            As cTestAgg
    Dim oColl           As cTestColl
    Dim oRs             As cRecordset
    Dim lIdx            As Long
    Dim sText           As String

    Set oCnn = pvOpen(sKey)
    Set oFunc = New cTestFunc
    Set oAgg = New cTestAgg
    Set oColl = New cTestColl
    pvCheck oCnn.AddUserDefinedFunction(oFunc), "Registrar VB_UPPER/VB_ADD/VB_FAIL"
    pvCheck oCnn.AddUserDefinedAggregateFunction(oAgg), "Registrar VB_SUM"
    pvCheck oCnn.AddUserDefinedCollation(oColl), "Registrar colacion VB_REVERSE"

    '--- "nandu aou" with n-tilde, u-acute and umlauts
    sText = ChrW$(&HF1) & "and" & ChrW$(&HFA) & " " & ChrW$(&HE4) & ChrW$(&HF6) & ChrW$(&HFC)
    Set oRs = oCnn.GetRs("SELECT VB_UPPER(?) AS U, VB_ADD(2, 3.5) AS S", sText)
    pvCheck oRs.Fields("U").Value = UCase$(sText), "VB_UPPER (callback CDecl): " & oRs.Fields("U").Value
    pvCheck oRs.Fields("S").Value = 5.5, "VB_ADD(2, 3.5) = 5.5"

    oCnn.Execute "CREATE TEMP TABLE Nums (N INTEGER, S TEXT)"
    For lIdx = 1 To 10
        oCnn.ExecCmd "INSERT INTO Nums VALUES (?, ?)", lIdx, Chr$(64 + lIdx)
    Next
    Set oRs = oCnn.GetRs("SELECT VB_SUM(N) AS Total FROM Nums")
    pvCheck oRs.Fields("Total").Value = 55, "VB_SUM(1..10) = 55"
    Set oRs = oCnn.GetRs("SELECT S FROM Nums ORDER BY S COLLATE VB_REVERSE")
    pvCheck oRs.Fields("S").Value = "J", "ORDER BY ... COLLATE VB_REVERSE empieza por 'J'"

    On Error Resume Next
    Set oRs = oCnn.GetRs("SELECT VB_FAIL()")
    pvCheck Err.Number <> 0, "Un error VB en la UDF llega como error SQL: " & Err.Description
    On Error GoTo 0

    pvCheck oCnn.RemoveUserDefinedFunction(oFunc), "Quitar VB_UPPER/VB_ADD/VB_FAIL"
    On Error Resume Next
    Set oRs = oCnn.GetRs("SELECT VB_UPPER('x')")
    pvCheck Err.Number <> 0, "VB_UPPER ya no existe tras quitarla"
    On Error GoTo 0
End Sub

Private Sub pvTestCommands(sKey As String)
    Dim oCnn            As cConnection
    Dim oCmd            As cCommand
    Dim oSel            As cSelectCommand
    Dim oRs             As cRecordset
    Dim lIdx            As Long

    Set oCnn = pvOpen(sKey)
    oCnn.Execute "CREATE TEMP TABLE Cmd (ID INTEGER, Txt TEXT, Big INTEGER)"
    Set oCmd = oCnn.CreateCommand("INSERT INTO Cmd VALUES (?, ?, ?)")
    pvCheck oCmd.StmtHdl <> 0, "StmtHdl es un LongPtr no nulo"
    oCnn.BeginTrans
    For lIdx = 1 To 1000
        oCmd.SetInt32 1, lIdx
        oCmd.SetText 2, "Fila " & lIdx
        oCmd.SetInt64 3, CLngLng(BIG_INT64) + lIdx
        oCmd.Execute
    Next
    oCnn.CommitTrans
    pvCheck pvCount(oCnn, "Cmd") = 1000, "1000 inserciones con cCommand"
    Set oSel = oCnn.CreateSelectCommand("SELECT Big FROM Cmd WHERE ID = ?")
    oSel.SetInt32 1, 500
    Set oRs = oSel.Execute
    pvCheck CStr(oRs.Fields(0).Value) = CStr(CLngLng(BIG_INT64) + 500), "cSelectCommand + SetInt64: " & CStr(oRs.Fields(0).Value)
End Sub

Private Sub pvTestReKey(sKey As String)
    Dim oCnn            As cConnection
    Dim sNewKey         As String

    sNewKey = sKey & "-nueva"
    Set oCnn = pvOpen(sKey)
    oCnn.ReKey sNewKey
    Set oCnn = Nothing
    Set oCnn = New cConnection
    pvCheck Not oCnn.OpenDB(m_sDbFile, sKey), "La clave antigua ya no abre"
    pvCheck oCnn.OpenDB(m_sDbFile, sNewKey), "La clave nueva abre"
    oCnn.ReKey
    Set oCnn = Nothing
    pvCheck pvFileIsPlainSqlite(m_sDbFile), "ReKey """" deja el fichero sin cifrar"
    Set oCnn = New cConnection
    pvCheck oCnn.OpenDB(m_sDbFile), "Abre sin clave"
    pvCheck pvCount(oCnn, "T") = 1, "Datos conservados tras descifrar"
    oCnn.ReKey sKey
    Set oCnn = Nothing
    pvCheck Not pvFileIsPlainSqlite(m_sDbFile), "ReKey vuelve a cifrar"
End Sub

Private Sub pvTestCiphers(sKey As String)
    Dim oCnn            As cConnection
    Dim aNames          As Variant
    Dim lCodec          As Long
    Dim sFile           As String

    aNames = Array(vbNullString, "AES128", "AES256", "ChaCha20", "SQLCipher", "RC4", "Ascon128", "AEGIS")
    For lCodec = 1 To 7
        sFile = pvTempFolder() & "sqlite3mc_cipher_" & lCodec & ".db"
        pvDeleteDb sFile
        Set oCnn = New cConnection
        oCnn.CodecType = lCodec
        If Not oCnn.CreateNewDB(sFile, sKey) Then
            If InStr(oCnn.OpenErrorMessage, "not available") > 0 Then
                pvSkip aNames(lCodec) & ": no incluido en esta sqlite3.dll"
            Else
                pvCheck False, aNames(lCodec) & ": " & oCnn.OpenErrorMessage
            End If
        Else
            oCnn.Execute "CREATE TABLE C (V TEXT)"
            oCnn.ExecCmd "INSERT INTO C VALUES (?)", pvUnicodeText()
            Set oCnn = Nothing
            Set oCnn = New cConnection
            oCnn.CodecType = lCodec
            If oCnn.OpenDB(sFile, sKey) Then
                pvCheck oCnn.GetRs("SELECT V FROM C").Fields(0).Value = pvUnicodeText(), aNames(lCodec) & ": cifrar, cerrar, reabrir y leer"
            Else
                pvCheck False, aNames(lCodec) & ": no reabre (" & oCnn.OpenErrorMessage & ")"
            End If
            Set oCnn = Nothing
            pvCheck Not pvFileIsPlainSqlite(sFile), aNames(lCodec) & ": fichero cifrado"
        End If
        Set oCnn = Nothing
        If chkKeepFiles.Value = vbUnchecked Then
            pvDeleteDb sFile
        End If
    Next
End Sub

Private Sub pvTestCopy(sKey As String)
    Dim oCnn            As cConnection
    Dim oCopy           As cConnection
    Dim sFile           As String
    Dim baBlob()        As Byte

    Set oCnn = pvOpen(sKey)
    '--- encrypted -> plain :memory: (SQLite3MC rejects the backup API here,
    '--- CopyDatabase falls back to a schema + row copy)
    Set oCopy = oCnn.CopyDatabase(":memory:")
    pvCheck pvCount(oCopy, "T") = 1, "CopyDatabase a memoria conserva los datos"
    pvCheck CStr(oCopy.GetRs("SELECT Big FROM T").Fields(0).Value) = BIG_INT64, "int64 intacto en la copia"
    baBlob = oCopy.GetRs("SELECT Bin FROM T").Fields(0).Value
    pvCheck pvSameBytes(baBlob, pvTestBlob()), "Blob intacto en la copia"
    Set oCopy = Nothing

    '--- encrypted -> encrypted file with another key
    sFile = pvTempFolder() & "sqlite3mc_copy.db"
    pvDeleteDb sFile
    Set oCopy = oCnn.CopyDatabase(sFile, sKey & "-copia")
    pvCheck pvCount(oCopy, "T") = 1, "CopyDatabase a fichero cifrado conserva los datos"
    Set oCopy = Nothing
    pvCheck Not pvFileIsPlainSqlite(sFile), "La copia en fichero esta cifrada"
    Set oCopy = New cConnection
    pvCheck oCopy.OpenDB(sFile, sKey & "-copia"), "La copia abre con su propia clave"
    Set oCopy = Nothing
    If chkKeepFiles.Value = vbUnchecked Then
        pvDeleteDb sFile
    End If
End Sub

'--- helpers --------------------------------------------------------------

Private Function pvOpen(sKey As String) As cConnection
    Set pvOpen = New cConnection
    If Not pvOpen.OpenDB(m_sDbFile, sKey) Then
        Err.Raise vbObjectError, , "No se puede abrir " & m_sDbFile & ": " & pvOpen.OpenErrorMessage
    End If
End Function

Private Function pvCount(oCnn As cConnection, sTable As String) As Long
    pvCount = oCnn.GetRs("SELECT COUNT(*) FROM " & sTable).Fields(0).Value
End Function

Private Function pvUnicodeText() As String
    '--- Spanish, German, Greek, Cyrillic, euro sign and a CJK character
    pvUnicodeText = "A" & ChrW$(&HF1) & "o " & ChrW$(&HF1) & "and" & ChrW$(&HFA) & " Gr" & ChrW$(&HF6) & ChrW$(&HDF) & "e " & ChrW$(&H3A9) & ChrW$(&H3B1) & " " & ChrW$(&H416) & " " & ChrW$(&H20AC) & " " & ChrW$(&H6F22)
End Function

Private Function pvTestBlob() As Byte()
    Dim baBuf(0 To 255) As Byte
    Dim lIdx            As Long

    For lIdx = 0 To 255
        baBuf(lIdx) = lIdx
    Next
    pvTestBlob = baBuf
End Function

Private Function pvTestDate() As Date
    pvTestDate = DateSerial(2024, 2, 29) + TimeSerial(23, 59, 58)
End Function

Private Function pvSameBytes(baA() As Byte, baB() As Byte) As Boolean
    Dim lIdx            As Long

    On Error GoTo QH
    If UBound(baA) <> UBound(baB) Then
        Exit Function
    End If
    For lIdx = 0 To UBound(baA)
        If baA(lIdx) <> baB(lIdx) Then
            Exit Function
        End If
    Next
    pvSameBytes = True
QH:
End Function

Private Function pvFileIsPlainSqlite(sFile As String) As Boolean
    Dim nFile           As Integer
    Dim baHeader(0 To 14) As Byte

    nFile = FreeFile
    Open sFile For Binary Access Read As #nFile
    Get #nFile, 1, baHeader
    Close #nFile
    pvFileIsPlainSqlite = (StrConv(baHeader, vbUnicode) = "SQLite format 3")
End Function

Private Function pvTempFolder() As String
    pvTempFolder = Environ$("TEMP")
    If Len(pvTempFolder) = 0 Then
        pvTempFolder = CurDir$
    End If
    If Right$(pvTempFolder, 1) <> "\" Then
        pvTempFolder = pvTempFolder & "\"
    End If
End Function

Private Sub pvDeleteDb(sFile As String)
    On Error Resume Next
    Kill sFile
    Kill sFile & "-journal"
    Kill sFile & "-wal"
    Kill sFile & "-shm"
End Sub

Private Sub pvCheck(ByVal bCondition As Boolean, sText As String)
    If bCondition Then
        m_lPassed = m_lPassed + 1
        pvLog "  [OK]    " & sText
    Else
        m_lFailed = m_lFailed + 1
        pvLog "  [FALLO] " & sText
    End If
End Sub

Private Sub pvSkip(sText As String)
    m_lSkipped = m_lSkipped + 1
    pvLog "  [SKIP]  " & sText
End Sub

Private Sub pvLog(sText As String)
    txtLog.SelStart = Len(txtLog.Text)
    txtLog.SelText = sText & vbCrLf
    DoEvents
End Sub
