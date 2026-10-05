# SQLite para twinBASIC (32 y 64 bits) con SQLite3 Multiple Ciphers

Envoltorio de SQLite, compatible con la API de RC6 (`cConnection`, `cRecordset`, `cCommand`...),
para **twinBASIC**. Compila igual en **Win32** y en **Win64** y usa como motor
[SQLite3 Multiple Ciphers](https://github.com/utelle/SQLite3MultipleCiphers)
distribuido como **`sqlite3.dll`**.

## Instalación

1. Descarga la versión de Windows desde
   [Releases de SQLite3MultipleCiphers](https://github.com/utelle/SQLite3MultipleCiphers/releases).
2. Copia la DLL que corresponda a la arquitectura del ejecutable, renómbrala a
   `sqlite3.dll` y déjala junto al `.exe` (o en otra carpeta del `PATH`):

   | Compilación twinBASIC | DLL del paquete                       | Renombrar a   |
   |-----------------------|---------------------------------------|---------------|
   | Win32                 | `sqlite3mc_x86.dll` (o `sqlite3mc.dll`) | `sqlite3.dll` |
   | Win64                 | `sqlite3mc_x64.dll`                   | `sqlite3.dll` |

3. En twinBASIC, importa todos los `.bas` y `.cls` del repositorio en el proyecto
   (*Project > Import file...*).

## Qué cambió respecto a la versión VB6

- **Convención de llamada `CDecl`.** La `sqlite3.dll` oficial (y la de SQLite3MC) usa
  `__cdecl`, no `StdCall` como `winsqlite3.dll`. En 32 bits esto es crítico: todos los
  `Declare` de `sqlite3win32stubs.bas` llevan `CDecl`, y los callbacks de funciones,
  agregados y colaciones definidos por el usuario (`mdUdf.bas`) también se declaran
  `CDecl`. En 64 bits la palabra clave no tiene efecto.
- **`PtrSafe` y `LongPtr`** en todos los `Declare` (también los de `kernel32`/`crypt32`), y
  los handles y punteros públicos (`DBHdl`, `StmtHdl`, `context`, `Set*Ptr`...) pasan a ser `LongPtr`.
- **`sqlite3_int64` es `LongLong` en 32 y 64 bits.** twinBASIC tiene `LongLong` también en
  Win32, así que se eliminaron los `#If Win64` y el truco de `Currency` × 10000. Los
  valores int64 (`LastInsertAutoID`, `GetInt64`, `UniqueID64`...) se devuelven como `Variant`
  de tipo `LongLong` (`VarType = 20`) en ambas plataformas.
- **API de cifrado de SQLite3MC**: `sqlite3_key[_v2]`, `sqlite3_rekey[_v2]`,
  `sqlite3mc_config`, `sqlite3mc_config_cipher`, `sqlite3mc_codec_data`,
  `sqlite3mc_cipher_count/index/name` y `sqlite3mc_version`. Se corrigieron las
  declaraciones de códec, que estaban dañadas (`stub_sqlite3x64.dll_config`).
- Gracias a `CDecl`, `sqlite3_db_config` y `sqlite3_config` (variádicas) se pueden llamar
  con aridad fija: `stub_sqlite3_db_config_int` y `stub_sqlite3_config_int`.

## Ejemplo

```vb
Dim Cnn As New cConnection
Cnn.CodecType = CODEC_TYPE_CHACHA20          ' opcional; 0 = cifrado por defecto (ChaCha20)
If Not Cnn.CreateNewDB(App.Path & "\datos.db", "MiClave") Then
    MsgBox "Error " & Cnn.OpenErrorCode & ": " & Cnn.OpenErrorMessage
    Exit Sub
End If
Debug.Print Cnn.Version, Cnn.CipherLibVersion

Cnn.Execute "CREATE TABLE IF NOT EXISTS T (ID INTEGER PRIMARY KEY, Nombre TEXT)"
Cnn.ExecCmd "INSERT INTO T (Nombre) VALUES (?)", "Hola"
Debug.Print Cnn.LastInsertAutoID

Dim Rs As cRecordset
Set Rs = Cnn.GetRs("SELECT * FROM T")
Do Until Rs.EOF
    Debug.Print Rs.Fields("ID").Value, Rs.Fields("Nombre").Value
    Rs.MoveNext
Loop

Cnn.ReKey "NuevaClave"   ' cambia la clave; Cnn.ReKey "" quita el cifrado
```

Si la clave o el cifrado no son correctos, `OpenDB` devuelve `False` y
`OpenFailedBadKey` es `True` (`SQLITE_NOTADB`, 26).
