# SQLite para twinBASIC (32 y 64 bits) con SQLite3 Multiple Ciphers

Envoltorio de SQLite, compatible con la API de RC6 (`cConnection`, `cRecordset`, `cCommand`...),
para **twinBASIC**. Compila igual en **Win32** y en **Win64** y usa como motor
[SQLite3 Multiple Ciphers](https://github.com/utelle/SQLite3MultipleCiphers)
distribuido como **`sqlite3.dll`**.

## Instalación

1. Descarga la versión de Windows desde
   [Releases de SQLite3MultipleCiphers](https://github.com/utelle/SQLite3MultipleCiphers/releases).
2. Las dos DLL deben llamarse **`sqlite3.dll`**, así que ponlas en subcarpetas `x86` y
   `x64` junto al proyecto (`.twinproj`) y junto al ejecutable:

   ```
   MiProyecto\
     MiProyecto.twinproj
     x86\sqlite3.dll      <- sqlite3mc_x86.dll (o sqlite3mc.dll) renombrada
     x64\sqlite3.dll      <- sqlite3mc_x64.dll renombrada
     Build\MiProyecto_win32.exe
     Build\x86\sqlite3.dll
     Build\MiProyecto_win64.exe
     Build\x64\sqlite3.dll
   ```

   `mdSqlite3Loader.bas` carga la DLL **por ruta completa** antes de la primera llamada.
   Busca primero en `App.Path\x86` o `App.Path\x64` (según la arquitectura del proceso),
   después en `App.Path` y por último en el directorio actual. Esto es necesario en el IDE:
   el proceso es el de twinBASIC, y Windows no busca DLL en la carpeta del proyecto.
   También se puede indicar la ruta: `Sqlite3EnsureLoaded "C:\ruta\sqlite3.dll"`.

3. En twinBASIC, importa todos los `.bas` y `.cls` del repositorio en el proyecto
   (*Project > Import file...*).

### Si no carga la DLL

`cConnection.OpenDB`/`CreateNewDB` devuelven `False` con `OpenErrorCode = -1`, y
`OpenErrorMessage` (o `Sqlite3LoadError`) explica la causa:

| Mensaje | Causa | Solución |
|---------|-------|----------|
| `No se encuentra sqlite3.dll` | No está en ninguna de las rutas probadas (se listan) | Copiarla a `x86\` / `x64\` junto al proyecto o al exe |
| `error 193` | DLL de la otra arquitectura (p.ej. la x64 en un proceso de 32 bits) | Usar la DLL que corresponda a la arquitectura del proceso |
| `error 126` | Falta una DLL de la que depende | Instalar *Microsoft Visual C++ Redistributable* (x86 y/o x64) |
| `error 5` | Fichero bloqueado (descargado de Internet) | *Propiedades > Desbloquear* |
| `es SQLite sin cifrado` | Es una `sqlite3.dll` normal, no la de SQLite3MC | Usar la DLL de SQLite3 Multiple Ciphers |

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

## Formulario de prueba (`Test/`)

`Test/frmTest.frm` es un banco de pruebas que ejercita el wrapper completo contra la
`sqlite3.dll` de SQLite3MC. Las clases `Test/cTestFunc.cls`, `Test/cTestAgg.cls` y
`Test/cTestColl.cls` implementan `IFunction`, `IAggregateFunction` e `ICollation` para
probar los callbacks `CDecl`.

1. Importa en el proyecto los módulos del wrapper (incluido `mdSqlite3Loader.bas`) y los
   cuatro ficheros de `Test/`.
2. Establece `frmTest` como formulario de inicio.
3. Compila **una vez como Win32 y otra como Win64**, cada una con su `sqlite3.dll`,
   y pulsa **Ejecutar pruebas**. Todas las líneas deben salir `[OK]`. Sale `[SKIP]`
   cuando un cifrado no está compilado en la DLL. La primera línea indica qué `sqlite3.dll`
   se ha cargado; si no se puede cargar, explica por qué.

Las pruebas comprueban:

| # | Prueba |
|---|--------|
| 1 | Plataforma, tamaño de `LongPtr`, versión de SQLite y de SQLite3MC, cifrados disponibles |
| 2 | Crear una BD cifrada; texto Unicode, int64 > 2^53, double, blob, fecha y NULL; `LastInsertAutoID` es `LongLong`; el fichero no tiene la cabecera `SQLite format 3` |
| 3 | Reabrir con la clave y verificar todos los valores e `integrity_check` |
| 4 | Clave incorrecta o sin clave: `OpenDB = False`, `OpenFailedBadKey`, `SQLITE_NOTADB` |
| 5 | `BeginTrans` / `RollbackTrans` / `CommitTrans` con savepoints anidados |
| 6 | UDF escalares, agregado y colación VB (callbacks `CDecl`), un error VB que llega como error SQL, y la eliminación de la UDF |
| 7 | `cCommand` (1000 inserciones con `SetInt64`) y `cSelectCommand` |
| 8 | `ReKey`: cambiar la clave, quitar el cifrado (`ReKey ""`) y volver a cifrar |
| 9 | Crear, reabrir y leer con cada cifrado: AES128, AES256, ChaCha20, SQLCipher, RC4, Ascon128 y AEGIS |
| 10 | `CopyDatabase` a `:memory:` (API de backup) |

Los ficheros de prueba se crean en `%TEMP%` y se borran al terminar, salvo que se marque
*Conservar ficheros*.
