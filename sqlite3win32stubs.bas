Attribute VB_Name = "sqlite3win32stubs"
'=========================================================================
' sqlite3win32stubs - stub_sqlite3_* declares for twinBASIC (Win32 + Win64)
' against SQLite3 Multiple Ciphers (github.com/utelle/SQLite3MultipleCiphers)
' deployed as sqlite3.dll.
'
' - Deploy the build that matches the bitness of your executable:
'     Win32 build -> sqlite3mc_x86.dll (or sqlite3mc.dll) renamed to sqlite3.dll
'     Win64 build -> sqlite3mc_x64.dll renamed to sqlite3.dll
'   next to the .exe (or anywhere on the DLL search path).
' - The official SQLite / SQLite3MC Windows binaries use the C calling
'   convention (__cdecl), NOT StdCall like winsqlite3.dll. On Win32 this
'   matters (the caller cleans the stack), so every declare is marked
'   CDecl; on Win64 there is a single calling convention and the keyword
'   is ignored. Callbacks handed to SQLite must therefore be CDecl too
'   (see mdUdf).
' - Handles/pointers are LongPtr; int is Long; sqlite3_int64 is LongLong
'   (twinBASIC supports LongLong in 32-bit builds as well).
' - SQLite is UTF-8: pass/receive LongPtr to byte buffers, not VB String.
' - The core constants live in sqlite3win32helper.bas; for the full set of
'   result codes, flags and options refer to sqlite3.h / sqlite3mc.h.
'=========================================================================
Option Explicit

'--- sqlite3.dll exports (standard sqlite3_* API)
Public Declare PtrSafe Function stub_sqlite3_aggregate_context CDecl Lib "sqlite3.dll" Alias "sqlite3_aggregate_context" (ByVal p1 As LongPtr, ByVal nBytes As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_aggregate_count CDecl Lib "sqlite3.dll" Alias "sqlite3_aggregate_count" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_auto_extension CDecl Lib "sqlite3.dll" Alias "sqlite3_auto_extension" (ByVal xEntryPoint As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_autovacuum_pages CDecl Lib "sqlite3.dll" Alias "sqlite3_autovacuum_pages" (ByVal db As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_backup_finish CDecl Lib "sqlite3.dll" Alias "sqlite3_backup_finish" (ByVal p As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_backup_init CDecl Lib "sqlite3.dll" Alias "sqlite3_backup_init" (ByVal pDest As LongPtr, ByVal zDestName As LongPtr, ByVal pSource As LongPtr, ByVal zSourceName As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_backup_pagecount CDecl Lib "sqlite3.dll" Alias "sqlite3_backup_pagecount" (ByVal p As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_backup_remaining CDecl Lib "sqlite3.dll" Alias "sqlite3_backup_remaining" (ByVal p As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_backup_step CDecl Lib "sqlite3.dll" Alias "sqlite3_backup_step" (ByVal p As LongPtr, ByVal nPage As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_blob CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_blob" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal n As Long, ByVal cb5 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_blob64 CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_blob64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongLong, ByVal cb5 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_double CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_double" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Double) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_int CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_int" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_int64 CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_int64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongLong) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_null CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_null" (ByVal p1 As LongPtr, ByVal p2 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_parameter_count CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_parameter_count" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_parameter_index CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_parameter_index" (ByVal p1 As LongPtr, ByVal zName As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_parameter_name CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_parameter_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_bind_pointer CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_pointer" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongPtr, ByVal cb5 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_text CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_text" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Long, ByVal cb5 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_text16 CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_text16" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As Long, ByVal cb5 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_text64 CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_text64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongLong, ByVal cb5 As LongPtr, ByVal encoding As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_value CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_value" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_zeroblob CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_zeroblob" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal n As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_bind_zeroblob64 CDecl Lib "sqlite3.dll" Alias "sqlite3_bind_zeroblob64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongLong) As Long
Public Declare PtrSafe Function stub_sqlite3_blob_bytes CDecl Lib "sqlite3.dll" Alias "sqlite3_blob_bytes" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_blob_close CDecl Lib "sqlite3.dll" Alias "sqlite3_blob_close" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_blob_open CDecl Lib "sqlite3.dll" Alias "sqlite3_blob_open" (ByVal p1 As LongPtr, ByVal zDb As LongPtr, ByVal zTable As LongPtr, ByVal zColumn As LongPtr, ByVal iRow As LongLong, ByVal flags As Long, ByVal ppBlob As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_blob_read CDecl Lib "sqlite3.dll" Alias "sqlite3_blob_read" (ByVal p1 As LongPtr, ByVal Z As LongPtr, ByVal N As Long, ByVal iOffset As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_blob_reopen CDecl Lib "sqlite3.dll" Alias "sqlite3_blob_reopen" (ByVal p1 As LongPtr, ByVal p2 As LongLong) As Long
Public Declare PtrSafe Function stub_sqlite3_blob_write CDecl Lib "sqlite3.dll" Alias "sqlite3_blob_write" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As Long, ByVal iOffset As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_busy_handler CDecl Lib "sqlite3.dll" Alias "sqlite3_busy_handler" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_busy_timeout CDecl Lib "sqlite3.dll" Alias "sqlite3_busy_timeout" (ByVal p1 As LongPtr, ByVal ms As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_cancel_auto_extension CDecl Lib "sqlite3.dll" Alias "sqlite3_cancel_auto_extension" (ByVal xEntryPoint As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_changes CDecl Lib "sqlite3.dll" Alias "sqlite3_changes" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_changes64 CDecl Lib "sqlite3.dll" Alias "sqlite3_changes64" (ByVal p1 As LongPtr) As LongLong
Public Declare PtrSafe Function stub_sqlite3_clear_bindings CDecl Lib "sqlite3.dll" Alias "sqlite3_clear_bindings" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_close CDecl Lib "sqlite3.dll" Alias "sqlite3_close" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_close_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_close_v2" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_collation_needed CDecl Lib "sqlite3.dll" Alias "sqlite3_collation_needed" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal cb3 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_collation_needed16 CDecl Lib "sqlite3.dll" Alias "sqlite3_collation_needed16" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal cb3 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_column_blob CDecl Lib "sqlite3.dll" Alias "sqlite3_column_blob" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_bytes CDecl Lib "sqlite3.dll" Alias "sqlite3_column_bytes" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_column_bytes16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_bytes16" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_column_count CDecl Lib "sqlite3.dll" Alias "sqlite3_column_count" (ByVal pStmt As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_column_database_name CDecl Lib "sqlite3.dll" Alias "sqlite3_column_database_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_database_name16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_database_name16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_decltype CDecl Lib "sqlite3.dll" Alias "sqlite3_column_decltype" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_decltype16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_decltype16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_double CDecl Lib "sqlite3.dll" Alias "sqlite3_column_double" (ByVal p1 As LongPtr, ByVal iCol As Long) As Double
Public Declare PtrSafe Function stub_sqlite3_column_int CDecl Lib "sqlite3.dll" Alias "sqlite3_column_int" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_column_int64 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_int64" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongLong
Public Declare PtrSafe Function stub_sqlite3_column_name CDecl Lib "sqlite3.dll" Alias "sqlite3_column_name" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_name16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_name16" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_origin_name CDecl Lib "sqlite3.dll" Alias "sqlite3_column_origin_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_origin_name16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_origin_name16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_table_name CDecl Lib "sqlite3.dll" Alias "sqlite3_column_table_name" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_table_name16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_table_name16" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_text CDecl Lib "sqlite3.dll" Alias "sqlite3_column_text" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_text16 CDecl Lib "sqlite3.dll" Alias "sqlite3_column_text16" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_column_type CDecl Lib "sqlite3.dll" Alias "sqlite3_column_type" (ByVal p1 As LongPtr, ByVal iCol As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_column_value CDecl Lib "sqlite3.dll" Alias "sqlite3_column_value" (ByVal p1 As LongPtr, ByVal iCol As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_commit_hook CDecl Lib "sqlite3.dll" Alias "sqlite3_commit_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_compileoption_get CDecl Lib "sqlite3.dll" Alias "sqlite3_compileoption_get" (ByVal N As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_compileoption_used CDecl Lib "sqlite3.dll" Alias "sqlite3_compileoption_used" (ByVal zOptName As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_complete CDecl Lib "sqlite3.dll" Alias "sqlite3_complete" (ByVal sql As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_complete16 CDecl Lib "sqlite3.dll" Alias "sqlite3_complete16" (ByVal sql As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_context_db_handle CDecl Lib "sqlite3.dll" Alias "sqlite3_context_db_handle" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_create_collation CDecl Lib "sqlite3.dll" Alias "sqlite3_create_collation" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_collation_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_create_collation_v2" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_collation16 CDecl Lib "sqlite3.dll" Alias "sqlite3_create_collation16" (ByVal p1 As LongPtr, ByVal zName As LongPtr, ByVal eTextRep As Long, ByVal pArg As LongPtr, ByVal xCompare As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_filename CDecl Lib "sqlite3.dll" Alias "sqlite3_create_filename" (ByVal zDatabase As LongPtr, ByVal zJournal As LongPtr, ByVal zWal As LongPtr, ByVal nParam As Long, ByVal azParam As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_create_function CDecl Lib "sqlite3.dll" Alias "sqlite3_create_function" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_function_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_create_function_v2" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_function16 CDecl Lib "sqlite3.dll" Alias "sqlite3_create_function16" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xFunc As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_module CDecl Lib "sqlite3.dll" Alias "sqlite3_create_module" (ByVal db As LongPtr, ByVal zName As LongPtr, ByVal p As LongPtr, ByVal pClientData As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_module_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_create_module_v2" (ByVal db As LongPtr, ByVal zName As LongPtr, ByVal p As LongPtr, ByVal pClientData As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_create_window_function CDecl Lib "sqlite3.dll" Alias "sqlite3_create_window_function" (ByVal db As LongPtr, ByVal zFunctionName As LongPtr, ByVal nArg As Long, ByVal eTextRep As Long, ByVal pApp As LongPtr, ByVal xStep As LongPtr, ByVal xFinal As LongPtr, ByVal xValue As LongPtr, ByVal xInverse As LongPtr, ByVal xDestroy As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_data_count CDecl Lib "sqlite3.dll" Alias "sqlite3_data_count" (ByVal pStmt As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_database_file_object CDecl Lib "sqlite3.dll" Alias "sqlite3_database_file_object" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_db_cacheflush CDecl Lib "sqlite3.dll" Alias "sqlite3_db_cacheflush" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_db_filename CDecl Lib "sqlite3.dll" Alias "sqlite3_db_filename" (ByVal db As LongPtr, ByVal zDbName As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_db_handle CDecl Lib "sqlite3.dll" Alias "sqlite3_db_handle" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_db_mutex CDecl Lib "sqlite3.dll" Alias "sqlite3_db_mutex" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_db_name CDecl Lib "sqlite3.dll" Alias "sqlite3_db_name" (ByVal db As LongPtr, ByVal N As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_db_readonly CDecl Lib "sqlite3.dll" Alias "sqlite3_db_readonly" (ByVal db As LongPtr, ByVal zDbName As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_db_release_memory CDecl Lib "sqlite3.dll" Alias "sqlite3_db_release_memory" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_db_status CDecl Lib "sqlite3.dll" Alias "sqlite3_db_status" (ByVal p1 As LongPtr, ByVal op As Long, ByVal pCur As LongPtr, ByVal pHiwtr As LongPtr, ByVal resetFlg As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_db_status64 CDecl Lib "sqlite3.dll" Alias "sqlite3_db_status64" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal p3 As LongPtr, ByVal p4 As LongPtr, ByVal p5 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_declare_vtab CDecl Lib "sqlite3.dll" Alias "sqlite3_declare_vtab" (ByVal p1 As LongPtr, ByVal zSQL As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_deserialize CDecl Lib "sqlite3.dll" Alias "sqlite3_deserialize" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal pData As LongPtr, ByVal szDb As LongLong, ByVal szBuf As LongLong, ByVal mFlags As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_drop_modules CDecl Lib "sqlite3.dll" Alias "sqlite3_drop_modules" (ByVal db As LongPtr, ByVal azKeep As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_enable_load_extension CDecl Lib "sqlite3.dll" Alias "sqlite3_enable_load_extension" (ByVal db As LongPtr, ByVal onoff As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_enable_shared_cache CDecl Lib "sqlite3.dll" Alias "sqlite3_enable_shared_cache" (ByVal p1 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_errcode CDecl Lib "sqlite3.dll" Alias "sqlite3_errcode" (ByVal db As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_errmsg CDecl Lib "sqlite3.dll" Alias "sqlite3_errmsg" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_errmsg16 CDecl Lib "sqlite3.dll" Alias "sqlite3_errmsg16" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_error_offset CDecl Lib "sqlite3.dll" Alias "sqlite3_error_offset" (ByVal db As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_errstr CDecl Lib "sqlite3.dll" Alias "sqlite3_errstr" (ByVal p1 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_exec CDecl Lib "sqlite3.dll" Alias "sqlite3_exec" (ByVal p1 As LongPtr, ByVal sql As LongPtr, ByVal callback As LongPtr, ByVal p4 As LongPtr, ByVal errmsg As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_expanded_sql CDecl Lib "sqlite3.dll" Alias "sqlite3_expanded_sql" (ByVal pStmt As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_expired CDecl Lib "sqlite3.dll" Alias "sqlite3_expired" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_extended_errcode CDecl Lib "sqlite3.dll" Alias "sqlite3_extended_errcode" (ByVal db As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_extended_result_codes CDecl Lib "sqlite3.dll" Alias "sqlite3_extended_result_codes" (ByVal p1 As LongPtr, ByVal onoff As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_file_control CDecl Lib "sqlite3.dll" Alias "sqlite3_file_control" (ByVal p1 As LongPtr, ByVal zDbName As LongPtr, ByVal op As Long, ByVal p4 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_filename_database CDecl Lib "sqlite3.dll" Alias "sqlite3_filename_database" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_filename_journal CDecl Lib "sqlite3.dll" Alias "sqlite3_filename_journal" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_filename_wal CDecl Lib "sqlite3.dll" Alias "sqlite3_filename_wal" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_finalize CDecl Lib "sqlite3.dll" Alias "sqlite3_finalize" (ByVal pStmt As LongPtr) As Long
Public Declare PtrSafe Sub stub_sqlite3_free CDecl Lib "sqlite3.dll" Alias "sqlite3_free" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_free_filename CDecl Lib "sqlite3.dll" Alias "sqlite3_free_filename" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_free_table CDecl Lib "sqlite3.dll" Alias "sqlite3_free_table" (ByVal result As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_get_autocommit CDecl Lib "sqlite3.dll" Alias "sqlite3_get_autocommit" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_get_auxdata CDecl Lib "sqlite3.dll" Alias "sqlite3_get_auxdata" (ByVal p1 As LongPtr, ByVal N As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_get_clientdata CDecl Lib "sqlite3.dll" Alias "sqlite3_get_clientdata" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_get_table CDecl Lib "sqlite3.dll" Alias "sqlite3_get_table" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal pazResult As LongPtr, ByVal pnRow As LongPtr, ByVal pnColumn As LongPtr, ByVal pzErrmsg As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_global_recover CDecl Lib "sqlite3.dll" Alias "sqlite3_global_recover" () As Long
Public Declare PtrSafe Function stub_sqlite3_hard_heap_limit64 CDecl Lib "sqlite3.dll" Alias "sqlite3_hard_heap_limit64" (ByVal N As LongLong) As LongLong
Public Declare PtrSafe Function stub_sqlite3_initialize CDecl Lib "sqlite3.dll" Alias "sqlite3_initialize" () As Long
Public Declare PtrSafe Sub stub_sqlite3_interrupt CDecl Lib "sqlite3.dll" Alias "sqlite3_interrupt" (ByVal p1 As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_is_interrupted CDecl Lib "sqlite3.dll" Alias "sqlite3_is_interrupted" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_keyword_check CDecl Lib "sqlite3.dll" Alias "sqlite3_keyword_check" (ByVal p1 As LongPtr, ByVal p2 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_keyword_count CDecl Lib "sqlite3.dll" Alias "sqlite3_keyword_count" () As Long
Public Declare PtrSafe Function stub_sqlite3_keyword_name CDecl Lib "sqlite3.dll" Alias "sqlite3_keyword_name" (ByVal p1 As Long, ByVal p2 As LongPtr, ByVal p3 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_last_insert_rowid CDecl Lib "sqlite3.dll" Alias "sqlite3_last_insert_rowid" (ByVal p1 As LongPtr) As LongLong
Public Declare PtrSafe Function stub_sqlite3_libversion CDecl Lib "sqlite3.dll" Alias "sqlite3_libversion" () As LongPtr
Public Declare PtrSafe Function stub_sqlite3_libversion_number CDecl Lib "sqlite3.dll" Alias "sqlite3_libversion_number" () As Long
Public Declare PtrSafe Function stub_sqlite3_limit CDecl Lib "sqlite3.dll" Alias "sqlite3_limit" (ByVal p1 As LongPtr, ByVal id As Long, ByVal newVal As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_load_extension CDecl Lib "sqlite3.dll" Alias "sqlite3_load_extension" (ByVal db As LongPtr, ByVal zFile As LongPtr, ByVal zProc As LongPtr, ByVal pzErrMsg As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_malloc CDecl Lib "sqlite3.dll" Alias "sqlite3_malloc" (ByVal p1 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_malloc64 CDecl Lib "sqlite3.dll" Alias "sqlite3_malloc64" (ByVal p1 As LongLong) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_memory_alarm CDecl Lib "sqlite3.dll" Alias "sqlite3_memory_alarm" (ByVal cb1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongLong) As Long
Public Declare PtrSafe Function stub_sqlite3_memory_highwater CDecl Lib "sqlite3.dll" Alias "sqlite3_memory_highwater" (ByVal resetFlag As Long) As LongLong
Public Declare PtrSafe Function stub_sqlite3_memory_used CDecl Lib "sqlite3.dll" Alias "sqlite3_memory_used" () As LongLong
Public Declare PtrSafe Function stub_sqlite3_msize CDecl Lib "sqlite3.dll" Alias "sqlite3_msize" (ByVal p1 As LongPtr) As LongLong
Public Declare PtrSafe Function stub_sqlite3_mutex_alloc CDecl Lib "sqlite3.dll" Alias "sqlite3_mutex_alloc" (ByVal p1 As Long) As LongPtr
Public Declare PtrSafe Sub stub_sqlite3_mutex_enter CDecl Lib "sqlite3.dll" Alias "sqlite3_mutex_enter" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_mutex_free CDecl Lib "sqlite3.dll" Alias "sqlite3_mutex_free" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_mutex_leave CDecl Lib "sqlite3.dll" Alias "sqlite3_mutex_leave" (ByVal p1 As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_mutex_try CDecl Lib "sqlite3.dll" Alias "sqlite3_mutex_try" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_next_stmt CDecl Lib "sqlite3.dll" Alias "sqlite3_next_stmt" (ByVal pDb As LongPtr, ByVal pStmt As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_open CDecl Lib "sqlite3.dll" Alias "sqlite3_open" (ByVal filename As LongPtr, ByVal ppDb As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_open_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_open_v2" (ByVal filename As LongPtr, ByVal ppDb As LongPtr, ByVal flags As Long, ByVal zVfs As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_open16 CDecl Lib "sqlite3.dll" Alias "sqlite3_open16" (ByVal filename As LongPtr, ByVal ppDb As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_os_end CDecl Lib "sqlite3.dll" Alias "sqlite3_os_end" () As Long
Public Declare PtrSafe Function stub_sqlite3_os_init CDecl Lib "sqlite3.dll" Alias "sqlite3_os_init" () As Long
Public Declare PtrSafe Function stub_sqlite3_overload_function CDecl Lib "sqlite3.dll" Alias "sqlite3_overload_function" (ByVal p1 As LongPtr, ByVal zFuncName As LongPtr, ByVal nArg As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_prepare CDecl Lib "sqlite3.dll" Alias "sqlite3_prepare" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_prepare_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_prepare_v2" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_prepare_v3 CDecl Lib "sqlite3.dll" Alias "sqlite3_prepare_v3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal prepFlags As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_prepare16 CDecl Lib "sqlite3.dll" Alias "sqlite3_prepare16" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_prepare16_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_prepare16_v2" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_prepare16_v3 CDecl Lib "sqlite3.dll" Alias "sqlite3_prepare16_v3" (ByVal db As LongPtr, ByVal zSql As LongPtr, ByVal nByte As Long, ByVal prepFlags As Long, ByVal ppStmt As LongPtr, ByVal pzTail As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_profile CDecl Lib "sqlite3.dll" Alias "sqlite3_profile" (ByVal p1 As LongPtr, ByVal xProfile As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare PtrSafe Sub stub_sqlite3_progress_handler CDecl Lib "sqlite3.dll" Alias "sqlite3_progress_handler" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal cb3 As LongPtr, ByVal p4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_randomness CDecl Lib "sqlite3.dll" Alias "sqlite3_randomness" (ByVal N As Long, ByVal P As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_realloc CDecl Lib "sqlite3.dll" Alias "sqlite3_realloc" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_realloc64 CDecl Lib "sqlite3.dll" Alias "sqlite3_realloc64" (ByVal p1 As LongPtr, ByVal p2 As LongLong) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_release_memory CDecl Lib "sqlite3.dll" Alias "sqlite3_release_memory" (ByVal p1 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_reset CDecl Lib "sqlite3.dll" Alias "sqlite3_reset" (ByVal pStmt As LongPtr) As Long
Public Declare PtrSafe Sub stub_sqlite3_reset_auto_extension CDecl Lib "sqlite3.dll" Alias "sqlite3_reset_auto_extension" ()
Public Declare PtrSafe Sub stub_sqlite3_result_blob CDecl Lib "sqlite3.dll" Alias "sqlite3_result_blob" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_blob64 CDecl Lib "sqlite3.dll" Alias "sqlite3_result_blob64" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongLong, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_double CDecl Lib "sqlite3.dll" Alias "sqlite3_result_double" (ByVal p1 As LongPtr, ByVal p2 As Double)
Public Declare PtrSafe Sub stub_sqlite3_result_error CDecl Lib "sqlite3.dll" Alias "sqlite3_result_error" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long)
Public Declare PtrSafe Sub stub_sqlite3_result_error_code CDecl Lib "sqlite3.dll" Alias "sqlite3_result_error_code" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare PtrSafe Sub stub_sqlite3_result_error_nomem CDecl Lib "sqlite3.dll" Alias "sqlite3_result_error_nomem" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_error_toobig CDecl Lib "sqlite3.dll" Alias "sqlite3_result_error_toobig" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_error16 CDecl Lib "sqlite3.dll" Alias "sqlite3_result_error16" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long)
Public Declare PtrSafe Sub stub_sqlite3_result_int CDecl Lib "sqlite3.dll" Alias "sqlite3_result_int" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare PtrSafe Sub stub_sqlite3_result_int64 CDecl Lib "sqlite3.dll" Alias "sqlite3_result_int64" (ByVal p1 As LongPtr, ByVal p2 As LongLong)
Public Declare PtrSafe Sub stub_sqlite3_result_null CDecl Lib "sqlite3.dll" Alias "sqlite3_result_null" (ByVal p1 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_pointer CDecl Lib "sqlite3.dll" Alias "sqlite3_result_pointer" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_subtype CDecl Lib "sqlite3.dll" Alias "sqlite3_result_subtype" (ByVal p1 As LongPtr, ByVal p2 As Long)
Public Declare PtrSafe Sub stub_sqlite3_result_text CDecl Lib "sqlite3.dll" Alias "sqlite3_result_text" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_text16 CDecl Lib "sqlite3.dll" Alias "sqlite3_result_text16" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_text16be CDecl Lib "sqlite3.dll" Alias "sqlite3_result_text16be" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_text16le CDecl Lib "sqlite3.dll" Alias "sqlite3_result_text16le" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long, ByVal cb4 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_text64 CDecl Lib "sqlite3.dll" Alias "sqlite3_result_text64" (ByVal p1 As LongPtr, ByVal z As LongPtr, ByVal n As LongLong, ByVal cb4 As LongPtr, ByVal encoding As Long)
Public Declare PtrSafe Sub stub_sqlite3_result_value CDecl Lib "sqlite3.dll" Alias "sqlite3_result_value" (ByVal p1 As LongPtr, ByVal p2 As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_result_zeroblob CDecl Lib "sqlite3.dll" Alias "sqlite3_result_zeroblob" (ByVal p1 As LongPtr, ByVal n As Long)
Public Declare PtrSafe Function stub_sqlite3_result_zeroblob64 CDecl Lib "sqlite3.dll" Alias "sqlite3_result_zeroblob64" (ByVal p1 As LongPtr, ByVal n As LongLong) As Long
Public Declare PtrSafe Function stub_sqlite3_rollback_hook CDecl Lib "sqlite3.dll" Alias "sqlite3_rollback_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_rtree_geometry_callback CDecl Lib "sqlite3.dll" Alias "sqlite3_rtree_geometry_callback" (ByVal db As LongPtr, ByVal zGeom As LongPtr, ByVal xGeom As LongPtr, ByVal pContext As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_rtree_query_callback CDecl Lib "sqlite3.dll" Alias "sqlite3_rtree_query_callback" (ByVal db As LongPtr, ByVal zQueryFunc As LongPtr, ByVal xQueryFunc As LongPtr, ByVal pContext As LongPtr, ByVal xDestructor As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_serialize CDecl Lib "sqlite3.dll" Alias "sqlite3_serialize" (ByVal db As LongPtr, ByVal zSchema As LongPtr, ByVal piSize As LongPtr, ByVal mFlags As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_set_authorizer CDecl Lib "sqlite3.dll" Alias "sqlite3_set_authorizer" (ByVal p1 As LongPtr, ByVal xAuth As LongPtr, ByVal pUserData As LongPtr) As Long
Public Declare PtrSafe Sub stub_sqlite3_set_auxdata CDecl Lib "sqlite3.dll" Alias "sqlite3_set_auxdata" (ByVal p1 As LongPtr, ByVal N As Long, ByVal p3 As LongPtr, ByVal cb4 As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_set_clientdata CDecl Lib "sqlite3.dll" Alias "sqlite3_set_clientdata" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal cb4 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_set_errmsg CDecl Lib "sqlite3.dll" Alias "sqlite3_set_errmsg" (ByVal db As LongPtr, ByVal errcode As Long, ByVal zErrMsg As LongPtr) As Long
Public Declare PtrSafe Sub stub_sqlite3_set_last_insert_rowid CDecl Lib "sqlite3.dll" Alias "sqlite3_set_last_insert_rowid" (ByVal p1 As LongPtr, ByVal p2 As LongLong)
Public Declare PtrSafe Function stub_sqlite3_setlk_timeout CDecl Lib "sqlite3.dll" Alias "sqlite3_setlk_timeout" (ByVal p1 As LongPtr, ByVal ms As Long, ByVal flags As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_shutdown CDecl Lib "sqlite3.dll" Alias "sqlite3_shutdown" () As Long
Public Declare PtrSafe Function stub_sqlite3_sleep CDecl Lib "sqlite3.dll" Alias "sqlite3_sleep" (ByVal p1 As Long) As Long
Public Declare PtrSafe Sub stub_sqlite3_soft_heap_limit CDecl Lib "sqlite3.dll" Alias "sqlite3_soft_heap_limit" (ByVal N As Long)
Public Declare PtrSafe Function stub_sqlite3_soft_heap_limit64 CDecl Lib "sqlite3.dll" Alias "sqlite3_soft_heap_limit64" (ByVal N As LongLong) As LongLong
Public Declare PtrSafe Function stub_sqlite3_sourceid CDecl Lib "sqlite3.dll" Alias "sqlite3_sourceid" () As LongPtr
Public Declare PtrSafe Function stub_sqlite3_sql CDecl Lib "sqlite3.dll" Alias "sqlite3_sql" (ByVal pStmt As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_status CDecl Lib "sqlite3.dll" Alias "sqlite3_status" (ByVal op As Long, ByVal pCurrent As LongPtr, ByVal pHighwater As LongPtr, ByVal resetFlag As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_status64 CDecl Lib "sqlite3.dll" Alias "sqlite3_status64" (ByVal op As Long, ByVal pCurrent As LongPtr, ByVal pHighwater As LongPtr, ByVal resetFlag As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_step CDecl Lib "sqlite3.dll" Alias "sqlite3_step" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_stmt_busy CDecl Lib "sqlite3.dll" Alias "sqlite3_stmt_busy" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_stmt_explain CDecl Lib "sqlite3.dll" Alias "sqlite3_stmt_explain" (ByVal pStmt As LongPtr, ByVal eMode As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_stmt_isexplain CDecl Lib "sqlite3.dll" Alias "sqlite3_stmt_isexplain" (ByVal pStmt As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_stmt_readonly CDecl Lib "sqlite3.dll" Alias "sqlite3_stmt_readonly" (ByVal pStmt As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_stmt_status CDecl Lib "sqlite3.dll" Alias "sqlite3_stmt_status" (ByVal p1 As LongPtr, ByVal op As Long, ByVal resetFlg As Long) As Long
Public Declare PtrSafe Sub stub_sqlite3_str_append CDecl Lib "sqlite3.dll" Alias "sqlite3_str_append" (ByVal p1 As LongPtr, ByVal zIn As LongPtr, ByVal N As Long)
Public Declare PtrSafe Sub stub_sqlite3_str_appendall CDecl Lib "sqlite3.dll" Alias "sqlite3_str_appendall" (ByVal p1 As LongPtr, ByVal zIn As LongPtr)
Public Declare PtrSafe Sub stub_sqlite3_str_appendchar CDecl Lib "sqlite3.dll" Alias "sqlite3_str_appendchar" (ByVal p1 As LongPtr, ByVal N As Long, ByVal C As Long)
Public Declare PtrSafe Function stub_sqlite3_str_errcode CDecl Lib "sqlite3.dll" Alias "sqlite3_str_errcode" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_str_finish CDecl Lib "sqlite3.dll" Alias "sqlite3_str_finish" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_str_length CDecl Lib "sqlite3.dll" Alias "sqlite3_str_length" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_str_new CDecl Lib "sqlite3.dll" Alias "sqlite3_str_new" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Sub stub_sqlite3_str_reset CDecl Lib "sqlite3.dll" Alias "sqlite3_str_reset" (ByVal p1 As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_str_value CDecl Lib "sqlite3.dll" Alias "sqlite3_str_value" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Sub stub_sqlite3_str_vappendf CDecl Lib "sqlite3.dll" Alias "sqlite3_str_vappendf" (ByVal p1 As LongPtr, ByVal zFormat As LongPtr, ByVal va_list As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_strglob CDecl Lib "sqlite3.dll" Alias "sqlite3_strglob" (ByVal zGlob As LongPtr, ByVal zStr As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_stricmp CDecl Lib "sqlite3.dll" Alias "sqlite3_stricmp" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_strlike CDecl Lib "sqlite3.dll" Alias "sqlite3_strlike" (ByVal zGlob As LongPtr, ByVal zStr As LongPtr, ByVal cEsc As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_strnicmp CDecl Lib "sqlite3.dll" Alias "sqlite3_strnicmp" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_system_errno CDecl Lib "sqlite3.dll" Alias "sqlite3_system_errno" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_table_column_metadata CDecl Lib "sqlite3.dll" Alias "sqlite3_table_column_metadata" (ByVal db As LongPtr, ByVal zDbName As LongPtr, ByVal zTableName As LongPtr, ByVal zColumnName As LongPtr, ByVal pzDataType As LongPtr, ByVal pzCollSeq As LongPtr, ByVal pNotNull As LongPtr, ByVal pPrimaryKey As LongPtr, ByVal pAutoinc As LongPtr) As Long
Public Declare PtrSafe Sub stub_sqlite3_thread_cleanup CDecl Lib "sqlite3.dll" Alias "sqlite3_thread_cleanup" ()
Public Declare PtrSafe Function stub_sqlite3_threadsafe CDecl Lib "sqlite3.dll" Alias "sqlite3_threadsafe" () As Long
Public Declare PtrSafe Function stub_sqlite3_total_changes CDecl Lib "sqlite3.dll" Alias "sqlite3_total_changes" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_total_changes64 CDecl Lib "sqlite3.dll" Alias "sqlite3_total_changes64" (ByVal p1 As LongPtr) As LongLong
Public Declare PtrSafe Function stub_sqlite3_trace CDecl Lib "sqlite3.dll" Alias "sqlite3_trace" (ByVal p1 As LongPtr, ByVal xTrace As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_trace_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_trace_v2" (ByVal p1 As LongPtr, ByVal uMask As Long, ByVal xCallback As LongPtr, ByVal pCtx As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_transfer_bindings CDecl Lib "sqlite3.dll" Alias "sqlite3_transfer_bindings" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_txn_state CDecl Lib "sqlite3.dll" Alias "sqlite3_txn_state" (ByVal p1 As LongPtr, ByVal zSchema As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_update_hook CDecl Lib "sqlite3.dll" Alias "sqlite3_update_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_uri_boolean CDecl Lib "sqlite3.dll" Alias "sqlite3_uri_boolean" (ByVal z As LongPtr, ByVal zParam As LongPtr, ByVal bDefault As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_uri_int64 CDecl Lib "sqlite3.dll" Alias "sqlite3_uri_int64" (ByVal p1 As LongPtr, ByVal p2 As LongPtr, ByVal p3 As LongLong) As LongLong
Public Declare PtrSafe Function stub_sqlite3_uri_key CDecl Lib "sqlite3.dll" Alias "sqlite3_uri_key" (ByVal z As LongPtr, ByVal N As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_uri_parameter CDecl Lib "sqlite3.dll" Alias "sqlite3_uri_parameter" (ByVal z As LongPtr, ByVal zParam As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_user_data CDecl Lib "sqlite3.dll" Alias "sqlite3_user_data" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_blob CDecl Lib "sqlite3.dll" Alias "sqlite3_value_blob" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_bytes CDecl Lib "sqlite3.dll" Alias "sqlite3_value_bytes" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_bytes16 CDecl Lib "sqlite3.dll" Alias "sqlite3_value_bytes16" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_double CDecl Lib "sqlite3.dll" Alias "sqlite3_value_double" (ByVal p1 As LongPtr) As Double
Public Declare PtrSafe Function stub_sqlite3_value_dup CDecl Lib "sqlite3.dll" Alias "sqlite3_value_dup" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_encoding CDecl Lib "sqlite3.dll" Alias "sqlite3_value_encoding" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Sub stub_sqlite3_value_free CDecl Lib "sqlite3.dll" Alias "sqlite3_value_free" (ByVal p1 As LongPtr)
Public Declare PtrSafe Function stub_sqlite3_value_frombind CDecl Lib "sqlite3.dll" Alias "sqlite3_value_frombind" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_int CDecl Lib "sqlite3.dll" Alias "sqlite3_value_int" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_int64 CDecl Lib "sqlite3.dll" Alias "sqlite3_value_int64" (ByVal p1 As LongPtr) As LongLong
Public Declare PtrSafe Function stub_sqlite3_value_nochange CDecl Lib "sqlite3.dll" Alias "sqlite3_value_nochange" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_numeric_type CDecl Lib "sqlite3.dll" Alias "sqlite3_value_numeric_type" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_pointer CDecl Lib "sqlite3.dll" Alias "sqlite3_value_pointer" (ByVal p1 As LongPtr, ByVal p2 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_subtype CDecl Lib "sqlite3.dll" Alias "sqlite3_value_subtype" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_value_text CDecl Lib "sqlite3.dll" Alias "sqlite3_value_text" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_text16 CDecl Lib "sqlite3.dll" Alias "sqlite3_value_text16" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_text16be CDecl Lib "sqlite3.dll" Alias "sqlite3_value_text16be" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_text16le CDecl Lib "sqlite3.dll" Alias "sqlite3_value_text16le" (ByVal p1 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_value_type CDecl Lib "sqlite3.dll" Alias "sqlite3_value_type" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vfs_find CDecl Lib "sqlite3.dll" Alias "sqlite3_vfs_find" (ByVal zVfsName As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_vfs_register CDecl Lib "sqlite3.dll" Alias "sqlite3_vfs_register" (ByVal p1 As LongPtr, ByVal makeDflt As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_vfs_unregister CDecl Lib "sqlite3.dll" Alias "sqlite3_vfs_unregister" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vmprintf CDecl Lib "sqlite3.dll" Alias "sqlite3_vmprintf" (ByVal p1 As LongPtr, ByVal va_list As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_vsnprintf CDecl Lib "sqlite3.dll" Alias "sqlite3_vsnprintf" (ByVal p1 As Long, ByVal p2 As LongPtr, ByVal p3 As LongPtr, ByVal va_list As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_vtab_collation CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_collation" (ByVal p1 As LongPtr, ByVal p2 As Long) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_vtab_distinct CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_distinct" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vtab_in CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_in" (ByVal p1 As LongPtr, ByVal iCons As Long, ByVal bHandle As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_vtab_in_first CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_in_first" (ByVal pVal As LongPtr, ByVal ppOut As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vtab_in_next CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_in_next" (ByVal pVal As LongPtr, ByVal ppOut As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vtab_nochange CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_nochange" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vtab_on_conflict CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_on_conflict" (ByVal p1 As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_vtab_rhs_value CDecl Lib "sqlite3.dll" Alias "sqlite3_vtab_rhs_value" (ByVal p1 As LongPtr, ByVal p2 As Long, ByVal ppVal As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_wal_autocheckpoint CDecl Lib "sqlite3.dll" Alias "sqlite3_wal_autocheckpoint" (ByVal db As LongPtr, ByVal N As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_wal_checkpoint CDecl Lib "sqlite3.dll" Alias "sqlite3_wal_checkpoint" (ByVal db As LongPtr, ByVal zDb As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_wal_checkpoint_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_wal_checkpoint_v2" (ByVal db As LongPtr, ByVal zDb As LongPtr, ByVal eMode As Long, ByVal pnLog As LongPtr, ByVal pnCkpt As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_wal_hook CDecl Lib "sqlite3.dll" Alias "sqlite3_wal_hook" (ByVal p1 As LongPtr, ByVal cb2 As LongPtr, ByVal p3 As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3_win32_set_directory CDecl Lib "sqlite3.dll" Alias "sqlite3_win32_set_directory" (ByVal vType As Long, ByVal zValue As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_win32_set_directory16 CDecl Lib "sqlite3.dll" Alias "sqlite3_win32_set_directory16" (ByVal vType As Long, ByVal zValue As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_win32_set_directory8 CDecl Lib "sqlite3.dll" Alias "sqlite3_win32_set_directory8" (ByVal vType As Long, ByVal zValue As LongPtr) As Long

'--- SQLite3 Multiple Ciphers: key handling (sqlite3mc.h)
Public Declare PtrSafe Function stub_sqlite3_key CDecl Lib "sqlite3.dll" Alias "sqlite3_key" (ByVal db As LongPtr, ByVal pKey As LongPtr, ByVal nKey As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_key_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_key_v2" (ByVal db As LongPtr, ByVal zDbName As LongPtr, ByVal pKey As LongPtr, ByVal nKey As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_rekey CDecl Lib "sqlite3.dll" Alias "sqlite3_rekey" (ByVal db As LongPtr, ByVal pKey As LongPtr, ByVal nKey As Long) As Long
Public Declare PtrSafe Function stub_sqlite3_rekey_v2 CDecl Lib "sqlite3.dll" Alias "sqlite3_rekey_v2" (ByVal db As LongPtr, ByVal zDbName As LongPtr, ByVal pKey As LongPtr, ByVal nKey As Long) As Long

'--- SQLite3 Multiple Ciphers: cipher configuration (sqlite3mc.h)
'---   paramName "cipher" selects the cipher of a connection; prefix with
'---   "default:" (db = 0) to change the process-wide default. newValue < 0
'---   only queries the current value.
Public Declare PtrSafe Function stub_sqlite3mc_config CDecl Lib "sqlite3.dll" Alias "sqlite3mc_config" (ByVal db As LongPtr, ByVal paramName As LongPtr, ByVal newValue As Long) As Long
Public Declare PtrSafe Function stub_sqlite3mc_config_cipher CDecl Lib "sqlite3.dll" Alias "sqlite3mc_config_cipher" (ByVal db As LongPtr, ByVal cipherName As LongPtr, ByVal paramName As LongPtr, ByVal newValue As Long) As Long
'--- returned buffer must be released with stub_sqlite3_free
Public Declare PtrSafe Function stub_sqlite3mc_codec_data CDecl Lib "sqlite3.dll" Alias "sqlite3mc_codec_data" (ByVal db As LongPtr, ByVal zDbName As LongPtr, ByVal paramName As LongPtr) As LongPtr
Public Declare PtrSafe Function stub_sqlite3mc_version CDecl Lib "sqlite3.dll" Alias "sqlite3mc_version" () As LongPtr
Public Declare PtrSafe Function stub_sqlite3mc_cipher_count CDecl Lib "sqlite3.dll" Alias "sqlite3mc_cipher_count" () As Long
Public Declare PtrSafe Function stub_sqlite3mc_cipher_index CDecl Lib "sqlite3.dll" Alias "sqlite3mc_cipher_index" (ByVal cipherName As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3mc_cipher_name CDecl Lib "sqlite3.dll" Alias "sqlite3mc_cipher_name" (ByVal cipherIndex As Long) As LongPtr

'--- C-variadic exports. Because the declares are CDecl (caller cleans the
'--- stack) a variadic function can be called through a fixed-arity
'--- declare; these cover the common integer forms:
'---   sqlite3_db_config(db, op, int, int*)  e.g. SQLITE_DBCONFIG_ENABLE_FKEY
'---   sqlite3_config(op, int)               (only before sqlite3_initialize)
Public Declare PtrSafe Function stub_sqlite3_db_config_int CDecl Lib "sqlite3.dll" Alias "sqlite3_db_config" (ByVal db As LongPtr, ByVal op As Long, ByVal newValue As Long, ByVal pOldValue As LongPtr) As Long
Public Declare PtrSafe Function stub_sqlite3_config_int CDecl Lib "sqlite3.dll" Alias "sqlite3_config" (ByVal op As Long, ByVal newValue As Long) As Long
'--- not declared (printf-style, do the formatting VB-side):
'---   sqlite3_log, sqlite3_mprintf, sqlite3_snprintf, sqlite3_str_appendf,
'---   sqlite3_test_control, sqlite3_vtab_config

'--- The following sqlite3 exports are data symbols or internal
'--- helpers with no public prototype in sqlite3.h. Declare when needed:
'---   data:     sqlite3_version (char[]), sqlite3_temp_directory,
'---             sqlite3_data_directory, sqlite3_fts3_may_be_corrupt
'---   internal: sqlite3_win32_is_nt, sqlite3_win32_sleep,
'---             sqlite3_win32_write_debug, sqlite3_win32_mbcs_to_utf8[_v2],
'---             sqlite3_win32_utf8_to_mbcs[_v2], sqlite3_win32_unicode_to_utf8,
'---             sqlite3_win32_utf8_to_unicode
