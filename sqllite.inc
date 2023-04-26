#IF NOT %DEF(%SQLITE_INC)
%SQLITE_INC = 1
'ver 2.23
'
'   modified 1/8/2006
'
'   Thanks to Terence McDonnell, Don Dickinson
'   SQLite 3 databases
'   public domain - use at your own risk
'
TYPE sqlRecSetType
    pDataArry AS STRING PTR  'data array
    pColArry AS STRING PTR   'column names
    RowCount AS LONG         'row count
    ColCount AS LONG         'column count
    phDB AS LONG PTR         'pointer to database handle
    Ndx AS LONG              'index marker for move functions
    IsEof AS LONG            'end of record set - Ndx > RowCount
    IsBof AS LONG            'beginning of record set - Ndx < 1
    HasKey AS LONG           'rowid is first column in results
END TYPE
'
    ''' use this macro to create a record set
    ''' takes care of DIMs
'
    ''' local record set
    ''' sqlRecSetNew(rs ,hDB)
MACRO sqlRecSetNew(tSqlRecSetType, hDB)
    MACROTEMP saSqlRecSetDataArry, saSqlRecSetColArry '''thanks Paul Squires
    LOCAL tSqlRecSetType AS SqlRecSetType
    LOCAL saSqlRecSetDataArry() AS STRING ': REDIM saSqlRecSetDataArry()
    LOCAL saSqlRecSetColArry() AS STRING ': REDIM saSqlRecSetColArry()
'
    tSqlRecSetType.pDataArry = VARPTR(saSqlRecSetDataArry())
    tSqlRecSetType.pColArry = VARPTR(saSqlRecSetColArry())
    tSqlRecSetType.phDB = VARPTR(hDB)
END MACRO
'
    ''' global record set
    ''' sqlRecSetNewGlobal(gRS, hDB)
MACRO sqlRecSetNewGlobal(tgSqlRecSetType, hDB)
    MACROTEMP gsaSqlRecSetDataArry, gsaSqlRecSetColArry
    GLOBAL tgSqlRecSetType AS SqlRecSetType
    GLOBAL gsaSqlRecSetDataArry() AS STRING ': REDIM gsaSqlRecSetDataArry()
    GLOBAL gsaSqlRecSetColArry() AS STRING ': REDIM gsaSqlRecSetColArry()
'
    tgSqlRecSetType.pDataArry = VARPTR(gsaSqlRecSetDataArry())
    tgSqlRecSetType.pColArry = VARPTR(gsaSqlRecSetColArry())
    tgSqlRecSetType.phDB = VARPTR(hDB)
END MACRO
'
'internal use - sqlite3.dll
DECLARE FUNCTION sqlite3_open CDECL LIB "sqlite3.dll" ALIAS "sqlite3_open" (zFilename AS ASCIIZ, hDB AS LONG) AS LONG
DECLARE SUB      sqlite3_close CDECL LIB "sqlite3.dll" ALIAS "sqlite3_close" (BYVAL hDB AS LONG)
DECLARE FUNCTION sqlite_get_table CDECL LIB "sqlite3.dll" ALIAS "sqlite3_get_table" (BYVAL hDB AS LONG, szSql AS ASCIIZ, lpTable AS LONG, nRow AS LONG, nColumn AS LONG, lpErrMsg AS LONG) AS LONG
DECLARE FUNCTION sqlite_free_table CDECL LIB "sqlite3.dll" ALIAS "sqlite3_free_table" (BYVAL lpTable AS LONG PTR) AS LONG
DECLARE FUNCTION sqlite3_errmsg CDECL LIB "sqlite3.dll" ALIAS "sqlite3_errmsg" (BYVAL hDB AS LONG) AS LONG
'internal use - sqlite3.dll
'
DECLARE FUNCTION sqlOpen( BYVAL sFileSpec AS STRING, BYREF hDB AS LONG ) AS LONG
DECLARE SUB sqlClose( BYREF hDB AS LONG )
DECLARE FUNCTION sqlErrMsg( BYVAL hDB AS LONG ) AS STRING
DECLARE SUB sqlRecSetClear( tRS AS sqlRecSetType )
DECLARE FUNCTION sqlSelect( tRS AS sqlRecSetType, sSql AS STRING ) AS LONG
DECLARE FUNCTION sqlExe( BYVAL hDB AS LONG, sSql AS STRING ) AS LONG
DECLARE SUB sqlMoveFirst( tRS AS sqlRecSetType )
DECLARE SUB sqlMoveNext( tRS AS sqlRecSetType )
DECLARE SUB sqlMovePrev( tRS AS sqlRecSetType )
DECLARE SUB sqlMoveTo( tRS AS sqlRecSetType, BYVAL lMoveTo AS LONG )
DECLARE SUB sqlMoveLast( tRS AS sqlRecSetType )
DECLARE FUNCTION sqlColNo( tRS AS sqlRecSetType, BYVAL sColName AS STRING ) AS LONG
DECLARE FUNCTION sqlColName( tRS AS sqlRecSetType, BYVAL lColNo AS LONG ) AS STRING
DECLARE FUNCTION sqlGet( tRS AS sqlRecSetType, BYVAL sColName AS STRING ) AS STRING
DECLARE FUNCTION sqlGetAt( tRS AS sqlRecSetType, BYVAL lColNo AS LONG ) AS STRING
DECLARE SUB sqlSet( tRS AS sqlRecSetType, BYVAL sColName AS STRING, BYVAL sValue AS STRING )
DECLARE SUB sqlSetAt( tRS AS sqlRecSetType, BYVAL lColNo AS LONG, BYVAL sValue AS STRING )
DECLARE FUNCTION sqlUpdateRow( tRS AS sqlRecSetType, BYVAL sTable AS STRING ) AS LONG
DECLARE FUNCTION sqlUpdateRecSet( tRS AS sqlRecSetType, BYVAL sTable AS STRING ) AS LONG
DECLARE SUB sqlTableList( BYVAL hDB AS LONG, saTableList() AS STRING )
DECLARE FUNCTION sqlTableExist( BYVAL hDB AS LONG, BYVAL sTable AS STRING ) AS LONG
DECLARE SUB sqlColList( BYVAL hDB AS LONG, BYVAL sTable AS STRING, saColList() AS STRING )
DECLARE SUB sqlIndexList( BYVAL hDB AS LONG, BYVAL sTable AS STRING, saIndexList() AS STRING )
DECLARE FUNCTION sqlFix( BYVAL sString AS STRING) AS STRING
DECLARE FUNCTION sqlDate( BYVAL sDate AS STRING ) AS STRING
DECLARE FUNCTION sqlTimeStamp() AS STRING
'
'
'#INCLUDE "Encapsule.inc"
'
' --------------------------------------------------
FUNCTION sqlOpen( BYVAL sFileSpec AS STRING, BYREF hDB AS LONG ) AS LONG
    LOCAL szFileSpec AS ASCIIZ * 400
    ''' create or open database
    ''' set database handle
    ''' success=-1 / error=0
    szFileSpec = sFileSpec
    FUNCTION = IIF&( sqlite3_open( szFileSpec, hDB )=0, -1, 0 )
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
SUB sqlClose( BYREF hDB AS LONG )
    IF ISTRUE hDB THEN
        sqlExe(hDB, "commit")
        sqlite3_close(hDB)
        hDB = 0
    END IF
END SUB
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlErrMsg( BYVAL hDB AS LONG ) AS STRING
    'Thanks to Don Dickinson
    DIM pzErr AS ASCIIZ PTR

    ''' get the SQLite error message
    ''' ? sqlErrMsg(hDB)

    pzErr = sqlite3_errmsg(hDB)
    IF pzErr = 0 THEN
        FUNCTION = ""
    ELSE
        FUNCTION = @pzErr
    END IF
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
SUB sqlRecSetClear( tRS AS sqlRecSetType )
    'clear array memory
    EncapStrArryReDim BYVAL tRS.pDataArry, -1
    EncapStrArryReDim BYVAL tRS.pColArry, -1
    tRS.RowCount = 0
    tRS.ColCount = 0
    tRS.Ndx = -1
    tRS.IsEof = -1
    tRS.IsBof = -1
    tRS.HasKey = 0
END SUB
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlSelect( tRS AS sqlRecSetType, sSql AS STRING ) AS LONG
    LOCAL i, x       AS LONG
    LOCAL lpTable    AS LONG PTR     ' Array of fields returned from get_table (starts with column names)
    LOCAL lpErrorSz  AS LONG         ' Error msg
    LOCAL pzField    AS ASCIIZ PTR   ' Field return from get_table (element in lpTable array)
    LOCAL lRow, lCol AS LONG
'
    ''' execute query and return record set
    '''
    ''' success=-1 / error=0
    '''
    ''' number rows = tRS.RowCount
    ''' number coulmns = tRS.ColCount
'
    ''' thanks to Terence McDonnell
'
    sqlRecSetClear tRS
    IF ISFALSE tRS.@phDB THEN EXIT FUNCTION
'
    IF sqlite_get_table(tRS.@phDB, BYVAL STRPTR(sSQL), lpTable, tRS.RowCount, tRS.ColCount, lpErrorSz)<>0 THEN
        FUNCTION = 0
    ELSEIF tRS.RowCount=0 THEN
        FUNCTION = -1
        sqlite_free_table lpTable
    ELSE
        'column names
        EncapStrArryReDim BYVAL tRS.pColArry, tRS.ColCount
        FOR i=1 TO tRS.ColCount
            pzField = @lpTable[i-1]
            EncapStrArrySet BYVAL tRS.pColArry, i, @pzField
        NEXT i
'
        'data
        EncapStrArryReDim2 BYVAL tRS.pDataArry, tRS.RowCount, tRS.ColCount
        x = tRS.ColCount 'skip col row
        lRow = 1
        FOR i = 1 TO tRS.RowCount
            FOR lCol=1 TO tRS.ColCount
                pzField = @lpTable[x]
                EncapStrArrySet2 BYVAL tRS.pDataArry, lRow, lCol, @pzField
                INCR x
            NEXT j
            INCR lRow
        NEXT i
        FUNCTION = -1
        IF LCASE$(sqlColName(tRS,1)) = "rowid" THEN
            tRS.HasKey = -1
        ELSE
            tRS.HasKey = 0
        END IF
        sqlite_free_table lpTable
    END IF
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlExe( BYVAL hDB AS LONG, sSql AS STRING ) AS LONG
    LOCAL lpTable    AS LONG PTR
    LOCAL lpErrorSz  AS LONG
    LOCAL RowCount&, ColCount&
'
    ''' execute no return SQL statement
    ''' success=-1 / error=0
    ''' ? sqlErrMsg(hDB)
'
    FUNCTION = IIF&( sqlite_get_table(hDB, BYVAL STRPTR(sSQL), lpTable, RowCount&, ColCount&, lpErrorSz)=0, -1, 0 )
    sqlite_free_table lpTable
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
SUB sqlMoveFirst( tRS AS sqlRecSetType )
    'move first row
    tRS.Ndx = 1
    tRS.IsBof = 0
    tRS.IsEof = 0
    IF tRS.RowCount<1 THEN
        tRS.IsBof = -1
        tRS.IsEof = -1
    ELSEIF tRS.Ndx>tRS.RowCount THEN
        tRS.IsEof = -1
    END IF
END SUB
' --------------------------------------------------
' --------------------------------------------------
SUB sqlMoveNext( tRS AS sqlRecSetType )
    'move next row
    INCR tRS.Ndx
    tRS.IsBof = 0
    tRS.IsEof = 0
    IF tRS.RowCount<1 THEN
        tRS.IsBof = -1
        tRS.IsEof = -1
    ELSEIF tRS.Ndx<1 THEN
        tRS.IsBof = -1
    ELSEIF tRS.Ndx>tRS.RowCount THEN
        tRS.IsEof = -1
    END IF
END SUB
' --------------------------------------------------
' --------------------------------------------------
SUB sqlMovePrev( tRS AS sqlRecSetType )
    'move backwards
    DECR tRS.Ndx
    tRS.IsBof = 0
    tRS.IsEof = 0
    IF tRS.RowCount<1 THEN
        tRS.IsBof = -1
        tRS.IsEof = -1
    ELSEIF tRS.Ndx<1 THEN
        tRS.IsBof = -1
    ELSEIF tRS.Ndx>tRS.RowCount THEN
        tRS.IsEof = -1
    END IF
END SUB
' --------------------------------------------------
' --------------------------------------------------
SUB sqlMoveTo( tRS AS sqlRecSetType, BYVAL lMoveTo AS LONG )
    'move to lMoveTo
    tRS.Ndx = lMoveTo
    tRS.IsBof = 0
    tRS.IsEof = 0
    IF tRS.RowCount<1 THEN
        tRS.IsBof = -1
        tRS.IsEof = -1
    ELSEIF tRS.Ndx<1 THEN
        tRS.IsBof = -1
    ELSEIF tRS.Ndx>tRS.RowCount THEN
        tRS.IsEof = -1
    END IF
END SUB
' --------------------------------------------------
' --------------------------------------------------
SUB sqlMoveLast( tRS AS sqlRecSetType )
    'move to last row
    tRS.Ndx = tRS.RowCount
    tRS.IsBof = 0
    tRS.IsEof = 0
    IF tRS.RowCount<1 THEN
        tRS.IsBof = -1
        tRS.IsEof = -1
    ELSEIF tRS.Ndx<1 THEN
        tRS.IsBof = -1
    ELSEIF tRS.Ndx>tRS.RowCount THEN
        tRS.IsEof = -1
    END IF
END SUB
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlColNo( tRS AS sqlRecSetType, BYVAL sColName AS STRING ) AS LONG
    ''' get column number from record set column list
    ''' return 0 = not found

    FUNCTION = EncapStrArryScanNC( BYVAL tRS.pColArry, sColName )
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlColName( tRS AS sqlRecSetType, BYVAL lColNo AS LONG ) AS STRING
    ''' get column name from record set column list
    ''' return "" = not found

    FUNCTION = EncapStrArryGet( BYVAL tRS.pColArry, lColNo )
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlGet( tRS AS sqlRecSetType, BYVAL sColName AS STRING ) AS STRING
    LOCAL x AS LONG

    'get column value by column name

    x = sqlColNo(tRS, sColName)
    IF x<1 THEN EXIT FUNCTION
    FUNCTION = EncapStrArryGet2( BYVAL tRS.pDataArry, tRS.Ndx, x )
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlGetAt( tRS AS sqlRecSetType, BYVAL lColNo AS LONG ) AS STRING
    'get column value by column number
    'first column = 1

    ''''''''' fixed - had tRS.RowCount
    IF lColNo<1 OR lColNo>tRS.ColCount THEN EXIT FUNCTION

    FUNCTION = EncapStrArryGet2( BYVAL tRS.pDataArry, tRS.Ndx, lColNo )
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
SUB sqlSet( tRS AS sqlRecSetType, BYVAL sColName AS STRING, BYVAL sValue AS STRING )
    LOCAL x AS LONG

    'set column value by column name

    x = sqlColNo(tRS, sColName)
    IF x<1 THEN EXIT SUB
    EncapStrArrySet2 BYVAL tRS.pDataArry, tRS.Ndx, x, sValue
END SUB
' --------------------------------------------------
' --------------------------------------------------
SUB sqlSetAt( tRS AS sqlRecSetType, BYVAL lColNo AS LONG, BYVAL sValue AS STRING )
    'set column value by column number
    'first column = 1

    EncapStrArrySet2 BYVAL tRS.pDataArry, tRS.Ndx, lColNo, sValue
END SUB
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlUpdateRow( tRS AS sqlRecSetType, BYVAL sTable AS STRING ) AS LONG
    LOCAL i AS LONG
    LOCAL sSql AS STRING
    LOCAL sKeyVal AS STRING

    ' update current row
    ' [rowid] must be the first column in record set
    ' "SELECT rowid,* FROM MyTable"
    ' rowid won't be in results unless you specifically ask for it
    '
    ' WARNING: you shouldn't use sqlUpdateRow() if you use direct
    ' SQL statements to modify the Table while moving through a
    ' RecordSet
    '
    ' Either use direct SQL statements OR sqlUpdateRow(), NOT both.
    '
    IF ISFALSE tRS.HasKey THEN EXIT FUNCTION
    IF tRS.ColCount<2 THEN EXIT FUNCTION

    sKeyVal = sqlGetAt(tRS, 1)
    sSql = "update ["+sTable+"] set "
    FOR i=2 TO tRS.ColCount
        sSql = sSql + "["+sqlColName(tRS,i)+"]='"+sqlFix(sqlGetAt(tRS,i))+"', "
    NEXT i
    sSql = RTRIM$(sSql,", ") + " where rowid="+sKeyVal
    FUNCTION = sqlExe(tRS.@phDB, sSql)
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlUpdateRecSet( tRS AS sqlRecSetType, BYVAL sTable AS STRING ) AS LONG
    ' update whole record set
    '
    'WARNING: you shouldn't use sqlUpdateRecSet() if you use direct
    'SQL statements to modify the Table while moving through a
    'RecordSet
    '
    'Either use direct SQL statements OR sqlUpdateRecSet(), NOT both
    '
    IF ISFALSE tRS.@phDB THEN EXIT FUNCTION
    sqlMoveFirst tRS
    WHILE NOT tRS.IsEof
        IF ISFALSE sqlUpdateRow(tRS,sTable) THEN EXIT FUNCTION
        sqlMoveNext tRS
    WEND
    FUNCTION = -1
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
SUB sqlTableList( BYVAL hDB AS LONG, saTableList() AS STRING )
    LOCAL i AS LONG
    ''' get Table List for database

    sqlRecSetNew(rs, hDB)

    REDIM saTableList()
    IF ISFALSE hDB THEN EXIT SUB

    IF ISFALSE sqlSelect(rs, "select name from sqlite_master where type = 'table' order by name collate nocase") THEN EXIT SUB
    IF rs.RowCount<1 THEN EXIT SUB

    REDIM saTableList(1 TO rs.RowCount)
    FOR i=1 TO rs.RowCount
        sqlMoveTo rs, i
        saTableList(i) = sqlGetAt(rs, 1)
    NEXT i
END SUB
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlTableExist( BYVAL hDB AS LONG, BYVAL sTable AS STRING ) AS LONG
    ''' see if sTable exist
    IF ISFALSE hDB THEN EXIT FUNCTION
    sqlRecSetNew(rs, hDB)
    IF ISFALSE sqlSelect(rs, "Select * From sqlite_master where name='"+sTable+"' and type='table'") THEN EXIT FUNCTION
    IF ISTRUE rs.RowCount THEN
        FUNCTION = -1
    ELSE
        FUNCTION = 0
    END IF
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
SUB sqlColList( BYVAL hDB AS LONG, BYVAL sTable AS STRING, saColList() AS STRING )
    LOCAL i AS LONG
    ''' get column list for sTable

    sqlRecSetNew(rs, hDB)

    REDIM saColList()
    IF ISFALSE hDB THEN EXIT SUB

    IF ISFALSE sqlSelect( rs, "PRAGMA table_info(["+sTable+"])") THEN EXIT SUB
    IF rs.RowCount<1 THEN EXIT SUB

    REDIM saColList(1 TO rs.RowCount)
    FOR i=1 TO rs.RowCount
        sqlMoveTo rs, i
        saColList(i) = sqlGetAt(rs, 2)
    NEXT i
END SUB
' --------------------------------------------------
' --------------------------------------------------
SUB sqlIndexList( BYVAL hDB AS LONG, BYVAL sTable AS STRING, saIndexList() AS STRING )
    LOCAL i AS LONG
    ''' get index list for sTable

    sqlRecSetNew(rs, hDB)

    REDIM saIndexList()
    IF ISFALSE hDB THEN EXIT SUB

    IF ISFALSE sqlSelect( rs, "PRAGMA index_list(["+sTable+"])") THEN EXIT SUB
    IF rs.RowCount<1 THEN EXIT SUB

    REDIM saIndexList(1 TO rs.RowCount)
    FOR i=1 TO rs.RowCount
        sqlMoveTo rs, i
        saIndexList(i) = sqlGetAt(rs, 2)
    NEXT i
END SUB
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlFix( BYVAL sString AS STRING) AS STRING
    ''' fix single quotes
    REPLACE "'" WITH "''" IN sString
    FUNCTION = sString
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlDate( BYVAL sDate AS STRING ) AS STRING
    LOCAL d$, m$, y$
    'change most dates to SQL date - 2006-01-25
    '
    'basic date$ - ok
    '11/1/05 - ok
    '11,01,05 - ok
    '11.1.2005 - ok
    '1-23-05 - ok
    'YYYYMMDD - ok
    '1-23-99 won't fix = 2099-01-23

    IF sDate="" THEN EXIT FUNCTION
    REPLACE "." WITH "," IN sDate
    REPLACE "-" WITH "," IN sDate
    REPLACE "/" WITH "," IN sDate
    REPLACE "\" WITH "," IN sDate

    IF INSTR(sDate,",")=0 THEN
        ''' date is in YYYYMMDD format?
        FUNCTION = LEFT$(sDate,4) +"-"+ MID$(sDate,5,2) +"-"+ RIGHT$(sDate,2)
        EXIT FUNCTION
    END IF

    d$ = PARSE$(sDate,",",2)
    m$ = PARSE$(sDate,",",1)
    y$ = PARSE$(sDate,",",3)
    IF LEN(d$)=1 THEN d$ = "0"+d$
    IF LEN(m$)=1 THEN m$ = "0"+m$
    IF LEN(y$)=2 THEN y$ = "20"+y$

    FUNCTION = y$+"-"+m$+"-"+d$
END FUNCTION
' --------------------------------------------------
' --------------------------------------------------
FUNCTION sqlTimeStamp() AS STRING
    'current SQL Time Stamp = "2005-11-30 17:15:46"
    FUNCTION = sqlDate(DATE$) +" "+ TIME$
END FUNCTION
' --------------------------------------------------

#ENDIF
