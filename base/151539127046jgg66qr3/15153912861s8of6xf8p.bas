Attribute VB_Name = "modConnect"
'---------------------------------------------------------------------------------------
' Module    : modConnect
' Release   : 2.0
' DateTime  : 22.01.2007 10:43
' Author    : DSonnyh (aka Joss)
' Purpose   : Автоматизация подключения таблиц из внешней базы даннех
'---------------------------------------------------------------------------------------
Option Compare Database
Option Explicit

'---------------------------------------------------------------------------------------
' Procedure : ClearTablesRefBase
' DateTime  : 08.01.2007 12:47
' Author    : DSonnyh
' Purpose   : Отключение подключенных таблиц, принадлежащих указанной базе
'---------------------------------------------------------------------------------------
'
Public Sub ClearTablesRefBase(strBaseName As String)
' подпрограмма отключения всех прилинкованных таблиц.
    Dim cat As ADOX.Catalog
    Dim tdf As ADOX.Table
    Dim i As Integer
    
   On Error GoTo ClearTablesRefBase_Error

    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
mmm:
    i = 0
    For Each tdf In cat.Tables
' Debug.Print cat.Tables(i).Name, cat.Tables(i).Type
' Проверка на линкованную таблицу
       If cat.Tables(i).Type = "LINK" Then
          If cat.Tables(i).Properties(6).Value = strBaseName Then
              cat.Tables.Delete (i)
              GoTo mmm
          End If
       End If
       i = i + 1
    Next tdf
    Set cat = Nothing
        

   On Error GoTo 0
Exit_ClearTablesRefBase:
   Exit Sub

ClearTablesRefBase_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре ClearTablesRefBase в Module Module1"
    Resume Exit_ClearTablesRefBase

End Sub
Public Sub ClearTablesRefADO()
' подпрограмма отключения всех прилинкованных таблиц.
    Dim cat As ADOX.Catalog
    Dim tdf As ADOX.Table
    Dim i As Integer
    
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
mmm:
    i = 0
    For Each tdf In cat.Tables
' Debug.Print cat.Tables(i).Name, cat.Tables(i).Type
       If cat.Tables(i).Type = "LINK" Then
          cat.Tables.Delete (i)
          GoTo mmm
       End If
       i = i + 1
    Next tdf
    Set cat = Nothing
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetTableRefBase
' DateTime  : 09.01.2007 15:16
' Author    : DSonnyh
' Purpose   : Подключение таблиц
'---------------------------------------------------------------------------------------
'
Public Sub SetTableRefBase(TName As String, NewTName As String, strBase As String, _
                            Optional intControl As Variant = 3)
On Error GoTo Err_SetTableRefBase

Select Case intControl
   Case 0
             Call SetTableRefBaseLite(TName, NewTName, strBase)
   Case 1
             Call SetTableRefBaseLiteOne(TName, NewTName, strBase)
   Case 2
             Call SetTableRefBaseLiteTwo(TName, NewTName, strBase)
   Case Else
             Call SetTableRefBaseFull(TName, NewTName, strBase)
End Select
   
   
Exit_SetTableRefBase:
    Exit Sub

Err_SetTableRefBase:
    MsgBox Err.Description
    Resume Exit_SetTableRefBase
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetTableRefBaseFull
' DateTime  : 22.01.2007 14:54
' Author    : DSonnyh
' Purpose   : Подключение таблиц с полной проверкой существования таблиц
'---------------------------------------------------------------------------------------
'
Public Sub SetTableRefBaseFull(TName As String, NewTName As String, strBase As String)

   On Error GoTo SetTableRefBaseFull_Error

Dim strCurrentDB As String
strCurrentDB = CurrentProject.FullName
' проверка наличия таблицы в подключаемой базе
   If IsTable(TName, strBase) = 1 Then
        If IsTable(NewTName, strCurrentDB) <> 1 Then
' проверка наличия таблицы в текущей базе
           DoCmd.TransferDatabase acLink, "Microsoft Access", _
           strBase, acTable, TName, NewTName, False, False
        Else
            Call MsgBox("В базе " & strCurrentDB _
                        & vbCrLf & "уже существует таблица " & NewTName _
                        & vbCrLf & "Подключение данной таблицы не проведено!" _
                        , vbExclamation, "Подключение таблиц")

        End If
   Else
        Call MsgBox("В базе " & strBase _
                    & vbCrLf & "отсутствует таблица " & TName _
                    & vbCrLf & "Подключение данной таблицы не проведено!" _
                    , vbExclamation, "Подключение таблиц")
   
   End If

   On Error GoTo 0
Exit_SetTableRefBaseFull:
   Exit Sub

SetTableRefBaseFull_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре SetTableRefBaseFull в Module modConnect"
    Resume Exit_SetTableRefBaseFull

End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetTableRefBaseLiteOne
' DateTime  : 11.01.2007 15:26
' Author    : DSonnyh
' Purpose   : Проверка только наличия таблицы в подключаемой базе
'---------------------------------------------------------------------------------------
'
Public Sub SetTableRefBaseLiteOne(TName As String, NewTName As String, strBase As String)
On Error GoTo Err_SetTableRefBaseLiteOne
Dim strCurrentDB As String
strCurrentDB = CurrentProject.FullName
' проверка наличия таблицы в подключаемой базе
   If IsTable(TName, strBase) = 1 Then
           DoCmd.TransferDatabase acLink, "Microsoft Access", _
           strBase, acTable, TName, NewTName, False, False
        Else
Call MsgBox("В базе " & strCurrentDB _
            & vbCrLf & "уже существует таблица " & NewTName _
            & vbCrLf & "Подключение данной таблицы не проведено!" _
            , vbExclamation, "Подключение таблиц")
   End If
   
Exit_SetTableRefBaseLiteOne:
    Exit Sub

Err_SetTableRefBaseLiteOne:
    MsgBox Err.Description
    Resume Exit_SetTableRefBaseLiteOne
End Sub
'---------------------------------------------------------------------------------------
' Procedure : SetTableRefBaseLiteTwo
' DateTime  : 22.01.2007 14:48
' Author    : DSonnyh
' Purpose   : Проверка только наличия таблицы в текущей базе
'---------------------------------------------------------------------------------------
'
Public Sub SetTableRefBaseLiteTwo(TName As String, NewTName As String, strBase As String)
On Error GoTo Err_SetTableRefBaseLiteTwo
           
Dim strCurrentDB As String
strCurrentDB = CurrentProject.FullName
        If IsTable(NewTName, strCurrentDB) <> 1 Then
' проверка наличия таблицы в текущей базе
           DoCmd.TransferDatabase acLink, "Microsoft Access", _
           strBase, acTable, TName, NewTName, False, False
        Else
Call MsgBox("В базе " & strCurrentDB _
            & vbCrLf & "уже существует таблица " & NewTName _
            & vbCrLf & "Подключение данной таблицы не проведено!" _
            , vbExclamation, "Подключение таблиц")

        End If
   
Exit_SetTableRefBaseLiteTwo:
    Exit Sub

Err_SetTableRefBaseLiteTwo:
    MsgBox Err.Description
    Resume Exit_SetTableRefBaseLiteTwo
End Sub

'---------------------------------------------------------------------------------------
' Procedure : SetTableRefBaseLite
' DateTime  : 22.01.2007 14:47
' Author    : DSonnyh
' Purpose   : Подключение таблиц без проверки
'---------------------------------------------------------------------------------------
'
Public Sub SetTableRefBaseLite(TName As String, NewTName As String, strBase As String)

   On Error GoTo SetTableRefBaseLite_Error

           DoCmd.TransferDatabase acLink, "Microsoft Access", _
           strBase, acTable, TName, NewTName, False, False

   On Error GoTo 0
Exit_SetTableRefBaseLite:
   Exit Sub

SetTableRefBaseLite_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре SetTableRefBaseLite в Module modConnect"
    Resume Exit_SetTableRefBaseLite

End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsTable, Release 2.0
' DateTime  : 15.01.2007 13:28
' Author    : DSonnyh
' Purpose   : Вариант вызова фукции, используемый для совместимости со старыми версиями
'---------------------------------------------------------------------------------------
'
Public Function IsTable(Name As String, Optional strBase As Variant, _
                        Optional intMetod As Variant = 0) As Integer

' проверка на наличие определенной таблицы
' Name - наименование таблицы
' strBase - полный путь к файлу базы (опционально), по умолчанию - текущая
' intMetod - используемый метод 0 - DAO или 1 - ADO (опционально), по умолчанию - DAO

   On Error GoTo IsTable_Error
   Dim strBaseName As String, intMetodA As Integer
      
   If IsMissing(strBase) Then
        If intMetod = 0 Then
            strBaseName = CurrentDb.Name
            IsTable = IsTableDAO(Name, strBaseName)
        Else
            strBaseName = CurrentProject.FullName
            IsTable = IsTableADO(Name, strBaseName)
        End If
   Else
        strBaseName = strBase
        If intMetod = 0 Then
            IsTable = IsTableDAO(Name, strBaseName)
        Else
            IsTable = IsTableADO(Name, strBaseName)
        End If
   End If

   On Error GoTo 0
Exit_IsTable:
   Exit Function

IsTable_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре IsTable " _
         & "в Module modConnect"
    Resume Exit_IsTable

End Function

'---------------------------------------------------------------------------------------
' Procedure : IsTableDAO, Release 2.0
' DateTime  : 15.01.2007 10:39
' Author    : DSonnyh
' Purpose   : Проверка наличия таблицы в указанной базе данных
'---------------------------------------------------------------------------------------
'
Public Function IsTableDAO(Name As String, strBase As String) As Integer
' проверка на наличие определенной таблицы
' Name - наименование таблицы
' strBase - полный путь к файлу базы
   On Error GoTo IsTableDAO_Error
    Dim dbs As Database, tdf As TableDef
    
    If strBase = CurrentDb.Name Then
        Set dbs = CurrentDb
    Else
        Set dbs = DBEngine.Workspaces(0).OpenDatabase(strBase)
    End If
    
    IsTableDAO = 0
    For Each tdf In dbs.TableDefs
        If tdf.Name = Name Then
            IsTableDAO = 1
            Exit For 'Function
        End If
    Next tdf
    Set tdf = Nothing
    dbs.Close
    Set dbs = Nothing

   On Error GoTo 0
Exit_IsTableDAO:
   Exit Function

IsTableDAO_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре IsTableDAO " _
         & "в Module modConnect"
    Resume Exit_IsTableDAO

End Function
'---------------------------------------------------------------------------------------
' Procedure : IsTableADO, Release 2.0
' DateTime  : 11.01.2007 16:59
' Author    : DSonnyh
' Purpose   : Проверка наличия таблицы в указанной базе данных
'---------------------------------------------------------------------------------------
'
Public Function IsTableADO(Name As String, strBase As String) As Integer
' проверка на наличие определенной таблицы. Метод ADO
' Name - наименование таблицы
' strBase - полный путь к файлу базы

    Dim cnn As ADODB.Connection
    Dim cat As ADOX.Catalog
    Dim tdf As ADOX.Table
   On Error GoTo IsTable_Error

    Set cnn = New ADODB.Connection
    Set cat = New ADOX.Catalog
    
    Dim strProvider As String
' строка подклюбчения
    strProvider = "Provider=Microsoft.Jet.OLEDB.4.0;User ID=;" _
            & "Data Source=" & strBase & ";Mode=Share Deny None;" _
            & "Extended Properties=" & "''" & ";Jet OLEDB:System database=;Jet O"
    
    If strBase = CurrentProject.FullName Then
        cat.ActiveConnection = CurrentProject.Connection
       Else
        cnn.ConnectionString = strProvider
        cnn.Open
        cat.ActiveConnection = cnn.ConnectionString
    End If
        
    IsTableADO = 0
    
    For Each tdf In cat.Tables
        If tdf.Name = Name Then
            IsTableADO = 1
            Exit For
        End If
    Next tdf
    
    Set tdf = Nothing
    Set cat = Nothing
    
    If strBase <> CurrentProject.FullName Then
        cnn.Close
    End If
    Set cnn = Nothing

   On Error GoTo 0
Exit_IsTable:
   Exit Function

IsTable_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре IsTable " _
          & "в Module1"
    Resume Exit_IsTable

End Function

'---------------------------------------------------------------------------------------
' Procedure : IsBases
' DateTime  : 10.01.2007 16:37
' Author    : DSonnyh
' Purpose   : Проверка наличия баз данных и доступа к ним
'---------------------------------------------------------------------------------------
'
Public Sub IsBases()
' проверка наличия баз данных и доступа к ним
On Error GoTo Err_IsBases
   Dim rstTemp2 As ADODB.Recordset
   Dim cnn As ADODB.Connection
   Dim PName As String
   Dim sss As String
   Dim strBasePath As String
   Dim i As Integer
   Dim a As Integer
    Dim strHide As String
    Dim iMsg As Integer
    
    strHide = Screen.ActiveForm.Name
    
    Dim stDocName As String
    Dim stLinkCriteria As String
    stDocName = "SystemBases"
   
   Set rstTemp2 = New ADODB.Recordset
   
   Set cnn = New ADODB.Connection
   
   
   rstTemp2.Open "SystemBases", CurrentProject.Connection
   
   Do Until rstTemp2.EOF
' перебор и проверказадания путей к базам данных
       
       strBasePath = Nz(rstTemp2!PathBase)
       If Len(strBasePath) = 0 Then
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
'          MsgBox "Не задан путь к базе данных " + rstTemp2!NickBase, vbOKOnly, "Путь к базе данных"
        iMsg = MsgBox("Не задан путь к базе данных '" & rstTemp2!NickBase & "'" & Chr(13) & "Будете задавать путь?", 68, "Путь к базе данных")

Select Case iMsg

    Case vbYes
    'Place conditional Code here
    stLinkCriteria = "[Id]=" & rstTemp2![Id]
    
    If Len(strHide) > 0 Then
        Screen.ActiveForm.Visible = False
    End If
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, "0"
    If Len(strHide) > 0 Then
        Forms(strHide).Visible = True
    End If
    

    Case vbNo
    'Place conditional Code here

End Select
        
        Else
          If Not fnFileExists(strBasePath) Then
'             MsgBox "Не найдена базе данных " + strBasePath, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
        iMsg = MsgBox("Не найдена базе данных '" & rstTemp2!NickBase & "'" & Chr(13) & "Будете задавать путь?", 68, "Путь к базе данных")

Select Case iMsg

    Case vbYes
    'Place conditional Code here
    stLinkCriteria = "[Id]=" & rstTemp2![Id]
    If Len(strHide) > 0 Then
        Screen.ActiveForm.Visible = False
    End If
    DoCmd.OpenForm stDocName, , , stLinkCriteria, , acDialog, "0"
    If Len(strHide) > 0 Then
        Forms(strHide).Visible = True
    End If

    Case vbNo
    'Place conditional Code here


End Select
          End If
       End If
       rstTemp2.MoveNext
   Loop
   rstTemp2.Close

Exit_IsBases:
    Exit Sub

Err_IsBases:
    If Err.Number = 2475 Then
        strHide = ""
        Resume Next
    End If
    
    MsgBox Err.Description
'!!!
'    DoCmd.Quit
    Resume Exit_IsBases

End Sub

'---------------------------------------------------------------------------------------
' Procedure : ClearTablesRefDAO
' DateTime  : 19.01.2007 10:26
' Author    : DSonnyh
' Purpose   : Отключение всех подлинкованных таблиц. Метод DAO
'---------------------------------------------------------------------------------------
'
Public Sub ClearTablesRefDAO()
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim sss As Variant, k As Integer
    Set dbs = CurrentDb
mmm:
    For Each tdf In dbs.TableDefs
       sss = tdf.Connect
        If Len(sss) > 0 Then
            dbs.TableDefs.Delete (tdf.Name)
            GoTo mmm
        End If
    Next tdf
End Sub

'---------------------------------------------------------------------------------------
' Procedure : IsBasesControlS
' DateTime  : 05.05.2006 11:29
' Author    : DSonnyh
' Purpose   : Проверка наличия подключаемых баз
'---------------------------------------------------------------------------------------
'
Public Function IsBasesControlS() As Boolean

    Dim bResult As Boolean

   On Error GoTo IsBasesControlS_Error

   Dim rstTemp2 As ADODB.Recordset
   Dim strBasePath As String
   
   Set rstTemp2 = New ADODB.Recordset
      
   rstTemp2.Open "SystemBases", CurrentProject.Connection
   bResult = True
   Do Until rstTemp2.EOF
' перебор и проверказадания путей к базам данных
       strBasePath = Nz(rstTemp2!PathBase)
       If Len(strBasePath) = 0 Then
'          MsgBox "Не задан путь к базе данных " + rstTemp2!NickBase, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
            bResult = False
            Exit Do
        Else
          If Not fnFileExists(strBasePath) Then
'             MsgBox "Не найдена базе данных " + strBasePath, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
            bResult = False
            Exit Do
          End If
       End If
       rstTemp2.MoveNext
   Loop
   rstTemp2.Close
   
   Set rstTemp2 = Nothing
   
   IsBasesControlS = bResult

Exit_IsBasesControlS:

   On Error GoTo 0
   Exit Function

IsBasesControlS_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsBasesControlS of Module Module1"
    Resume Exit_IsBasesControlS

End Function

'---------------------------------------------------------------------------------------
' Procedure : IsBasesControlA
' DateTime  : 05.05.2006 11:29
' Author    : DSonnyh
' Purpose   : Проверка наличия подключаемых баз
'---------------------------------------------------------------------------------------
'
Public Function IsBasesControlA() As Boolean

    Dim bResult As Boolean

   On Error GoTo IsBasesControlS_Error

   Dim rstTemp2 As ADODB.Recordset
   Dim strBasePath As String
   Dim strPB As String
   
   Set rstTemp2 = New ADODB.Recordset
      
   rstTemp2.Open "SystemBases", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   bResult = True
   Do Until rstTemp2.EOF
' перебор и проверка задания путей к базам данных
       strBasePath = Nz(rstTemp2!PathBase)
       If Len(strBasePath) = 0 Then
'          MsgBox "Не задан путь к базе данных " + rstTemp2!NickBase, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.

        Select Case MsgBox("Не задан путь к базе данных - " & rstTemp2!NickBase _
                           & vbCrLf & "Будете задавать сейчас" _
                           , vbYesNo Or vbExclamation Or vbDefaultButton1, "Подключаемые базы")
        
            Case vbYes

                If GetDBFileNameDlg(0, strBasePath) Then
                     strPB = strBasePath
                     strPB = Left(strPB, InStr(strPB, Chr(0)) - 1)
                     rstTemp2.Fields("PathBase") = strPB
                     rstTemp2.Update
                    Else
                     bResult = False
                     Exit Do
                End If

            Case vbNo
                bResult = False
                Exit Do

        End Select
        
        Else
          If Not fnFileExists(strBasePath) Then
'             MsgBox "Не найдена базе данных " + strBasePath, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
Select Case MsgBox("Не найдена базе данных - " & rstTemp2!NickBase _
                   & vbCrLf & "Будете задавать сейчас" _
                   , vbYesNo Or vbExclamation Or vbDefaultButton1, "Подключаемые базы")

    Case vbYes

           If GetDBFileNameDlg(0, strBasePath) Then
                strPB = strBasePath
                strPB = Left(strPB, InStr(strPB, Chr(0)) - 1)
                rstTemp2.Fields("PathBase") = strPB
                rstTemp2.Update
               Else
                bResult = False
                Exit Do
           End If

    Case vbNo
            bResult = False
            Exit Do

End Select
            
          End If
       End If
       rstTemp2.MoveNext
   Loop
   rstTemp2.Close
   
   Set rstTemp2 = Nothing
   
   IsBasesControlA = bResult

Exit_IsBasesControlS:

   On Error GoTo 0
   Exit Function

IsBasesControlS_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsBasesControlS of Module Module1"
    Resume Exit_IsBasesControlS

End Function

'---------------------------------------------------------------------------------------
' Procedure : IsBasesConnect
' DateTime  : 17.05.2006 14:15
' Author    : DSonnyh
' Purpose   : подпрограмма проверки правильности подключения к базам. Метод ADO
'---------------------------------------------------------------------------------------
'
Public Function IsBasesConnect() As Boolean
   On Error GoTo IsBasesConnect_Error

    Dim bResult As Boolean
    Dim tdf As Table   'tdf As TableDef
    Dim sss As Variant, k As Integer, intCount As Integer
    Dim strTblName As String, sst As String, blnCBase As Boolean
    Dim cat As ADOX.Catalog
    
    Set cat = New ADOX.Catalog
    cat.ActiveConnection = CurrentProject.Connection
' определяем число прилинкованных таблиц в таблице SystemTables
    Dim rst As ADODB.Recordset
    Set rst = New ADODB.Recordset
    Dim strSQL As String
    strSQL = "select * from SystemTB;"
    rst.Open strSQL, CurrentProject.Connection, adOpenKeyset, adLockReadOnly
    intCount = rst.RecordCount
    rst.Close
    Set rst = Nothing
    
    blnCBase = True
    k = 0
    For Each tdf In cat.Tables
        If tdf.Type = "LINK" Then
            k = k + 1
            strTblName = tdf.Name
            sss = tdf.Properties.Item(6)
            sst = DLookup("[PathBase]", "SystemTB", "TName = '" + strTblName + "'")
            If sss <> sst Then
                blnCBase = False
                Exit For
            End If
        End If
    Next tdf
    
    If k <> intCount Then
        blnCBase = False
    End If
        
    IsBasesConnect = blnCBase 'bResult

   On Error GoTo 0
   Exit Function

IsBasesConnect_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure IsBasesConnect of Module Module1"

End Function

'---------------------------------------------------------------------------------------
' Procedure : fnPathDB
' DateTime  : 17.08.2006 12:58
' Author    : DSonnyh
' Purpose   : Определение пути к тукущей базе данных
'---------------------------------------------------------------------------------------
'
Public Function fnPathDB() As String

    Dim sResult As String
   On Error GoTo fnPathDB_Error

' Использовать в версиях начиная с А2000
    fnPathDB = CurrentProject.Path

' Использовать в версии А97
'    sResult = CurrentDb.Name
'    sResult = fnGetParentFolderName(sResult)
'    fnPathDB = sResult
    
   On Error GoTo 0
Exit_fnPathDB:
   Exit Function

fnPathDB_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnPathDB в Module Module1"
    Resume Exit_fnPathDB

End Function

'---------------------------------------------------------------------------------------
' Procedure : GetTablesRefDAO
' DateTime  : 19.01.2007 10:21
' Author    : DSonnyh
' Purpose   : проверка на наличие линкованных таблиц , метод DAO
'---------------------------------------------------------------------------------------
'
Public Function GetTablesRefDAO() As Variant
'установка связи между таблицами
    Dim dbs As DAO.Database
    Dim tdf As DAO.TableDef
    Dim sss As Variant, k As Integer
    Set dbs = CurrentDb
    For Each tdf In dbs.TableDefs
       sss = tdf.Connect
        If Len(sss) > 0 Then
            k = InStr(1, sss, "=")
            If k = 0 Then
                GetTablesRefDAO = ""
                Exit Function
            End If
            GetTablesRefDAO = Mid(sss, k + 1)
            Exit Function
        End If
    Next tdf
    GetTablesRefDAO = ""
End Function

'---------------------------------------------------------------------------------------
' Procedure : SetReferences
' DateTime  : 22.01.2007 11:14
' Author    : DSonnyh
' Purpose   : Подпрограмма определения привязки баз данных
'---------------------------------------------------------------------------------------
'
Public Sub SetReferences()
' Подпрограмма определения привязки баз данных
' метод ADO
On Error GoTo Err_SetReferences
   Dim rstTemp1 As ADODB.Recordset
   Dim rstTemp2 As ADODB.Recordset
   Dim PName As String
   Dim strBasePath As String, strPB As String
   Dim strSQL As String
   Dim strNameAlt As String, strNameNew As String
   
   Set rstTemp1 = New ADODB.Recordset
   Set rstTemp2 = New ADODB.Recordset
   
' отключение всех подключенных таблиц
   Call ClearTablesRef
   
   strSQL = "SELECT * FROM SystemTables WHERE blnConnect=True ;"
   rstTemp1.Open strSQL, CurrentProject.Connection
   rstTemp2.Open "SystemBases", CurrentProject.Connection, adOpenKeyset, adLockOptimistic
   
   Do Until rstTemp2.EOF
' перебор и проверказадания путей к базам данных
       strBasePath = Nz(rstTemp2!PathBase)
       If Len(strBasePath) = 0 Then
'          MsgBox "Не задан путь к базе данных " + rstTemp2!NickBase, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
        Select Case MsgBox("Не задан путь к базе данных - " & rstTemp2!NickBase _
                           & vbCrLf & "Будете задавать сейчас" _
                           , vbYesNo Or vbExclamation Or vbDefaultButton1, "Подключаемые базы")
        
            Case vbYes

                If GetDBFileNameDlg(0, strBasePath) Then
                     strPB = strBasePath
                     strPB = Left(strPB, InStr(strPB, Chr(0)) - 1)
                     rstTemp2.Fields("PathBase") = strPB
                     rstTemp2.Update
                    Else
'                     bResult = False
                     Exit Do
                End If

            Case vbNo
'                bResult = False
                Exit Do

        End Select
            
        Else
          If Not fnFileExists(strBasePath) Then
'             MsgBox "Не найдена базе данных " + strBasePath, vbOKOnly, "Путь к базе данных"
'! здесь необходимо вставить обработку критической ситуации. Задание пути к базе данных.
            Select Case MsgBox("Не найдена базе данных - " & rstTemp2!NickBase _
                               & vbCrLf & strBasePath _
                               & vbCrLf & "Будете задавать сейчас" _
                               , vbYesNo Or vbExclamation Or vbDefaultButton1, "Подключаемые базы")
            
                Case vbYes
            
                       If GetDBFileNameDlg(0, strBasePath) Then
                            strPB = strBasePath
                            strPB = Left(strPB, InStr(strPB, Chr(0)) - 1)
                            rstTemp2.Fields("PathBase") = strPB
                            rstTemp2.Update
                           Else
'                            bResult = False
                            Exit Do
                       End If
            
                Case vbNo
'                        bResult = False
                        Exit Do
            
            End Select

          End If
       End If
       rstTemp2.MoveNext
   Loop
   rstTemp2.Close
   
   
   With rstTemp1
      Do While Not .EOF
         rstTemp2.Open "SystemBases", CurrentProject.Connection
         Do Until rstTemp2.EOF
            If !BName = rstTemp2!Id Then
               PName = rstTemp2!PathBase
               Exit Do
            End If
            rstTemp2.MoveNext
            
         Loop
         rstTemp2.Close
             
        strNameAlt = !TName
        If Len(Nz(!TNameNew, "")) = 0 Then
           strNameNew = !TName
          Else
           strNameNew = !TNameNew
        End If
        Call SetTableRefBase(strNameAlt, strNameNew, PName)
       
         .MoveNext
      Loop
      .Close
   End With
    
    Set rstTemp1 = Nothing
    Set rstTemp2 = Nothing

Exit_SetReferences:
    Exit Sub

Err_SetReferences:
    MsgBox Err.Description
'!!!
    If Not blnDesignChanges Then
' выход из программы в рабочем режиме
        DoCmd.Quit
    End If
    Resume Exit_SetReferences
End Sub


'---------------------------------------------------------------------------------------
' Procedure : ClearTablesRef
' DateTime  : 19.01.2007 17:16
' Author    : DSonnyh
' Purpose   : отключение подключенных таблиц
'---------------------------------------------------------------------------------------
'
Public Sub ClearTablesRef(Optional intMetod As Variant = 0)
' отключение подключенных таблиц
' intMetod - используемый метод 0 - DAO или 1 - ADO (опционально), по умолчанию - DAO

   On Error GoTo ClearTablesRef_Error

        If intMetod = 0 Then
            Call ClearTablesRefDAO
        Else
            Call ClearTablesRefADO
        End If
    

   On Error GoTo 0
Exit_ClearTablesRef:
   Exit Sub

ClearTablesRef_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре ClearTablesRef в Module modConnect"
    Resume Exit_ClearTablesRef

End Sub
