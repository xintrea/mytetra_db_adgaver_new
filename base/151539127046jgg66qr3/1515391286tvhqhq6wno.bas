Attribute VB_Name = "objFSO"
Option Compare Database
Option Explicit

' Краткая библиотека методов FSO (FileSystemObject)
' Release 1.3 от 22.10.2015
' Автор: Дмитрий Сонных (aka Joss) banderlogi@bk.ru
' Лицензия: freeware - свободное использование. Ссылка необязательна, но желательна.
'
' по мотивам справочника В.И.Короля "Visual Basic 6.0, Visual Basic for Applications 6.0"
'
' Объектная модель FileSystemObject представляет собой неиерархическую структуру
' объектов (классов), позволяющих получить информацию о файловой системе компьютера
' и выполнить различные операции и каталогами этой системы.
' Объектная модель включает следующие классы
' FileSystemObject - обеспечивает доступ к файловой системе компьютера
' Drives - содержит объекты Drive, каждый из которых ассоциируется ровно с одним
'          диском в файловой системе компьютера с учетом сети
' Drive - обеспечивает информацией о заданном диске компьютера
' Folders - семейство Folders содержит объекты Folder, каждый из которых
'           ассоциирован с одним подкаталогом заданного каталога
' Folder - обеспечивает доступ к информации озаданной папке, о содержащихся в неё
'          папкахи каталогах, а также о методах перемещения папки и создании
'          текстового файла.
' Files - семейство Files содержит объекты File, каждый из которых ассоциирован
'         ровно с одним файлом
' File - обеспечивает доступ к информации о заданном файле, методов перемещения
'        и открытия файла
' TextStream - обеспечивает операции чтения/записи для текстового файла, открытого
'              в режиме последовательного доступа
'
' Данная библиотека содержит почти все методы объекта FileSystemObject (кроме тех,
' что устанавливают ссылки на другие объекты).
'
' Примечание: Здесь может возникнуть небольшая путаница: сама модель называется FileSystemObject и
' верхний объект модели называется FileSystemObject.
'
' FileSystemObject (объект) - обеспечивает доступ к файловой системе компьютера. Будучи
' объектом верхнего уровня объектной модели FileSystemObject, является "точкой входа" в
' в файловую систему компьютера. Только после его создания возможен доступ к другим
' объектам модели, их методам и свойствам
'

'---------------------------------------------------------------------------------------
' Procedure : fnBuildPath
' DateTime  : 13.06.2006 14:24
' Author    : DSonnyh
' Purpose   : Создание строки путем слияния аргументов и добавления между ними \ (если его нет)
'---------------------------------------------------------------------------------------
'
Public Function fnBuildPath(strPath As String, strName As String) As String
' strPath - строка, имеющая смысл полного или относительного каталога
' strName - строка, имеющая смысл относительного имени каталога или файла
   On Error GoTo fnBuildPath_Error
    
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")

    fnBuildPath = objFSO.BuildPath(strPath, strName)
    Set objFSO = Nothing

   On Error GoTo 0
   Exit Function

fnBuildPath_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnBuildPath of Module objFSO"
    Set objFSO = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : sbCopyFile
' DateTime  : 22.06.2006 16:54
' Author    : DSonnyh
' Purpose   : копирование одного или нескольких файлов из одной папки в другую
'---------------------------------------------------------------------------------------
'
Public Sub sbCopyFile(strSource As String, strDestination As String, Optional blnOverwriteFiles As Boolean = True)
' копирование одного или нескольких файлов из одной папки в другую
' strSource - путь и имя копируемого файла
' strDestination - путь и необязательное имя файла, в который будет производится копирование
' blnOverwriteFiles - флаг, задающий, будет ли копируемый файл записываться поверх существующего
'   с тем же именем. True (по умолчанию) означает запись файла поверх существующего без предупреждения

   On Error GoTo sbCopyFile_Error


   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")

   objFSO.CopyFile strSource, strDestination, blnOverwriteFiles
   Set objFSO = Nothing

   On Error GoTo 0
   Exit Sub

sbCopyFile_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbCopyFile of Module objFSO"
    Set objFSO = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sbCopyFolder
' DateTime  : 13.06.2006 14:13
' Author    : DSonnyh
' Purpose   : Копирование содержимого папки со всем содержимым в заданное место
'---------------------------------------------------------------------------------------
'
Public Sub sbCopyFolder(strSource As String, strDestination As String, Optional blnOverwriteFiles As Boolean = True)
' strSource - путь и имя копируемой папки
' strDestination - путь, указывающий, куда будет копироваться папка
' blnOverwriteFiles - флаг, задающий, будет ли копируемый файл записываться поверх существующего
'   с тем же именем. True (по умолчанию) означает запись файла поверх существующего без предупреждения
' При установке флага blnOverwriteFiles в False, если папка strSource или что-то и её содержания существует
' в strDestination, генерируется ошибка времени исполнения 58: File already exists
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null
' Если любой из файлов, существующий одновременно и в strSource и в strDestination имеет в последнем атрибут
' "Read only" (Только для чтения), в независимости от установки флага blnOverwriteFiles генерируется ошибка
' времени исполнения 70: Permission Denied
'
   On Error GoTo sbCopyFolder_Error
    
   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.CopyFolder strSource, strDestination, blnOverwriteFiles
   Set objFSO = Nothing

   On Error GoTo 0
   Exit Sub

sbCopyFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbCopyFolder of Module objFSO"
    Set objFSO = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sbCreateFolder
' DateTime  : 13.06.2006 14:44
' Author    : DSonnyh
' Purpose   : создание новой папки с заданным именем
'---------------------------------------------------------------------------------------
'
Public Sub sbCreateFolder(strPath As String)
' strPath - имя создаваемой папки, может быть относительным именем. Если strPath содержит
' только имя папки, она создается в текущей папке на текущем диске
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null
' Если strPath определяет уже существующую папку, генерируется ошибка 58: File already exists

   On Error GoTo sbCreateFolder_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder strPath
    Set objFSO = Nothing

   On Error GoTo 0
   Exit Sub

sbCreateFolder_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure sbCreateFolder of Module objFSO"
    Set objFSO = Nothing

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fnCreateFolder
' Author    : Dmitriy
' Date      : 26.04.2015
' Purpose   :
'---------------------------------------------------------------------------------------
'
Public Function fnCreateFolder(strPath As String) As Boolean
' strPath - имя создаваемой папки, может быть относительным именем. Если strPath содержит
' только имя папки, она создается в текущей папке на текущем диске
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null
' Если strPath определяет уже существующую папку, генерируется ошибка 58: File already exists

    Dim bResult As Boolean


   On Error GoTo fnCreateFolder_Error

    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.CreateFolder strPath
    Set objFSO = Nothing
    bResult = True
    fnCreateFolder = bResult

   On Error GoTo 0
   Exit Function

fnCreateFolder_Error:

'    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnCreateFolder of Module mdl_FSO"
    Set objFSO = Nothing
    Err.Clear

End Function


'---------------------------------------------------------------------------------------
' Procedure : sbDeleteFile
' DateTime  : 22.05.2006 16:45
' Author    : DSonnyh
' Purpose   : удаление одного или нескольких файлов
'---------------------------------------------------------------------------------------
'
Public Sub sbDeleteFile(strFileSpec As String, Optional blnForce As Boolean = False)
' удаление файлов происходит окончательно и бесповоротно - они не попадают в корзину.
' objFSO - ссылка на созданный объект FileSystemObject
' strFileSpec - путь и имя удаляемого файла (файлов). Может быть как абсолютным, так и
'            относительным. Если он опущен, то считается, что удаляемые файлы находятся в
'            текущем каталоге. Может содержать символы (* и !), но только в имени и расширении файла
'            Если файл не существует (или не существует ни одного файла, соответствующего заданному шаблону),
'            генерируется ошибка 53: File not found
' blnForce - (boolean) флаг, задающий, будут ли удаляться файлы с атрибутом Read only (Только для чтения)
'            False (по умолчанию) не позволяет удалять такие файлы
' Если удаляемый файл занят или имеет атрибут "Только для чтения" (Read only) и флаг Force не выставлен в True,
' генерируется ошибка 70: Permission Denied
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo sbDeleteFile_Error
   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
    objFSO.DeleteFile strFileSpec, blnForce
    Set objFSO = Nothing
    

   On Error GoTo 0
Exit_sbDeleteFile:
   Exit Sub

sbDeleteFile_Error:

If Err.Number = 70 Then
'    Set objFSO = Nothing
    Resume Next
Else
    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре sbDeleteFile в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_sbDeleteFile
End If

End Sub

'---------------------------------------------------------------------------------------
' Procedure : sbDeleteFolder
' DateTime  : 02.10.2006 15:08
' Author    : DSonnyh
' Purpose   : удаление одной или нескольких папок вместе со всем содержимым.
'---------------------------------------------------------------------------------------
'
Public Sub sbDeleteFolder(strFolderSpec As String, Optional blnForce As Boolean = False)
' удаление папок происходит окончательно и бесповоротно - они не попадают в корзину.
' objFSO - ссылка на созданный объект FileSystemObject
' strFolderSpec - путь и имя удаляеой папки (папок). Может быть как абсолютным, так и
'            относительным. Если он опущен, то считается, что удаляемые папки находятся в текущем
'            каталоге текущего диска. Может содержать символы (* и !), но только в последней части
'            Если папка не существует (или не существует ни одной папки, соответствующей заданному шаблону),
'            генерируется ошибка 76: Path not found
' blnForce - (boolean) флаг, задающий, будут ли удаляться файлы с атрибутом Read only (Только для чтения)
'            False (по умолчанию) не позволяет удалять такие файлы
' Если удаляемая папка или что-то из её содержимого имеет атрибут "Только для чтения" (Read only) и флаг
' Force не выставлен в True, генерируется ошибка 70: Permission Denied
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo sbDeleteFolder_Error

    
   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.DeleteFolder strFolderSpec, blnForce
   Set objFSO = Nothing
    

   On Error GoTo 0
Exit_sbDeleteFolder:
   Exit Sub

sbDeleteFolder_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре sbDeleteFolder в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_sbDeleteFolder

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fnDriveExists
' DateTime  : 02.10.2006 16:20
' Author    : DSonnyh
' Purpose   : Проверка существования диска с указанным именем на локальной машине
'---------------------------------------------------------------------------------------
'
Public Function fnDriveExists(strDriveSpec As String) As Boolean
' Возвращает True - если диск существует
' objFSO - ссылка на созданный объект FileSystemObject
' strDrive - Имя проверяемого диска. Для буквенного обозначения диска strDriveSpec двоеточие
' после буквы не обязательно
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnDriveExists_Error
    
   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
    fnDriveExists = objFSO.DriveExists(strDriveSpec)
    Set objFSO = Nothing

   On Error GoTo 0
Exit_fnDriveExists:
   Exit Function

fnDriveExists_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnDriveExists в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnDriveExists


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnFileExists
' DateTime  : 17.08.2006 13:06
' Author    : DSonnyh
' Purpose   : проверка существования файла на локальной машине или в сети
'---------------------------------------------------------------------------------------
'
Public Function fnFileExists(strFileName As String) As Boolean
' objFSO - ссылка на созданный объект FileSystemObject
' strFileSpec - путь и имя проверяемого файла. Может быть как абсолютным, так и относительным.
'           Не может содержать символы шаблонов. Если путь опущен, файл ищется в текущем каталоге
'           текущего диска.
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnFileExists_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnFileExists = objFSO.FileExists(strFileName)
   Set objFSO = Nothing

   On Error GoTo 0
Exit_fnFileExists:
   Exit Function

fnFileExists_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnFileExists в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnFileExists

End Function
'---------------------------------------------------------------------------------------
' Procedure : fnFolderExists
' DateTime  : 17.08.2006 13:06
' Author    : DSonnyh
' Purpose   : Проверка существования папки с заданным именем на локальной машине или в сети
'---------------------------------------------------------------------------------------
'
Public Function fnFolderExists(strFolderName As String) As Boolean
' objFSO - ссылка на созданный объект FileSystemObject
' strFolderSpec - путь и имя проверяемого файла. Может быть как абсолютным, так и относительным.
'           Не может содержать символы шаблонов. Если путь опущен, папка ищется в текущем каталоге
'           текущего диска.
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnFolderExists_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnFolderExists = objFSO.FolderExists(strFolderName)
   Set objFSO = Nothing

   On Error GoTo 0
Exit_fnFolderExists:
   Exit Function

fnFolderExists_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnFolderExists в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnFolderExists


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetAbsolutePathName
' DateTime  : 02.10.2006 16:08
' Author    : DSonnyh
' Purpose   : Получение полного имени файла или папки по его относительному имени
'---------------------------------------------------------------------------------------
'
Public Function fnGetAbsolutePathName(strPath As String) As String
' objFSO - ссылка на созданный объект FileSystemObject
' strPath - строка, имеющая смысл полного или относительного имени файла или папки
' Символы шаблона могут включаться в любую часть Path
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnGetAbsolutePathName_Error
    
   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnGetAbsolutePathName = objFSO.GetAbsolutePathName(strPath)
   Set objFSO = Nothing

   On Error GoTo 0
Exit_fnGetAbsolutePathName:
   Exit Function

fnGetAbsolutePathName_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnGetAbsolutePathName в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnGetAbsolutePathName


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetBaseName
' DateTime  : 24.05.2006 12:36
' Author    : DSonnyh
' Purpose   : Получение последнего компонента - имени папки или файла (без расширения)
'           : по его полному или относительному имени
'---------------------------------------------------------------------------------------
'
Public Function fnGetBaseName(strPath As String) As String
' objFSO - ссылка на созданный объект FileSystemObject
' strPath - строка, имеющая смысл полного или относительного имени файла или папки
' Символы шаблона могут включаться в любую часть Path
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnGetBaseName_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnGetBaseName = objFSO.GetBaseName(strPath)
   Set objFSO = Nothing

   On Error GoTo 0
Exit_fnGetBaseName:
   Exit Function

fnGetBaseName_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnGetBaseName в Module objFSOx"
    Set objFSO = Nothing
    Resume Exit_fnGetBaseName


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetDriveName
' DateTime  : 02.10.2006 16:43
' Author    : DSonnyh
' Purpose   : Получение имени диска из имени папки или файла
'---------------------------------------------------------------------------------------
'
Public Function fnGetDriveName(strPath As String) As String
' objFSO - ссылка на созданный объект FileSystemObject
' strPath - строка, имеющая смысл пути (имени файла или папки)
' Если по заданному имени нельзя определить имя диска, метод возвращает пустую строку
' По относительному пути определить имя диска нельзя
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnGetDriveName_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnGetDriveName = objFSO.GetDriveName(strPath)
   Set objFSO = Nothing

   On Error GoTo 0
Exit_fnGetDriveName:
   Exit Function

fnGetDriveName_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnGetDriveName в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnGetDriveName


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetExtensionName
' DateTime  : 06.10.2006 14:20
' Author    : DSonnyh
' Purpose   : Получение расширения из заданного имени файла
'---------------------------------------------------------------------------------------
'
Public Function fnGetExtensionName(strPath As String) As String
' objFSO - ссылка на созданный объект FileSystemObject
' strPath - строка, имеющая смысл полного или относительного пути имени файла
' Если расширение не обнаружено, возвращается пустая строка
' Передача в качестве аргумента Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnGetExtensionName_Error
    
   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnGetExtensionName = objFSO.GetExtensionName(strPath)
   Set objFSO = Nothing


   On Error GoTo 0
Exit_fnGetExtensionName:
   Exit Function

fnGetExtensionName_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnGetExtensionName в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnGetExtensionName


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetFileName
' DateTime  : 06.10.2006 14:25
' Author    : DSonnyh
' Purpose   : Получение имени (с расширением) из полного имени (пути) файла
'---------------------------------------------------------------------------------------
'
Public Function fnGetFileName(strPath As String) As String
' objFSO - ссылка на созданный объект FileSystemObject
' strPath - строка, имеющая смысл полного или относительного пути имени файла
' Если по заданному имени невозможно определить имя файла, возвращается пустая строка
' Передача в качестве аргумента Null генерирует ошибку 94: Invalid Use of Null
   
   On Error GoTo fnGetFileName_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
    fnGetFileName = objFSO.GetFileName(strPath)
    Set objFSO = Nothing

   On Error GoTo 0
Exit_fnGetFileName:
   Exit Function

fnGetFileName_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnGetFileName в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnGetFileName


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetParentFolderName
' DateTime  : 24.05.2006 12:21
' Author    : DSonnyh
' Purpose   : получение имени папки, являющейся предпоследним компонентом полного имени
'---------------------------------------------------------------------------------------
'
Public Function fnGetParentFolderName(strFileName As String) As String
' objFSO - ссылка на созданный объект FileSystemObject
' strPath - строка, имеющая смысл полного или относительного пути имени файла
' получение имени папки, являющейся предпоследним компонентом
' полного имени (пути) файла или папки
' Если по заданному имени невозможно определить имя папки, возвращается пустая строка
' Передача в качестве аргумента Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo fnGetParentFolderName_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnGetParentFolderName = objFSO.GetParentFolderName(strFileName)
   Set objFSO = Nothing

   On Error GoTo 0
   Exit Function

fnGetParentFolderName_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnGetParentFolderName of Module objFSO"
    Set objFSO = Nothing

End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetTempName
' DateTime  : 10.10.2006 16:57
' Author    : DSonnyh
' Purpose   : Получение имени для временного файла
'---------------------------------------------------------------------------------------
'
Public Function fnGetTempName() As String
' метод не создает файл, он только придумывает имя.
' для временной папки можно использовать myTmpFolderName = objFSO.GetBaseName(objFSO.GetTempName)
    
   On Error GoTo fnGetTempName_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   fnGetTempName = objFSO.GetTempName
   Set objFSO = Nothing

   On Error GoTo 0
Exit_fnGetTempName:
   Exit Function

fnGetTempName_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре fnGetTempName в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_fnGetTempName


End Function


'---------------------------------------------------------------------------------------
' Procedure : sbMoveFile
' DateTime  : 10.10.2006 17:05
' Author    : DSonnyh
' Purpose   : Перемещение одного или нескольких файлов из одной папки в другую
'---------------------------------------------------------------------------------------
'
Public Sub sbMoveFile(strSource As String, strDestination As String)
' objFSO - ссылка на созданный объект FileSystemObject
' strSource - путь и имя перемещаемого файла
' strDestination - путь, определяющий, куда будет производится перемещение
' strDestination не может содержать символы шаблонов
' strSource может содержать символы шаблонов, но только в имени файла
' если файл strSource не существует, генерируется ошибка времени исполнения 53: File not found
' Если strDestination уже существует, генерируется ошибка времени исполнения 58: File already exists
' Если strDestination не содержит разделитель \ в качестве последнего символа, в strSource нет символов шаблона,
' генерируется ошибка времени исполнения 70: Permission Denied
' Если файл strDestination имеет атрибут "Только для чтения" (Real only), генерируется ошибка времени
' исполнения 70: Permission Denied
' Передача в качестве аргумента Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo sbMoveFile_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.MoveFile strSource, strDestination
   Set objFSO = Nothing

   On Error GoTo 0
Exit_sbMoveFile:
   Exit Sub

sbMoveFile_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре sbMoveFile в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_sbMoveFile

End Sub


'---------------------------------------------------------------------------------------
' Procedure : sbMoveFolder
' DateTime  : 10.10.2006 17:25
' Author    : DSonnyh
' Purpose   : Перемещение папки со всеми содержащимися в ней папками в заданное место
'---------------------------------------------------------------------------------------
'
Public Sub sbMoveFolder(strSource As String, strDestination As String)
' objFSO - ссылка на созданный объект FileSystemObject
' strSource - путь и имя перемещаемой папки
' strDestination - путь, определяющий, куда будет производится перемещение
' strDestination не может содержать символы шаблонов
' strSource может содержать символы шаблонов, но только в имени файла
' если strSource не существует, генерируется ошибка времени исполнения 76: Path not found
' Если любой из файлов, существующий одновременно и в strSource и в strDestination имеет в последнем атрибут
' "Read only" (Только для чтения),генерируется ошибка времени исполнения 70: Permission Denied
' Передача в качестве любого из аргументов Null генерирует ошибку 94: Invalid Use of Null

   On Error GoTo sbMoveFolder_Error

   Dim objFSO As Object
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.MoveFolder strSource, strDestination
   Set objFSO = Nothing

   On Error GoTo 0
Exit_sbMoveFolder:
   Exit Sub

sbMoveFolder_Error:

    MsgBox "Ошибка " & Err.Number & " (" & Err.Description & ") в процедуре sbMoveFolder в Module objFSO"
    Set objFSO = Nothing
    Resume Exit_sbMoveFolder

End Sub

'---------------------------------------------------------------------------------------
' Procedure : fnGetSpecialFolder
' DateTime  : 22.10.2015 13:06
' Author    : DSonnyh
' Purpose   : Полечение пути к системным папкам
'---------------------------------------------------------------------------------------
'
Public Function fnGetSpecialFolder(intConst As Integer) As String
' Назначение. Получение ссылки на объект Folder, связанный с одной из трех специальных папок
' - папки Windows, системной папки и папки временных файлов.
' Возвращает Объектную ссылку на объект типа Folder.
' objFSO - ссылка на созданный объект FileSystemObject
' Имя константы   Значение    Описание
' WindowsFolder   0   Папка Windows, содержащая файлы, установленные операционной системой Windows
' SystemFolder    1   Системная папка, содержащая библиотеки, шрифты и драйверы устройств.
' TemporaryFolder 2   Папка системы для хранения временных файлов.
'                     Устанавливается переменной среды TMP.

    Dim sResult As String
    On Error GoTo fnGetSpecialFolder_Error
   
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sResult = objFSO.GetSpecialFolder(intConst)
    Set objFSO = Nothing

    fnGetSpecialFolder = sResult

Exit_fnGetSpecialFolder:

    On Error GoTo 0
    Exit Function

fnGetSpecialFolder_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnGetSpecialFolder of Module mdl_FSO"
     Resume Exit_fnGetSpecialFolder


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetWindowsFolder
' DateTime  : 22.10.2015 15:43
' Author    : DSonnyh
' Purpose   : Получение ссылки на объект Folder, связанный папкой Windows.
'---------------------------------------------------------------------------------------
'
' Назначение. Получение ссылки на объект Folder, связанный папкой Windows.
' Возвращает Объектную ссылку на объект типа Folder.
' objFSO - ссылка на созданный объект FileSystemObject
Public Function fnGetWindowsFolder() As String

    Dim sResult As String

    On Error GoTo fnGetWindowsFolder_Error
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sResult = objFSO.GetSpecialFolder(0)
    Set objFSO = Nothing
    fnGetWindowsFolder = sResult

Exit_fnGetWindowsFolder:

    On Error GoTo 0
    Exit Function

fnGetWindowsFolder_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnGetWindowsFolder of Module mdl_FSO"
     Resume Exit_fnGetWindowsFolder


End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetSystemFolder
' DateTime  : 22.10.2015 15:44
' Author    : DSonnyh
' Purpose   : Получение ссылки на объект Folder, связанный папкой System.
'---------------------------------------------------------------------------------------
'
' Назначение. Получение ссылки на объект Folder, связанный папкой System.
' Возвращает Объектную ссылку на объект типа Folder.
' objFSO - ссылка на созданный объект FileSystemObject
Public Function fnGetSystemFolder() As String

    Dim sResult As String

    On Error GoTo fnGetSystemFolder_Error
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sResult = objFSO.GetSpecialFolder(1)
    Set objFSO = Nothing
    fnGetSystemFolder = sResult

Exit_fnGetSystemFolder:

    On Error GoTo 0
    Exit Function

fnGetSystemFolder_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnGetSystemFolder of Module mdl_FSO"
     Resume Exit_fnGetSystemFolder

End Function

'---------------------------------------------------------------------------------------
' Procedure : fnGetTemporaryFolder
' DateTime  : 22.10.2015 15:44
' Author    : DSonnyh
' Purpose   : Получение ссылки на объект Folder, связанный папкой временных файлов
'---------------------------------------------------------------------------------------
'
' Назначение. Получение ссылки на объект Folder, связанный папкой временных файлов
' Возвращает Объектную ссылку на объект типа Folder.
' objFSO - ссылка на созданный объект FileSystemObject
Public Function fnGetTemporaryFolder() As String

    Dim sResult As String


    On Error GoTo fnGetTemporaryFolder_Error
    Dim objFSO As Object
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    sResult = objFSO.GetSpecialFolder(2)
    Set objFSO = Nothing
    fnGetTemporaryFolder = sResult

Exit_fnGetTemporaryFolder:

    On Error GoTo 0
    Exit Function

fnGetTemporaryFolder_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure fnGetTemporaryFolder of Module mdl_FSO"
     Resume Exit_fnGetTemporaryFolder

End Function
