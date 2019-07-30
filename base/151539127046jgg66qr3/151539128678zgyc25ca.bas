Attribute VB_Name = "Module1"
Option Compare Database
Option Explicit
Dim DATADbName As Variant
Type OpenFileName
        lStructSize As Long
        hwndOwner As Long
        hInstance As Long
        lpstrFilter As String
        lpstrCustomFilter As String
        nMaxCustFilter As Long
        nFilterIndex As Long
        lpstrFile As String
        nMaxFile As Long
        lpstrFileTitle As String
        nMaxFileTitle As Long
        lpstrInitialDir As String
        lpstrTitle As String
        flags As Long
        nFileOffset As Integer
        nFileExtension As Integer
        lpstrDefExt As String
        lCustData As Long
        lpfnHook As Long
        lpTemplateName As String
End Type

Public Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" _
                (ByVal hwnd As Long, _
                ByVal lpOperation As String, _
                ByVal lpFile As String, _
                ByVal lpParameters As String, _
                ByVal lpDirectory As String, _
                ByVal nShowCmd As Long) As Long

Public Const SW_SHOWMAXIMIZED = 3
Public Const SW_SHOWDEFAULT = 10
Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOWNORMAL = 1

Public Const blnDesignChanges As Boolean = True 'False 'True

Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OpenFileName) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public lngNewId As Long

'---------------------------------------------------------------------------------------
' Procedure : GetDbPath
' DateTime  : 19.01.2007 10:23
' Author    : DSonnyh
' Purpose   : Задание пути к подключаемой базе данных
'---------------------------------------------------------------------------------------
'
Public Function GetDbPath(PName As String) As Integer
    Dim i As Long
    i = MsgBox("Неправильно задана или не задана ссылка на базу данных с таблицами!" _
    & Chr(13) & "Будете задавать ссылку сейчас?", vbOKCancel + vbExclamation)
    If i <> vbOK Then
      DoCmd.Quit
      Exit Function
    End If
    i = GetDBFileNameDlg(0, PName)
    If i = 0 Then
      GetDbPath = False
      DoCmd.Quit
      Exit Function
    End If
    
    GetDbPath = True
End Function
'---------------------------------------------------------------------------------------
' Procedure : GetDBFileNameDlg
' DateTime  : 19.01.2007 10:25
' Author    : DSonnyh
' Purpose   : Диалог задания пути к подключаемой базе MDB
'---------------------------------------------------------------------------------------
'
Public Function GetDBFileNameDlg(hWn As Long, fName As String) As Integer

On Error GoTo Err_

Dim l As Long
Dim pOpenfilename As OpenFileName
Dim strPatchFi As String
Dim strFolderName As String

'   fName = ""
   
   GetDBFileNameDlg = False
   If Right(fName, 1) = "\" Then
        strPatchFi = ""
        strFolderName = fName
       
    Else
        strPatchFi = fnGetFileName(fName)
        strFolderName = fnGetParentFolderName(fName)
   
   End If
   
pOpenfilename.lStructSize = Len(pOpenfilename)
pOpenfilename.lpstrFilter = "*.mdb" + Chr(0) + "*.mdb" + Chr(0) + "*.*" + Chr(0) + "*.*" + Chr(0) + Chr(0)
'pOpenfilename.lpstrFile = String(255, Chr(0))
pOpenfilename.lpstrFile = strPatchFi + String(255 - Len(strPatchFi), Chr(0)) '""
pOpenfilename.nMaxFile = 255
pOpenfilename.lpstrTitle = "Выберите файл с таблицами базы данных"
pOpenfilename.hwndOwner = hWn
pOpenfilename.lpstrInitialDir = strFolderName 'fnGetParentFolderName(fName) '"C:\"

l = GetOpenFileName(pOpenfilename)
If l = 1 Then
   fName = pOpenfilename.lpstrFile
   GetDBFileNameDlg = True
End If

Exit_:
   Exit Function

Err_:
    MsgBox Err.Description
    Resume Exit_

End Function
'---------------------------------------------------------------------------------------
' Procedure : IsFormOpen
' DateTime  : 15.08.2006 16:43
' Author    : DSonnyh
' Purpose   : Проверка, является ли форма открытой
'---------------------------------------------------------------------------------------
'
Public Function IsFormOpen(Name As Variant) As Integer
Dim f As Form
For Each f In Forms
    If f.Name = Name Then
        IsFormOpen = True
        Exit Function
    End If
Next
IsFormOpen = False
End Function

'---------------------------------------------------------------------------------------
' Procedure : IsReportOpen
' DateTime  : 19.01.2007 10:26
' Author    : DSonnyh
' Purpose   : Проверка на открытие указанного отчета
'---------------------------------------------------------------------------------------
'
Public Function IsReportOpen(Name As Variant) As Integer
' проверка на текущее открытие указанного отчета
Dim rep As Access.Report
For Each rep In Reports
    If rep.Name = Name Then
        IsReportOpen = True
        Exit Function
    End If
Next
IsReportOpen = False
End Function
'---------------------------------------------------------------------------------------
' Procedure : subBookmark
' DateTime  : 03.11.2015 15:05
' Author    : DSonnyh
' Purpose   : Работа с завладками ленточной формы
'---------------------------------------------------------------------------------------
'
Public Sub subBookmark(objForm As Form, ctrlId As String, fldId As String, _
                    Optional varNew As Variant = Null, Optional blnDel As Boolean = 0)
Dim lngID As Variant

    On Error GoTo subBookmark_Error

    objForm.Painting = False
    With objForm.Recordset
        Select Case True
        Case blnDel     'Если происходит удаление
            .MoveNext
            If Not .EOF = True Then lngID = objForm.Controls(ctrlId) Else Err.Clear: .MovePrevious: .MovePrevious
            If Not .BOF = True Then lngID = objForm.Controls(ctrlId) Else Err.Clear: lngID = 0
        Case Not IsNull(varNew)     'если вводим новую строку
            lngID = varNew
        Case Else       'другие если
            lngID = objForm.Controls(ctrlId)
        End Select
        
        objForm.RecordSource = objForm.RecordSource
        
        If TypeOf objForm.Recordset Is DAO.Recordset Then
            objForm.Recordset.FindFirst fldId & "=" & lngID
        ElseIf TypeOf objForm.Recordset Is ADODB.Recordset Then
            objForm.Recordset.Find fldId & "=" & lngID
        End If
    End With
    objForm.Painting = True

Exit_subBookmark:

    On Error GoTo 0
    Exit Sub

subBookmark_Error:

     MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure subBookmark of Module Module1"
     Resume Exit_subBookmark

End Sub


