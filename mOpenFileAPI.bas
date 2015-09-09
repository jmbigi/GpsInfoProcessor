Attribute VB_Name = "mOpenFileAPI"
Option Compare Database
Option Explicit

Type tagOPENFILENAME
  lStructSize As Long
  hWndOwner As Long
  hInstance As Long
  strFilter As String
  strCustomFilter As String
  nMaxCustFilter As Long
  nFilterIndex As Long
  strFile As String
  nMaxFile As Long
  strFileTitle As String
  nMaxFileTitle As Long
  strInitialDir As String
  strTitle As String
  flags As Long
  nFileOffset As Integer
  nFileExtension As Integer
  strDefExt As String
  lCustData As Long
  lpfnHook As Long
  lpTemplateName As String
End Type
Declare Function aht_apiGetOpenFileName Lib "comdlg32.dll" _
    Alias "GetOpenFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function aht_apiGetSaveFileName Lib "comdlg32.dll" _
    Alias "GetSaveFileNameA" (OFN As tagOPENFILENAME) As Boolean
Declare Function CommDlgExtendedError Lib "comdlg32.dll" () As Long
Global Const ahtOFN_READONLY = &H1
Global Const ahtOFN_OVERWRITEPROMPT = &H2
Global Const ahtOFN_HIDEREADONLY = &H4
Global Const ahtOFN_NOCHANGEDIR = &H8
Global Const ahtOFN_SHOWHELP = &H10
Global Const ahtOFN_NOVALIDATE = &H100
Global Const ahtOFN_ALLOWMULTISELECT = &H200
Global Const ahtOFN_EXTENSIONDIFFERENT = &H400
Global Const ahtOFN_PATHMUSTEXIST = &H800
Global Const ahtOFN_FILEMUSTEXIST = &H1000
Global Const ahtOFN_CREATEPROMPT = &H2000
Global Const ahtOFN_SHAREAWARE = &H4000
Global Const ahtOFN_NOREADONLYRETURN = &H8000
Global Const ahtOFN_NOTESTFILECREATE = &H10000
Global Const ahtOFN_NONETWORKBUTTON = &H20000
Global Const ahtOFN_NOLONGNAMES = &H40000
Global Const ahtOFN_EXPLORER = &H80000
Global Const ahtOFN_NODEREFERENCELINKS = &H100000
Global Const ahtOFN_LONGNAMES = &H200000

Global Const ahtOFS_FILE_OPEN_FLAGS = ahtOFN_EXPLORER _
    Or ahtOFN_LONGNAMES _
    Or ahtOFN_NODEREFERENCELINKS

Global Const ahtOFS_FILE_SAVE_FLAGS = ahtOFN_EXPLORER _
    Or ahtOFN_LONGNAMES

Function TestIt()

  Dim strFilter As String
  Dim lngFlags As Long
  strFilter = ahtAddFilterItem(strFilter, "Access Files (*.mda, *.mdb)", _
      "*.MDA;*.MDB")
  strFilter = ahtAddFilterItem(strFilter, "dBASE Files (*.dbf)", "*.DBF")
  strFilter = ahtAddFilterItem(strFilter, "Text Files (*.txt)", "*.TXT")
  strFilter = ahtAddFilterItem(strFilter, "All Files (*.*)", "*.*")
  MsgBox "You selected: " & ahtCommonFileOpenSave(InitialDir:="C:\", _
      Filter:=strFilter, FilterIndex:=3, flags:=lngFlags, _
      DialogTitle:="Hello! Open Me!")
End Function

Function GetOpenFile(Optional varDirectory As Variant, _
    Optional varTitleForDialog As Variant) As Variant
  Dim strFilter As String
  Dim lngFlags As Long
  Dim varFileName As Variant
  lngFlags = ahtOFN_FILEMUSTEXIST Or _
      ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
  If IsMissing(varDirectory) Then
    varDirectory = ""
  End If
  If IsMissing(varTitleForDialog) Then
    varTitleForDialog = ""
  End If

  strFilter = ahtAddFilterItem(strFilter, _
      "Access (*.mdb)", "*.MDB;*.MDA")
  varFileName = ahtCommonFileOpenSave( _
      OpenFile:=True, _
      InitialDir:=varDirectory, _
      Filter:=strFilter, _
      flags:=lngFlags, _
      DialogTitle:=varTitleForDialog)
  If Not IsNull(varFileName) Then
    varFileName = TrimNull(varFileName)
  End If
  GetOpenFile = varFileName
End Function

Function ahtCommonFileOpenSave( _
    Optional ByRef flags As Variant, _
    Optional ByVal InitialDir As Variant, _
    Optional ByVal Filter As Variant, _
    Optional ByVal FilterIndex As Variant, _
    Optional ByVal DefaultExt As Variant, _
    Optional ByVal FileName As Variant, _
    Optional ByVal DialogTitle As Variant, _
    Optional ByVal hwnd As Variant, _
    Optional ByVal OpenFile As Variant) As Variant
  Dim OFN As tagOPENFILENAME
  Dim strFileName As String
  Dim strFileTitle As String
  Dim fResult As Boolean
  If IsMissing(InitialDir) Then InitialDir = CurDir
  If IsMissing(Filter) Then Filter = ""
  If IsMissing(FilterIndex) Then FilterIndex = 1
  If IsMissing(flags) Then flags = 0&
  If IsMissing(DefaultExt) Then DefaultExt = ""
  If IsMissing(FileName) Then FileName = ""
  If IsMissing(DialogTitle) Then DialogTitle = ""
  If IsMissing(hwnd) Then hwnd = Application.hWndAccessApp
  If IsMissing(OpenFile) Then OpenFile = True
  strFileName = Left(FileName & String(25600, 0), 25600)
  strFileTitle = String(25600, 0)
  With OFN
    .lStructSize = Len(OFN)
    .hWndOwner = hwnd
    .strFilter = Filter
    .nFilterIndex = FilterIndex
    .strFile = strFileName
    .nMaxFile = Len(strFileName)
    .strFileTitle = strFileTitle
    .nMaxFileTitle = Len(strFileTitle)
    .strTitle = DialogTitle
    .flags = flags Or (IIf(OpenFile, ahtOFS_FILE_OPEN_FLAGS, ahtOFS_FILE_SAVE_FLAGS))
    .strDefExt = DefaultExt
    .strInitialDir = InitialDir
    .hInstance = 0
    .strCustomFilter = ""
    .nMaxCustFilter = 0
    .lpfnHook = 0
    .strCustomFilter = String(9999, 0)
    .nMaxCustFilter = 9999
  End With
  If OpenFile Then
    fResult = aht_apiGetOpenFileName(OFN)
  Else
    fResult = aht_apiGetSaveFileName(OFN)
  End If

  If fResult Then
    If Not IsMissing(flags) Then flags = OFN.flags
    ahtCommonFileOpenSave = TrimToLastNull(OFN.strFile)
  Else
    ahtCommonFileOpenSave = vbNullString
  End If
End Function
Function ahtAddFilterItem(strFilter As String, _
    strDescription As String, Optional varItem As Variant) As String

  If IsMissing(varItem) Then varItem = "*.*"
  ahtAddFilterItem = strFilter & _
      strDescription & vbNullChar & _
      varItem & vbNullChar
End Function
Private Function TrimNull(ByVal strItem As String) As String
  Dim intPos As Integer
  intPos = InStr(strItem, vbNullChar)
  If intPos > 0 Then
    TrimNull = Left(strItem, intPos - 1)
  Else
    TrimNull = strItem
  End If
End Function

Private Function TrimToLastNull(ByVal strItem As String) As String
  Dim intPos As Integer
  intPos = InStrRev(strItem, vbNullChar)
  If intPos > 0 Then
    Do While intPos > 0
      If Mid(strItem, intPos, 1) = vbNullChar Then
        intPos = intPos - 1
      Else
        Exit Do
      End If
    Loop
    If intPos = 0 Then
      TrimToLastNull = ""
    Else
      TrimToLastNull = Left(strItem, intPos)
    End If
  Else
    TrimToLastNull = strItem
  End If
End Function

Sub Just_More_Notes()
  Dim strFilter As String
  Dim strInputFileName As String

  strFilter = ahtAddFilterItem(strFilter, "Excel Files (*.XLS)", "*.XLS")
  strInputFileName = ahtCommonFileOpenSave(Filter:=strFilter, OpenFile:=True, _
      DialogTitle:="Please select an input file...", _
      flags:=ahtOFN_HIDEREADONLY Or ahtOFN_LONGNAMES)
End Sub

Function OpenTextFile(Optional varDirectory As Variant) As Variant
  Dim strFilter As String, lngFlags As Long, varFileName As Variant

  lngFlags = ahtOFN_FILEMUSTEXIST Or _
      ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR
  If IsMissing(varDirectory) Then
    varDirectory = ""
  End If
  strFilter = ahtAddFilterItem(strFilter, "Text files (*.txt)", "*.txt")
  strFilter = ahtAddFilterItem(strFilter, "Any files (*.*)", "*.*")
  varFileName = ahtCommonFileOpenSave( _
      OpenFile:=True, _
      InitialDir:=varDirectory, _
      Filter:=strFilter, _
      flags:=lngFlags, _
      DialogTitle:="Open a text file ...")
  If Not IsNull(varFileName) Then
    varFileName = TrimNull(varFileName)
  End If
  OpenTextFile = varFileName
End Function

Function SaveCsvFile(Optional varDirectory As Variant) As Variant
  Dim strFilter As String, lngFlags As Long, varFileName As Variant
  lngFlags = ahtOFN_OVERWRITEPROMPT Or _
      ahtOFN_HIDEREADONLY Or ahtOFN_NOCHANGEDIR

  If IsMissing(varDirectory) Then
    varDirectory = ""
  End If
  strFilter = ahtAddFilterItem(strFilter, "Csv files (*.csv)", "*.csv")
  strFilter = ahtAddFilterItem(strFilter, "Any files (*.*)", "*.*")
  varFileName = ahtCommonFileOpenSave( _
      OpenFile:=True, _
      InitialDir:=varDirectory, _
      Filter:=strFilter, _
      flags:=lngFlags, _
      DialogTitle:="Save a csv file ...")
  If Not IsNull(varFileName) Then
    varFileName = TrimNull(varFileName)
  End If
  OpenTextFile = varFileName
End Function

Public Function StripDelimitedItem(startStrg As String, _
    delimiter As String) As String
  Dim pos As Long
  Dim item As String
  '
  pos = InStr(1, startStrg, delimiter)
  If pos > 0 Then
    StripDelimitedItem = Mid$(startStrg, 1, pos - 1)
    startStrg = Mid$(startStrg, pos + 1, Len(startStrg))
  Else
    StripDelimitedItem = startStrg
    startStrg = ""
  End If
  '
End Function


