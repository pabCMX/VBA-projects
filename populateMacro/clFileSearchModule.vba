VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clFileSearchModule"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' clFileSearch -- class module to replace Office FileSearch Object
' Version 1.0a
' Created May 24, 2009, Last Updated 11/16/2009
' Created by David W. Fenton
' David Fenton Associates
'   http://dfenton.com/DFA/
' Current version of this class always available at:
'   http://dfenton.com/DFA/download/Access/FileSearch/
' in order to search document properties, requires registered DLL:
'   DSO OLE Document Propoerties Reader 2.1
'   dsofile.dll
'   http://support.microsoft.com/kb/224351
' to search text/files without registering dsofile.dll, use
'   the TextOnly property instead of TextOrProperty
'
' FileSearch Object Properties and methods
' PROPERTIES
' Application                : irrelevant
' Creator (r/w)              : implemented
' ExecuteCancel (r/w)        : added by DWF
' FileName (r/w)             : implemented
' FilesToCheck               : added by DWF
' FileType (r/w)             : implemented
' FileTypes                  : implemented
' FileTypeSpecify (w)        : added by DWF
' FoundFiles                 : implemented
' FSO                        : added by DWF
' LastModified (r/w)         : implemented
' LastModifiedSpecify (w)    : added by DWF
' LastModifiedSpecifyEnd     : added by DWF
' LastModifiedSpecifyStart   : added by DWF
' LookIn (r/w)               : implemented
' MatchAllWordForms          : not implemented
' MatchTextExactly (r/w)     : implemented
' MatchTextType (r/w)        : added by DWF
' ProgressBarForm (w)        : added by DWF
' PropertyTests              : not implemented
' SearchFolders              : not implemented
' SearchPropertiesOnly (r/w) : added by DWF
' SearchScopes               : not implemented
' SearchSubFolders (r/w)     : implemented
' TextOnly (r/w)             : added by DWF
' TextOrProperty (r/w)       : implemented
'
' METHODS
' Execute                    : implemented -- extended by DWF
' NewSearch                  : implemented
' RefreshScopes              : not implemented

'Option Compare Database
Option Explicit
 
' VARIABLES FOR PUBLIC PROPERTIES/COLLECTIONS
Private strCreator As String
Private strFileName As String
Private intFileType As Integer
Private mcolFileTypes As Collection
Private mcolFoundFiles As Collection
Private intLastModified As Integer
Private dteStartDate As Date
Private dteEndDate As Date
Private strLookIn As String
Private bolMatchTextExactly As Boolean
Private intMatchTextType As Integer
Private bolSearchPropertiesOnly As Boolean
Private bolSearchSubFolders As Boolean
Private bolTextOnly As Boolean
Private strTextOrProperty As String
Private bolCancelExecution As Boolean
'Private frmProgressBar As Form
 
' VARIABLES FOR PRIVATE PROPERTIES/COLLECTIONS
Private mcolFilesToCheck As Collection
Private objDSOFileDoc As Object
Private objFSO As Object

' ENUMERATIONS
Public Enum MsoLastModified
  msoLastModifiedAnyTime = 7
  msoLastModifiedLastMonth = 5
  msoLastModifiedLastWeek = 3
  msoLastModifiedThisMonth = 6
  msoLastModifiedThisWeek = 4
  msoLastModifiedToday = 2
  msoLastModifiedYesterday = 1
  msoLastModifiedSpecify = 8 ' DWF-Added Constant
End Enum
Public Enum MsoFileType
  msoFileTypeAllFiles = 1
  msoFileTypeBinders = 6
  msoFileTypeCalendarItem = 11
  msoFileTypeContactItem = 12
  msoFileTypeDatabases = 7
  msoFileTypeDataConnectionFiles = 17
  msoFileTypeDesignerFiles = 22
  msoFileTypeDocumentImagingFiles = 20
  msoFileTypeExcelWorkbooks = 4
  msoFileTypeJournalItem = 14
  msoFileTypeMailItem = 10
  msoFileTypeNoteItem = 13
  msoFileTypeOfficeFiles = 2
  msoFileTypeOutlookItems = 9
  msoFileTypePhotoDrawFiles = 16
  msoFileTypePowerPointPresentations = 5
  msoFileTypeProjectFiles = 19
  msoFileTypePublisherFiles = 18
  msoFileTypeTaskItem = 15
  msoFileTypeTemplates = 8
  msoFileTypeVisioFiles = 21
  msoFileTypeWebPages = 23
  msoFileTypeWordDocuments = 3
  msoFileTypeUserSpecified = 1000 ' DWF-Added Constant
End Enum
Public Enum osfMatchTextType
  osfMatchTextTypeMatchAll = 1
  osfMatchTextTypeMatchExact = 2
  osfMatchTextTypeMatchAny = 3
End Enum

'=====================================================================
' Public Properties/Collections
'=====================================================================
Public Property Let Creator(pstrCreator As String)
  strCreator = pstrCreator
End Property

Public Property Get Creator() As String
  Creator = strCreator
End Property

Public Property Let ExecuteCancel(pbolCancelExecution As Boolean)
  bolCancelExecution = pbolCancelExecution
End Property

Public Property Get ExecuteCancel() As Boolean
  ExecuteCancel = bolCancelExecution
End Property

Public Property Let fileName(pstrFileName As String)
  strFileName = pstrFileName
End Property

Public Property Get fileName() As String
  fileName = strFileName
End Property

Public Property Let FileType(pintFileType As MsoFileType)
  intFileType = pintFileType
  'ProcessFileTypes
End Property

Public Property Get FileType() As MsoFileType
  FileType = intFileType
End Property

Public Property Let FileTypeSpecify(pstrFileType As String)
  Dim strExtensionList() As String
  Dim i As Integer
  Dim lngUBound As Long
  Dim strTemp As String
  
  If InStr(pstrFileType, ",") Then
     strExtensionList() = Split(pstrFileType, ",")
     lngUBound = UBound(strExtensionList)
     For i = 0 To lngUBound
       strTemp = Trim(GetFileExtension(strExtensionList(i)))
       If Len(strTemp) > 0 Then
          mcolFileTypes.Add strTemp, strTemp
       End If
     Next i
  Else
     strTemp = GetFileExtension(pstrFileType)
     mcolFileTypes.Add strTemp, strTemp
  End If
End Property

Public Property Get FileTypes() As Collection
  Set FileTypes = mcolFileTypes
End Property

Public Property Get FilesToCheck() As Collection
  Set FilesToCheck = mcolFilesToCheck
End Property

Public Property Get FoundFiles() As Collection
  Set FoundFiles = mcolFoundFiles
End Property

Public Property Get FSO() As Object
  If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set FSO = objFSO
End Property

Public Property Let LastModified(pintLastModified As MsoLastModified)
  intLastModified = pintLastModified
  dteEndDate = Date + 1
  Select Case intLastModified
    Case msoLastModifiedAnyTime
      ' do nothing as this will cause the date to be ignored
    Case msoLastModifiedLastMonth
      dteStartDate = DateAdd("m", -2, Date)
      dteEndDate = DateAdd("m", -1, Date) + 1
    Case msoLastModifiedLastWeek
      dteStartDate = DateAdd("ww", -2, Date)
      dteEndDate = DateAdd("ww", -1, Date) + 1
    Case msoLastModifiedThisMonth
      dteStartDate = DateAdd("m", -1, Date)
    Case msoLastModifiedThisWeek
      dteStartDate = DateAdd("ww", -1, Date)
    Case msoLastModifiedToday
      dteStartDate = Date
    Case msoLastModifiedYesterday
      dteStartDate = Date - 1
      dteEndDate = Date
  End Select
End Property

Public Property Get LastModified() As MsoLastModified
  LastModified = intLastModified
End Property

Public Sub LastModifiedSpecify(Optional pdteStartDate As Date = 0, Optional pdteEndDate As Variant)
  LastModified = msoLastModifiedSpecify
  dteStartDate = pdteStartDate
  If IsMissing(pdteEndDate) Then
     dteEndDate = Date + 1
  Else
     dteEndDate = CDate(pdteEndDate) + 1
  End If
End Sub

Public Property Get LastModifiedSpecifyStart() As Date
  LastModifiedSpecifyStart = dteStartDate
End Property

Public Property Get LastModifiedSpecifyEnd() As Date
  LastModifiedSpecifyEnd = dteEndDate
End Property

Public Property Let LookIn(pstrLookin As String)
  strLookIn = pstrLookin
End Property

Public Property Get LookIn() As String
  LookIn = strLookIn
End Property

Public Property Let MatchTextExactly(pbolMatchTextExactly As Boolean)
  bolMatchTextExactly = pbolMatchTextExactly
  intMatchTextType = osfMatchTextTypeMatchExact
End Property

Public Property Get MatchTextExactly() As Boolean
  MatchTextExactly = bolMatchTextExactly
End Property

Public Property Let MatchTextType(pintMatchTextType As osfMatchTextType)
  intMatchTextType = pintMatchTextType
  bolMatchTextExactly = (pintMatchTextType = osfMatchTextTypeMatchExact)
End Property

Public Property Get MatchTextType() As osfMatchTextType
  MatchTextType = intMatchTextType
End Property

'Public Property Let ProgressBarForm(pfrmProgressBar As Form)
'  Set frmProgressBar = pfrmProgressBar
'End Property

Public Property Let SearchPropertiesOnly(pbolSearchPropertiesOnly As Boolean)
  bolSearchPropertiesOnly = pbolSearchPropertiesOnly
End Property

Public Property Get SearchPropertiesOnly() As Boolean
  SearchPropertiesOnly = bolSearchPropertiesOnly
End Property

Public Property Let SearchSubFolders(pbolSearchSubFolders As Boolean)
  bolSearchSubFolders = pbolSearchSubFolders
End Property

Public Property Get SearchSubFolders() As Boolean
  SearchSubFolders = bolSearchSubFolders
End Property

Public Property Let TextOnly(pbolTextOnly As Boolean)
  bolTextOnly = pbolTextOnly
End Property

Public Property Get TextOnly() As Boolean
  TextOnly = bolTextOnly
End Property

Public Property Let TextOrProperty(pstrTextOrProperty As String)
  strTextOrProperty = pstrTextOrProperty
End Property

Public Property Get TextOrProperty() As String
  TextOrProperty = strTextOrProperty
End Property

'=====================================================================
' Public Methods
'=====================================================================
Public Function Execute(Optional strSpecificProperty As String) As Integer
On Error GoTo errHandler
  Dim strCriteria() As String
  Dim varItem As Variant
  Dim strFileMatch As String
  Dim intSearchReturnValue As Integer
  Dim bolAddFile As Boolean

  Call ProcessFileTypes
  'Call UpdateProgressBar
  Call PopulateFileList(strLookIn)
  'Call UpdateProgressBar
  DoEvents
  If Me.ExecuteCancel Then GoTo exitRoutine
  If Len(strTextOrProperty) > 0 Then
     Select Case MatchTextType
       Case osfMatchTextTypeMatchExact
         ReDim strCriteria(0)
         strCriteria(0) = strTextOrProperty
       Case osfMatchTextTypeMatchAll, osfMatchTextTypeMatchAny
         strCriteria() = Split(strTextOrProperty, " ")
     End Select
  End If
  For Each varItem In FilesToCheck
    strFileMatch = FilesToCheck(varItem)
    'If Len(Me.TextOrProperty) > 0 Then
    '   DoCmd.Echo True, strFileMatch
    'End If
    If Len(Me.TextOrProperty) <> 0 Then
       'Call UpdateProgressBar
       DoEvents
       If Me.ExecuteCancel Then GoTo exitRoutine
       intSearchReturnValue = SearchFileForText(strFileMatch, strCriteria(), strSpecificProperty)
       If intSearchReturnValue = 429 Then
          MsgBox "Cancelling search because of unregistered DLL (DSOFile.dll)", vbInformation, "Cancelling search!"
          Exit For
       End If
       'Call UpdateProgressBar
       bolAddFile = (intSearchReturnValue = True)
    Else
       bolAddFile = True
    End If
    If bolAddFile Then
       mcolFoundFiles.Add strFileMatch, strFileMatch
    End If
  Next varItem
  Execute = mcolFoundFiles.Count
  
exitRoutine:
  Exit Function
  
errHandler:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in clFileSearch.Execute()"
  Resume exitRoutine
End Function

Public Sub NewSearch()
  strLookIn = vbNullString
  strFileName = vbNullString
  strTextOrProperty = vbNullString
  dteEndDate = Date
  dteStartDate = 0
  intLastModified = msoLastModifiedAnyTime
  intFileType = msoFileTypeAllFiles
  intMatchTextType = osfMatchTextTypeMatchAll
  bolMatchTextExactly = False
  bolCancelExecution = False
  Set mcolFilesToCheck = New Collection
  Set mcolFoundFiles = New Collection
  Set mcolFileTypes = New Collection
  If Not objDSOFileDoc Is Nothing Then
     objDSOFileDoc.Close
  End If
  'Set frmProgressBar = Nothing
End Sub

'=====================================================================
' Private Properties/Collections
'=====================================================================
Private Property Get DSOFileDoc() As Object
' see http://support.microsoft.com/kb/224351 to download DSOFile.dll
'   register it to make code work
On Error GoTo errHandler

  If objDSOFileDoc Is Nothing Then
     Set objDSOFileDoc = CreateObject("DSOFile.OleDocumentProperties")
  End If
  Set DSOFileDoc = objDSOFileDoc

exitRoutine:
  Exit Property

errHandler:
  Select Case Err.Number
    Case 429 ' ActiveX component can't create object
      MsgBox Err.Number & ": " & Err.Description & vbCrLf & " " & vbCrLf & "Likely the DSOFile.dll library is not registered. See http://support.microsoft.com/kb/224351 to download it.", vbInformation, "Error initializing DSOFile object (DSOFileDoc Property Get)"
      Err.Clear
    Case Else
      MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error DSOFileDoc Property Get"
  End Select
  Resume exitRoutine
End Property

'=====================================================================
' Private Methods
'=====================================================================
Private Function CheckCollectionForValue(mcolCollection As Collection, strValue As String) As Boolean
On Error GoTo errHandler

  CheckCollectionForValue = (Len(mcolCollection(strValue)) <> 0)

exitRoutine:
  Exit Function
  
errHandler:
  Select Case Err.Number
    Case 5 ' invalid procedure call or argument
      ' ignore the error
    Case Else
      MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in CheckCollectionForValue()"
      Resume exitRoutine
  End Select
End Function

Private Function CheckExtension(strExtension As String) As Boolean
  CheckExtension = CheckCollectionForValue(Me.FileTypes, strExtension)
End Function

Private Function GetFileExtension(strFileName As String) As String
  Dim i As Integer
  Dim tmpOutput As String

  If (InStr(strFileName, ".") = 0) Then Exit Function
  For i = Len(strFileName) To 1 Step -1
    If Mid(strFileName, i, 1) = "." Then Exit For
    tmpOutput = Mid(strFileName, i, 1) & tmpOutput
  Next i
  GetFileExtension = tmpOutput
End Function

Private Function ExecuteSearchProperties(strFileName As String, strCriteria() As String, Optional strSpecificProperty As String)
On Error GoTo errHandler
  Dim bolTemp As Boolean
  Dim i As Integer
  Dim lngUBound As Long
  Dim strSearchText As String
     
  'Call UpdateProgressBar
  DSOFileDoc.Open strFileName, True, 2 ' dsoOptionOpenReadOnlyIfNoWriteAccess
  If Not objDSOFileDoc Is Nothing Then
     lngUBound = UBound(strCriteria)
     For i = 0 To lngUBound
       'Call UpdateProgressBar
       DoEvents
       If Me.ExecuteCancel Then GoTo exitRoutine
       strSearchText = strCriteria(i)
       'Call UpdateProgressBar
       If IsDate(strSearchText) Then
          bolTemp = SearchPropertiesDate(DateValue(strSearchText), strSpecificProperty)
       Else
          bolTemp = SearchPropertiesText(strSearchText, strSpecificProperty)
       End If
       Select Case Me.MatchTextType
         Case osfMatchTextTypeMatchAll, osfMatchTextTypeMatchExact
           If Not bolTemp Then i = UBound(strCriteria)
         Case osfMatchTextTypeMatchAny
           If bolTemp Then i = UBound(strCriteria)
       End Select
     Next i
  End If
  ExecuteSearchProperties = bolTemp

exitRoutine:
  DSOFileDoc.Close
  'Call UpdateProgressBar
  Exit Function

errHandler:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in ExecuteSearchProperties()"
  Resume exitRoutine
End Function

Private Function ExecuteSearchFile(strFileName As String, strCriteria() As String) As Boolean
On Error GoTo errHandler
  Dim intFile As Integer
  Dim strFileContent As String
  Dim bolTemp As Boolean
  Dim i As Integer
  Dim lngUBound As Long
     
  'Call UpdateProgressBar
  intFile = FreeFile
  Open strFileName For Binary As #intFile
  strFileContent = String(LOF(intFile), " ")
  Get #intFile, , strFileContent
  ' might use RegExp CreateObject("VBScript.Regexp") Microsoft VBScript Regular Expressions
  lngUBound = UBound(strCriteria)
  For i = 0 To lngUBound
    'Debug.Print GetFileName(strFileName)
    'Call UpdateProgressBar
    DoEvents
    If Me.ExecuteCancel Then GoTo exitRoutine
    bolTemp = (InStr(strFileContent, strCriteria(i)) <> 0)
    Select Case MatchTextType
      Case osfMatchTextTypeMatchAll, osfMatchTextTypeMatchExact
        bolTemp = (InStr(strFileContent, strCriteria(i)) <> 0)
        If Not bolTemp Then i = UBound(strCriteria())
      Case osfMatchTextTypeMatchAny
        If bolTemp Then i = UBound(strCriteria())
    End Select
  Next i
  ExecuteSearchFile = bolTemp

exitRoutine:
  Close #intFile
  'Call UpdateProgressBar
  Exit Function
  
errHandler:
  MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in ExecuteFileSearch()"
  Resume exitRoutine
End Function

Private Sub PopulateFileList(strSearchFolder As String)
  Dim strFileMatch As String
  Dim strExtension As String
  Dim objFSOParentFolder As Object
  Dim objFSOFolder As Object
  
  strFileMatch = Dir(strSearchFolder & "\" & strFileName)
  If Len(strFileMatch) = 0 And Not bolSearchSubFolders Then
     Exit Sub
  ElseIf Len(strFileMatch) = 0 Then
     GoTo SearchSubFolders
  End If
  Do Until Len(strFileMatch) = 0
     DoEvents
     'Call UpdateProgressBar
     If Me.ExecuteCancel Then GoTo exitRoutine
     'If Len(Me.TextOrProperty) = 0 Then
     '   DoCmd.Echo True, strSearchFolder & "\" & strFileMatch
     'End If
     strExtension = GetFileExtension(strFileMatch)
     If FileType = msoFileTypeAllFiles Or CheckExtension(strExtension) Then
        mcolFilesToCheck.Add strSearchFolder & "\" & strFileMatch, strSearchFolder & "\" & strFileMatch
     End If
     strFileMatch = Dir
  Loop
  
SearchSubFolders:
  If bolSearchSubFolders Then
     Set objFSOParentFolder = FSO.GetFolder(strSearchFolder)
     If objFSOParentFolder.SubFolders.Count > 0 Then
        For Each objFSOFolder In objFSOParentFolder.SubFolders
          DoEvents
          If Me.ExecuteCancel Then GoTo exitRoutine
          Call PopulateFileList(objFSOFolder.Path)
        Next
     End If
     Set objFSOParentFolder = Nothing
  End If

exitRoutine:
  'Call UpdateProgressBar
  Exit Sub
End Sub

Private Sub ProcessFileTypes()
' reference for extensions:
'   http://msdn.microsoft.com/es-es/library/microsoft.office.core.msofiletype(VS.80).aspx
' template extensions commented out except for msoFileTypeTemplates
  If intFileType = 0 Then
     Set mcolFileTypes = Nothing
  Else
     Select Case intFileType
       Case msoFileTypeAllFiles
         mcolFileTypes.Add "*", "*"
       Case msoFileTypeBinders
         mcolFileTypes.Add "obd", "obd"
         'mcolFileTypes.Add "obt", "obt"
       Case msoFileTypeCalendarItem
         mcolFileTypes.Add "ics", "ics"
         mcolFileTypes.Add "vsc", "vsc"
       Case msoFileTypeContactItem
         mcolFileTypes.Add "vcf", "vcf"
       Case msoFileTypeDatabases
         mcolFileTypes.Add "mdb", "mdb"
         mcolFileTypes.Add "mde", "mde"
         mcolFileTypes.Add "mdr", "mdr"
         mcolFileTypes.Add "accdb", "accdb"
         mcolFileTypes.Add "accde", "accde"
         mcolFileTypes.Add "accdr", "accdr"
       Case msoFileTypeDataConnectionFiles
         mcolFileTypes.Add "mdf", "mdf"
       Case msoFileTypeDesignerFiles
         mcolFileTypes.Add "dsr", "dsr"
       Case msoFileTypeDocumentImagingFiles
         mcolFileTypes.Add "mdi", "mdi"
       Case msoFileTypeExcelWorkbooks
         mcolFileTypes.Add "xls", "xls"
         'mcolFileTypes.Add "xlt", "xlt"
         mcolFileTypes.Add "wbk", "wbk"
         mcolFileTypes.Add "xlsx", "xlsx"
         mcolFileTypes.Add "xlsm", "xlsm"
         'mcolFileTypes.Add "xltx", "xltx"
         'mcolFileTypes.Add "xltm", "xltm"
         mcolFileTypes.Add "xlsb", "xlsb"
         mcolFileTypes.Add "xlam", "xlam"
       Case msoFileTypeJournalItem
         ' NOT IMPLEMENTED
         MsgBox "Outlook Journal Items (msoFileTypeJournalItem) not implemented"
       Case msoFileTypeMailItem
         mcolFileTypes.Add "msg", "msg"
       Case msoFileTypeNoteItem
         ' NOT IMPLEMENTED
         MsgBox "Outlook Note Items (msoFileTypeNoteItem) not implemented"
       Case msoFileTypeOfficeFiles
         mcolFileTypes.Add "doc", "doc"
         mcolFileTypes.Add "docx", "docx"
         mcolFileTypes.Add "docm", "docm"
         'mcolFileTypes.Add "dot", "dot"
         'mcolFileTypes.Add "dotx", "dotx"
         mcolFileTypes.Add "dotm", "dotm"
         mcolFileTypes.Add "xls", "xls"
         mcolFileTypes.Add "xlsx", "xlsx"
         mcolFileTypes.Add "xlsm", "xlsm"
         'mcolFileTypes.Add "xlt", "xlt"
         'mcolFileTypes.Add "xltx", "xltx"
         'mcolFileTypes.Add "xltm", "xltm"
         mcolFileTypes.Add "xlsb", "xlsb"
         mcolFileTypes.Add "xlam", "xlam"
         'mcolFileTypes.Add "ppt", "ppt"
         mcolFileTypes.Add "pot", "pot"
         mcolFileTypes.Add "pps", "pps"
         'mcolFileTypes.Add "pptx", "pptx"
         mcolFileTypes.Add "ppsx", "ppsx"
         'mcolFileTypes.Add "pptm", "pptm"
         mcolFileTypes.Add "ppsm", "ppsm"
         'mcolFileTypes.Add "potx", "potx"
         mcolFileTypes.Add "potm", "potm"
         mcolFileTypes.Add "ppam", "ppam"
         mcolFileTypes.Add "sldx", "sldx"
         mcolFileTypes.Add "sldm", "sldm"
         mcolFileTypes.Add "thmx", "thmx"
         mcolFileTypes.Add "obd", "obd"
         'mcolFileTypes.Add "obt", "obt"
         mcolFileTypes.Add "mdb", "mdb"
         mcolFileTypes.Add "mde", "mde"
         mcolFileTypes.Add "mdr", "mdr"
         mcolFileTypes.Add "accdb", "accdb"
         mcolFileTypes.Add "accde", "accde"
         mcolFileTypes.Add "accdr", "accdr"
         mcolFileTypes.Add "htm", "htm"
         mcolFileTypes.Add "html", "html"
       Case msoFileTypeOutlookItems
         ' NOT IMPLEMENTED
         MsgBox "All Outlook items (msoFileTypeOutlookItems) not implemented"
       Case msoFileTypePhotoDrawFiles
         mcolFileTypes.Add "mix", "mix"
       Case msoFileTypePowerPointPresentations
         mcolFileTypes.Add "ppt", "ppt"
         'mcolFileTypes.Add "pot", "pot"
         mcolFileTypes.Add "pps", "pps"
         mcolFileTypes.Add "pptx", "pptx"
         mcolFileTypes.Add "ppsx", "ppsx"
         'mcolFileTypes.Add "pptm", "pptm"
         mcolFileTypes.Add "ppsm", "ppsm"
         'mcolFileTypes.Add "potx", "potx"
         mcolFileTypes.Add "potm", "potm"
         mcolFileTypes.Add "ppam", "ppam"
         mcolFileTypes.Add "ppsx", "ppsx"
         mcolFileTypes.Add "ppsm", "ppsm"
       Case msoFileTypeProjectFiles
         mcolFileTypes.Add "mpd", "mpd"
       Case msoFileTypePublisherFiles
         mcolFileTypes.Add "pub", "pub"
       Case msoFileTypeTaskItem
         ' NOT IMPLEMENTED
         MsgBox "Outlook Task items (msoFileTypeOutlookItems) not implemented"
       Case msoFileTypeTemplates
         mcolFileTypes.Add "dot", "dot"
         mcolFileTypes.Add "dotx", "dotx"
         mcolFileTypes.Add "xlt", "xlt"
         mcolFileTypes.Add "xltx", "xltx"
         mcolFileTypes.Add "xltm", "xltm"
         mcolFileTypes.Add "pot", "pot"
         mcolFileTypes.Add "potx", "potx"
         mcolFileTypes.Add "potm", "potm"
       Case msoFileTypeVisioFiles
         mcolFileTypes.Add "vsd", "vsd"
       Case msoFileTypeWebPages
         mcolFileTypes.Add "htm", "htm"
         mcolFileTypes.Add "html", "html"
       Case msoFileTypeWordDocuments
         mcolFileTypes.Add "doc", "doc"
         'mcolFileTypes.Add "dot", "dot"
         mcolFileTypes.Add "docx", "docx"
         mcolFileTypes.Add "docm", "docm"
         'mcolFileTypes.Add "dotx", "dotx"
         'mcolFileTypes.Add "dotm", "dotm"
     End Select
  End If
End Sub

Private Function SearchFileForText(strFileName As String, strCriteria() As String, Optional strSpecificProperty As String) As Integer
On Error GoTo errHandler
  Dim bolCheckText As Boolean
  Dim bolCheckDate As Boolean
  Dim bolCheckProp As Boolean
  Dim bolMatchText As Boolean
  Dim bolMatchDate As Boolean
  Dim bolMatchProp As Boolean
  Dim intFile As Integer
  Dim strFileContent As String
  Dim i As Integer
  Dim dteLastSaved As Date
  
  ' CHECK FULL TEXT
  bolCheckText = (Not bolSearchPropertiesOnly) And (Len(Me.TextOrProperty) > 0)
  If bolCheckText Then
     bolMatchText = ExecuteSearchFile(strFileName, strCriteria())
  End If
  ' CHECK PROPERTIES
  bolCheckProp = ((Len(Me.TextOrProperty) > 0) And Not bolTextOnly) Or (Len(Me.Creator) > 0)
  If bolCheckProp Then
     bolMatchProp = ExecuteSearchProperties(strFileName, strCriteria())
  End If
  ' CHECK FILE DATE
  bolCheckDate = (Me.LastModified <> msoLastModifiedAnyTime)
  If bolCheckDate Then
     dteLastSaved = FileDateTime(strFileName)
     bolMatchDate = (dteLastSaved >= dteStartDate And dteLastSaved < dteEndDate)
  End If
  ' PARSE RESULTS
  If bolCheckText And Not bolCheckProp Then ' checked ONLY text
     If bolCheckDate Then ' text only and date
        SearchFileForText = (bolMatchProp And bolMatchDate)
     Else                 ' text ONLY
        SearchFileForText = bolMatchText
     End If
  ElseIf bolCheckText And bolCheckProp Then ' checked both properties and text
     If bolCheckDate Then ' text/properties and date
        SearchFileForText = (bolMatchText Or bolMatchProp) And bolMatchDate
     Else                 ' text/properties ONLY
        SearchFileForText = (bolMatchText Or bolMatchProp)
     End If
  ElseIf Not bolCheckText And bolCheckProp Then ' checked properties but not text
     If bolCheckDate Then ' properties and date
        SearchFileForText = bolMatchProp And bolMatchDate
     Else                 ' properties ONLY
        SearchFileForText = bolMatchProp
     End If
  ElseIf bolCheckDate And Not (bolCheckText Or bolCheckProp) Then ' checked file date only
     SearchFileForText = bolMatchDate
  End If
  
exitRoutine:
  Exit Function
  
errHandler:
  Select Case objDSOFileDoc Is Nothing
    Case True
      If Err.Number = 91 Then
         'Object variable or With block variable not set
         ' ignore error as it's already been handled in the DSOFileDoc Property Get
         SearchFileForText = 429
         Err.Clear
      End If
    Case False
      MsgBox Err.Number & ": " & Err.Description, vbExclamation, "Error in clFileSearch.SearchFileForText()"
  End Select
  Resume exitRoutine
End Function

Private Function SearchPropertiesDate(dteSearch As Date, Optional strPropertyName As String) As Boolean
  Dim bolSingleProperty As Boolean
  Dim bolTemp As Boolean
  Dim dteLastSaved As Date
  
  bolSingleProperty = (Len(strPropertyName) > 0)
  If bolSingleProperty Then
    Select Case strPropertyName
      Case "DateLastSaved"
        GoTo DateLastSaved
      Case "DateCreated"
        GoTo DateCreated
      Case "DateLastPrinted"
        GoTo DateLastPrinted
    End Select
  End If
  If dteSearch = 0 Then GoTo exitRoutine
DateLastSaved:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.DateLastSaved, dteSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
DateCreated:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.DateCreated, dteSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
DateLastPrinted:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.DateLastPrinted, dteSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine

exitRoutine:
  SearchPropertiesDate = bolTemp
End Function

Private Function SearchPropertiesText(strSearch As String, Optional strPropertyName As String) As Boolean
  Dim bolSingleProperty As Boolean
  Dim bolTemp As Boolean
  Dim dteLastSaved As Date
  Dim i As Integer
  Dim strTemp As String
  
  bolSingleProperty = (Len(strPropertyName) > 0)
  If bolSingleProperty Then
    Select Case strPropertyName
      Case "ApplicationName"
        GoTo ApplicationName
      Case "Version"
        GoTo Version
      Case "Title"
        GoTo title
      Case "Subject"
        GoTo Subject
      Case "Category"
        GoTo Category
      Case "Company"
        GoTo Company
      Case "Keywords"
        GoTo Keywords
      Case "Manager"
        GoTo Manager
      Case "LastSavedBy"
        GoTo LastSavedBy
      Case "WordCount"
        GoTo WordCount
      Case "PageCount"
        GoTo PageCount
      Case "ParagraphCount"
        GoTo ParagraphCount
      Case "LineCount"
        GoTo LineCount
      Case "ParagraphCount"
        GoTo ParagraphCount
      Case "CharacterCount"
        GoTo CharacterCount
      Case "CharacterCountWithSpaces"
        GoTo CharacterCountWithSpaces
      Case "ByteCount"
        GoTo ByteCount
      Case "PresentationFormat"
        GoTo PresentationFormat
      Case "SlideCount"
        GoTo SlideCount
      Case "NoteCount"
        GoTo NoteCount
      Case "HiddenSlideCount"
        GoTo HiddenSlideCount
      Case "MultimediaClipCount"
        GoTo MultimediaClipCount
      Case "TotalEditTime"
        GoTo TotalEditTime
      Case "Template"
        GoTo Template
      Case "RevisionNumber"
        GoTo RevisionNumber
      Case "SharedDocument"
        GoTo SharedDocument
      Case Else
        GoTo Custom
    End Select
  End If
  If Len(Me.Creator) Then
     bolTemp = InStr(DSOFileDoc.SummaryProperties.Author, strSearch)
     If bolTemp Or bolSingleProperty Or bolSingleProperty Then GoTo exitRoutine
  End If
  If Len(strSearch) = 0 Then GoTo exitRoutine
ApplicationName:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.ApplicationName, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Version:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Version, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
title:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.title, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Subject:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Subject, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Category:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Category, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Company:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Company, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Keywords:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Keywords, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Manager:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Manager, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
LastSavedBy:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.LastSavedBy, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
WordCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.WordCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
PageCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.PageCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
ParagraphCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.ParagraphCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
LineCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.LineCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
CharacterCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.CharacterCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
CharacterCountWithSpaces:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.CharacterCountWithSpaces, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
ByteCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.ByteCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
PresentationFormat:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.PresentationFormat, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
SlideCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.SlideCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
NoteCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.NoteCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
HiddenSlideCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.HiddenSlideCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
MultimediaClipCount:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.MultimediaClipCount, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
TotalEditTime:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.TotalEditTime, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Template:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.Template, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
RevisionNumber:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.RevisionNumber, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
SharedDocument:
  bolTemp = InStr(DSOFileDoc.SummaryProperties.SharedDocument, strSearch)
  If bolTemp Or bolSingleProperty Then GoTo exitRoutine
Custom:
' custom properties do not seem to work for all document types
  For i = 0 To DSOFileDoc.CustomProperties.Count - 1
    If bolSingleProperty And (DSOFileDoc.CustomProperties(i).Name <> strPropertyName) Then
       ' do nothing
    Else
       strTemp = CStr(DSOFileDoc.CustomProperties(i).Value)
       bolTemp = InStr(strTemp, strSearch)
       If bolTemp Then Exit For
    End If
  Next i

exitRoutine:
  SearchPropertiesText = bolTemp
End Function

'Private Sub UpdateProgressBar()
'  If Not (frmProgressBar Is Nothing) Then
'     Call frmProgressBar.ProgressBar
'  End If
'End Sub

'=====================================================================
' Class Operations
'=====================================================================
Private Sub Class_Initialize()
  Set mcolFilesToCheck = New Collection
  Set mcolFoundFiles = New Collection
  Set mcolFileTypes = New Collection
  intLastModified = msoLastModifiedAnyTime
  intFileType = msoFileTypeAllFiles
  bolMatchTextExactly = False
  intMatchTextType = osfMatchTextTypeMatchAll
  bolCancelExecution = False
  'Set frmProgressBar = Nothing
End Sub

Private Sub Class_Terminate()
  Set objDSOFileDoc = Nothing
End Sub

