Attribute VB_Name = "FolderRecurse"
Option Explicit
  
Option Compare Text
  
 'Recursion Bas Module -- Except that there is no recursion.  Other than that it is right, it is a Basic Module.
  
  Public Const MAX_UPDT_INT = 2999  ' Minimum number of seconds to wait before updating the progress stats.  If over, update.
  
  Public bDoDebugPrints As Boolean
  
  Public bIgnoreClicks As Boolean
  
  Public Const ARRAY_EXPANSION = 100  ' Number of elements to add to arrays when they are full and need another element added.
  Public iScanPass As Integer         ' 0 for Master scan, 1 for Slave scan
  Public iLogFile As Integer          ' Write out progress information.
  Public sMasterPath As String        ' (-mp) Master folder path from comand line
  Public sSlavePath As String         ' (-sp) Slave folder path from command line
  Public bCopyMissing As Boolean      ' (-cm) Command line said to copy Missing files from Master to Slave
  Public bCopyNewer As Boolean        ' (-cn) Command line said to copy Newer files on Master to Slave
  Public bCopyBack As Boolean         ' (-cb) Command line said (ouch!) to copy back Newer files on Slave to Master.  Take care with this one.
  Public bRemoveOrphans As Boolean    ' (-do) Command line said to delete files on Slave no longer on Master.
  Public bAutomaticRun As Boolean     ' If any option is set, set this to show to run selected options and exit.
  
  Public LstItem As ListItem
  
  Public AllItems() As ItemInfo
  Public lAllItems As Long  ' Used to count what is in AllItems and also to set a max for the gauge.
  
  Public MissingItems() As ItemInfo
  Public lMissingItems As Long
  
  Public NewerMaster() As ItemInfo
  Public lNewerMaster As Long
  
  Public NewerSlave() As ItemInfo
  Public lNewerSlave As Long
  
  Public Orphans() As ItemInfo  ' Files/Folders in Slave folder but not in Master folder.
  Public lOrphans As Long
  
  Public lTotalFilesFound As Long    ' Count of all files found
  Public lTotalFoldersFound As Long  ' Count of all folders found
  
  Public StopRequested As Boolean
  
  Public oFso As New Scripting.FileSystemObject
  
  Public Type ItemInfo
    ItemName As String
    ItemType As Boolean
  End Type
  
  Public Type ItemInfoD
    ItemName As String
    ItemType As Boolean
    DateCreated As String
  End Type
  
  Public Const ITEM_FILE = False
  Public Const ITEM_FOLDER = True
  
  Public Type FileInformation
      Folder As String
      Path As String
      Title As String
      Size As Long
  End Type
  
  Public bCancelled As Boolean  ' True if user clicks Stop button
  
  Public gsDirsQueue As New Collection

  Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
  
  Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                  ByVal wParam As Long, lParam As Any) As Long

  Public Const LVM_FIRST = &H1000
  
 'Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
  
  Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
  
  Declare Function timeGetTime Lib "winmm.dll" () As Long
  Declare Function timeBeginPeriod Lib "winmm.dll" (ByVal uPeriod As Integer) As Long
  Declare Function timeEndPeriod Lib "winmm.dll" (ByVal uPeriod As Integer) As Long
  
Public Function FixPath(sPath As String) As String

  While Right(sPath, 1) = "\"
    sPath = Left(sPath, Len(sPath) - 1)
  Wend
  FixPath = sPath
  
End Function

Sub MLMkdir(sPath As String)

  Dim i As Integer
  sPath = FixPath(sPath) & "\"
  On Error Resume Next
  For i = 1 To Len(sPath)
    If Mid$(sPath, i, 1) = "\" Then
      If i > 3 Then
        Debug.Print "Trying to Mkdir " & Mid$(sPath, 1, i - 1)
        On Error GoTo BadCreate
        MkDir Mid$(sPath, 1, i - 1)
      End If
    End If
  Next
  Exit Sub
  
BadCreate:
  If Err.Number = 75 Then
    Resume Next
  Else
    MsgBox "Error " & Err.Number & " -- " & Err.Description & " trying to create " & Mid$(sPath, 1, i - 1)
    StopRequested = True
    Exit Sub
  End If
  
End Sub
Public Function KillFolder(ByVal FullPath As String) As Boolean
     
 '******************************************
 'PURPOSE: DELETES A FOLDER, INCLUDING ALL SUB-
 '         DIRECTORIES, FILES, REGARDLESS OF THEIR
 '         ATTRIBUTES
 '
 'PARAMETER: FullPath = FullPath of Folder to Delete
 '
 'RETURNS:   True is successful, false otherwise
 '
 'REQUIRES:  'VB6
 '            Reference to Microsoft Scripting Runtime
 '            Caution in use for obvious reasons
 '
 'EXAMPLE:   'KillFolder("D:\MyOldFiles")
  
 '******************************************
  
 'deletefolder method does not like the "\" at end of fullpath

 'If Right(FullPath, 1) = "\" Then FullPath = Left(FullPath, Len(FullPath) - 1)
  FullPath = FixPath(FullPath)
  If oFso.FolderExists(FullPath) Then
      
    'Setting the 2nd parameter to true forces deletion of read-only files and subfolders.
    oFso.DeleteFolder FullPath, True
    
    KillFolder = Err.Number = 0 And oFso.FolderExists(FullPath) = False
  
  End If

End Function

Sub FindFilesAndFolders(sPath As String)

  Dim sTemp As String
  Dim sDirTemp As String
  Dim iUpdateSeconds As Long
  
  timeBeginPeriod 1
  iUpdateSeconds = timeGetTime

  Set gsDirsQueue = Nothing
  gsDirsQueue.Add sPath  ' Prime the file pump
  
  lAllItems = 0
  lTotalFilesFound = 0
  lTotalFoldersFound = 0
  
  While gsDirsQueue.Count > 0 And Not bCancelled
    If timeGetTime - iUpdateSeconds > MAX_UPDT_INT Then
      iUpdateSeconds = timeGetTime
      frmManualRAID1.lblCurrentFolder = sDirTemp
      frmManualRAID1.lblTotalFiles(iScanPass) = Format$(lTotalFilesFound, "##,###,##0")
      frmManualRAID1.lblTotalFolders(iScanPass) = Format$(lTotalFoldersFound, "##,###,##0")
      frmManualRAID1.lblFoldersQ = gsDirsQueue.Count
      DoEvents
    End If
    sDirTemp = FixPath(gsDirsQueue(1))
    sTemp = ""
    On Error Resume Next
    sTemp = Dir$(sDirTemp & "\*.*", vbNormal + vbReadOnly + vbHidden + vbDirectory)
    On Error GoTo 0
    While sTemp <> ""
      If sTemp <> "." And sTemp <> ".." Then
        'If (GetAttr(sDirTemp & "\" & sTemp) And &H10) = vbDirectory Then
        If (GetFileAttributes(sDirTemp & "\" & sTemp) And &H10) = vbDirectory Then
          gsDirsQueue.Add Item:=sDirTemp & "\" & sTemp, After:=DirSearchB(gsDirsQueue, sDirTemp & "\" & sTemp) ' Comes and goes here.
          lAllItems = lAllItems + 1
          If lAllItems > UBound(AllItems) Then ReDim Preserve AllItems(lAllItems + ARRAY_EXPANSION)
          AllItems(lAllItems).ItemName = sDirTemp & "\" & sTemp
          AllItems(lAllItems).ItemType = ITEM_FOLDER
          lTotalFoldersFound = lTotalFoldersFound + 1
        Else  ' We have a file.
          lAllItems = lAllItems + 1
          If lAllItems > UBound(AllItems) Then ReDim Preserve AllItems(lAllItems + ARRAY_EXPANSION)
          AllItems(lAllItems).ItemName = sDirTemp & "\" & sTemp
          AllItems(lAllItems).ItemType = ITEM_FILE
          lTotalFilesFound = lTotalFilesFound + 1
        End If
      End If
      sTemp = Dir$()
    Wend
    
    gsDirsQueue.Remove (1)
  Wend
  
  frmManualRAID1.lblCurrentFolder = ""
  frmManualRAID1.lblTotalFiles(iScanPass) = Format$(lTotalFilesFound, "##,###,##0")
  frmManualRAID1.lblTotalFolders(iScanPass) = Format$(lTotalFoldersFound, "##,###,##0")
  
  timeEndPeriod 1
  
End Sub

Function FormatPath(ByVal FolderPath As String) As String

 'Simple function that does some checking of a folderpath and removes trailing slashes
  
  On Error Resume Next
  If Len(FolderPath) > 2 Then
    Do Until Right$(FolderPath, 1) <> "\"
      FolderPath = Left$(FolderPath, Len(FolderPath) - 1)
    Loop
    FolderPath = Replace$(FolderPath, "/", "\")
  End If
  
  If Len(FolderPath) > 2 Then
    FormatPath = Left$(FolderPath, 2) & Replace$(Mid$(FolderPath, 3), "\\", "\")
  Else
    FormatPath = FolderPath
  End If

End Function

Public Function TimeString(Seconds As Long, Optional Verbose As Boolean = False) As String

  'if verbose = false, returns
  'something like
  '02:22.08
  'if true, returns
  '2 hours, 22 minutes, and 8 seconds
  
  Dim lHrs As Long
  Dim lMinutes As Long
  Dim lSeconds As Long
  Dim sAns As String
  
  lSeconds = Seconds
  
  lHrs = Int(lSeconds / 3600)
  lMinutes = (Int(lSeconds / 60)) - (lHrs * 60)
  lSeconds = Int(lSeconds Mod 60)
  
  If lSeconds = 60 Then
    lMinutes = lMinutes + 1
    lSeconds = 0
  End If
  
  If lMinutes = 60 Then
    lMinutes = 0
    lHrs = lHrs + 1
  End If
  
  sAns = Format(CStr(lHrs), "#####0") & ":" & _
         Format(CStr(lMinutes), "00") & "." & _
         Format(CStr(lSeconds), "00")
  If Verbose Then sAns = TimeStringtoEnglish(sAns)
  
  TimeString = sAns

End Function

Public Function TimeStringtoEnglish(sTimeString As String) As String

  Dim sAns As String
  Dim sHour, sMin As String, sSec As String
  Dim iTemp As Integer, sTemp As String
  Dim iPos As Integer
  iPos = InStr(sTimeString, ":") - 1
  
  sHour = Left$(sTimeString, iPos)
  If CLng(sHour) <> 0 Then
      sAns = CLng(sHour) & " hour"
      If CLng(sHour) > 1 Then sAns = sAns & "s"
      sAns = sAns & ", "
  End If
  
  sMin = Mid$(sTimeString, iPos + 2, 2)
  
  iTemp = sMin
  
  If sMin = "00" Then
     sAns = IIf(Len(sAns), sAns & "0 minutes, and ", "")
  Else
     sTemp = IIf(iTemp = 1, " minute", " minutes")
     sTemp = IIf(Len(sAns), sTemp & ", and ", sTemp & " and ")
     sAns = sAns & Format$(iTemp, "##") & sTemp
  End If
  
  iTemp = Val(Right$(sTimeString, 2))
  sSec = Format$(iTemp, "#0")
  sAns = sAns & sSec & " second"
  If iTemp <> 1 Then sAns = sAns & "s"
  
  TimeStringtoEnglish = sAns

End Function

Public Function DirSearch(col As Collection, sDirToAdd As String) As Long

  If col.Count = 1 Then
    DirSearch = 1
    Exit Function
  End If
  
  Dim i As Long
  Dim iStart As Long
  Dim iMidPoint As Long
  
  Dim bFound As Boolean
  
  bFound = False
 'If bDoDebugPrints Then Debug.Print "There are " & col.Count & " items in the collection."
  
  iMidPoint = col.Count / 2
  If iMidPoint = 0 Then iMidPoint = 1
  
 'If bDoDebugPrints Then Debug.Print "Check midpoint item " & sDirToAdd & " > " & iMidPoint & ": " & col.Item(iMidPoint)
  
  If sDirToAdd > col.Item(iMidPoint) Then
    iStart = col.Count
  Else
    iStart = col.Count / 2 + 2
    If iStart > col.Count Then iStart = col.Count
  End If
  
  For i = iStart To 1 Step -1
   'If bDoDebugPrints Then Debug.Print "Comparing " & sDirToAdd & " > " & col.Item(i)
    If sDirToAdd > col.Item(i) Then
      bFound = True
      Exit For
    End If
  Next
  
  If bFound Then
    DirSearch = i
  Else
    DirSearch = col.Count
  End If
  
End Function

Function DirSearchB(col As Collection, sDirToAdd As String) As Long

 'Binary search now, for speed I hope.
 
 'Input: The complete folder path to add after it's partent in one of the directory collections.
 'Output: The correctly placed folder, added to the collection.
 'Process: See numbered comments, below.
 
 'But, first, if this is small just do a sequential search and save the overhead here.
  If col.Count < 30 Then
    DirSearchB = DirSearch(col, sDirToAdd)
    Exit Function
  End If
  
  Dim iLow As Long
  Dim iHigh As Long
  Dim iPivot As Long
  Dim iAdjust As Long
  
 '1. Initialize: iHigh=high element, iLow=low element, iPivot=(iHigh - iLow + 1) \ 2 + iLow
  iLow = 1: iHigh = col.Count: iPivot = (iHigh - iLow + 1) \ 2 + iLow
 '1a. Start loop and determine ending condition
  iAdjust = 1
 'Debug.Print "Top: " & "L:" & iLow & " H:" & iHigh; " P:" & iPivot & " A:" & iAdjust
  Do While iAdjust > 0
 '2. Compare New item to collection pivot element
   'Debug.Print sDirToAdd & " > " & col.Item(iPivot)
    If StrComp(sDirToAdd, col.Item(iPivot)) = 1 Then
 '3. When the new item is HIGHER than the collection element, set LOW = iPivot
      iLow = iPivot
    Else
 '4. When the new item is LOWER than the collection element, set HIGH = iPivot
      iHigh = iPivot
    End If
 '5. Find new iAdjust and iPivot: (iHigh - iLow + 1) \ 2 + iLow
    iAdjust = (iHigh - iLow + 1) \ 2  ' Used as one of the two exits from the loop
    iPivot = iAdjust + iLow
 '6. When the adjustment value is 0, we have found the two items between which the new item is to be placed.
   'Debug.Print "End: " & "L:" & iLow & " H:" & iHigh; " P:" & iPivot & " A:" & iAdjust
 '6a. Well, this is odd.  Not sure how this condition occurs, maybe an odd number of items in the collection.
 '    But, however it comes about, this is the right way to take care of it.  Tested it 1.6 million times!
    If iHigh - iLow = 1 Then
      DirSearchB = iLow  ' Return the item to add the new path after
      Exit Function
    End If
  Loop
  
  DirSearchB = iPivot  ' Return the item to add the new path after
 
End Function


