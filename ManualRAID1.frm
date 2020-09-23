VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.OCX"
Begin VB.Form frmManualRAID1 
   Caption         =   "Simulate RAID 1 Mirrored File/Folder Duplication"
   ClientHeight    =   9720
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11700
   LinkTopic       =   "Form1"
   ScaleHeight     =   9720
   ScaleWidth      =   11700
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboChoices 
      Height          =   330
      Left            =   120
      TabIndex        =   24
      Text            =   "cboChoices"
      Top             =   1340
      Width           =   4100
   End
   Begin VB.CommandButton cmdStop 
      BackColor       =   &H00C0E0FF&
      Caption         =   "&Stop"
      Height          =   420
      Left            =   1480
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   900
      Width           =   1360
   End
   Begin VB.CommandButton cmdPreview 
      BackColor       =   &H00C0FFC0&
      Caption         =   "&Preview"
      Default         =   -1  'True
      Height          =   420
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   900
      Width           =   1360
   End
   Begin VB.CommandButton cmdSync 
      BackColor       =   &H00FFFFC0&
      Caption         =   "S&ync"
      Enabled         =   0   'False
      Height          =   760
      Left            =   4200
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   900
      Width           =   1350
   End
   Begin VB.CommandButton cmdBFFSlave 
      BackColor       =   &H00C0E0FF&
      Caption         =   "..."
      Height          =   360
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   500
      Width           =   375
   End
   Begin VB.TextBox txtSlave 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Text            =   "C:\temp slave"
      Top             =   500
      Width           =   8265
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H008080FF&
      Cancel          =   -1  'True
      Caption         =   "E&xit"
      Height          =   420
      Left            =   2840
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   900
      Width           =   1360
   End
   Begin VB.CommandButton cmdBFFMaster 
      BackColor       =   &H00C0FFFF&
      Caption         =   "..."
      Height          =   360
      Left            =   8460
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   120
      Width           =   375
   End
   Begin VB.TextBox txtMaster 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\temp master"
      Top             =   120
      Width           =   8265
   End
   Begin MSComctlLib.ProgressBar pbAll 
      Height          =   270
      Left            =   120
      TabIndex        =   10
      Top             =   1670
      Visible         =   0   'False
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   476
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lstMissing 
      Height          =   915
      Left            =   120
      TabIndex        =   19
      Top             =   2040
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstMasterNew 
      Height          =   915
      Left            =   1420
      TabIndex        =   20
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "SourceDate"
         Text            =   "Master Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "ImageDate"
         Text            =   "Slave Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstSlaveNew 
      Height          =   915
      Left            =   2730
      TabIndex        =   21
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "SourceDate"
         Text            =   "Master Date"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "ImageDate"
         Text            =   "Slave Date"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstSizeVariations 
      Height          =   915
      Left            =   4040
      TabIndex        =   22
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "SourceSize"
         Text            =   "Master Size"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Key             =   "ImageSize"
         Text            =   "Slave Size"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.ListView lstOrphans 
      Height          =   915
      Left            =   5340
      TabIndex        =   23
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1614
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      MousePointer    =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "FileName"
         Text            =   "File Name"
         Object.Width           =   14111
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label lblFoldQ 
      Alignment       =   1  'Right Justify
      Caption         =   "Folder Queue:"
      Height          =   195
      Left            =   5520
      TabIndex        =   27
      Top             =   1470
      Visible         =   0   'False
      Width           =   1350
   End
   Begin VB.Label lblFoldersQ 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      Height          =   195
      Left            =   6900
      TabIndex        =   26
      Top             =   1470
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   7860
      X2              =   8760
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   6900
      X2              =   7800
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label lblSubState 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   9000
      TabIndex        =   25
      Top             =   1440
      UseMnemonic     =   0   'False
      Width           =   2535
      WordWrap        =   -1  'True
   End
   Begin VB.Label lblTotalFiles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   7860
      TabIndex        =   18
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label lblTotalFolders 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   7860
      TabIndex        =   17
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0E0FF&
      Caption         =   "Slave"
      Height          =   195
      Index           =   3
      Left            =   7860
      TabIndex        =   16
      Top             =   885
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      Caption         =   "Master"
      Height          =   195
      Index           =   2
      Left            =   6900
      TabIndex        =   15
      Top             =   885
      Width           =   900
   End
   Begin VB.Label lblTotalFolders 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   6900
      TabIndex        =   14
      Top             =   1275
      Width           =   900
   End
   Begin VB.Label lblTotalFiles 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   6900
      TabIndex        =   13
      Top             =   1080
      Width           =   900
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Folders Found:"
      Height          =   195
      Index           =   1
      Left            =   5520
      TabIndex        =   12
      Top             =   1275
      Width           =   1350
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Files Found:"
      Height          =   195
      Index           =   0
      Left            =   5520
      TabIndex        =   11
      Top             =   1080
      Width           =   1350
   End
   Begin VB.Label lblCurrentFolder 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Manual RAID 1  v.2  is Ready for You..."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   9360
      UseMnemonic     =   0   'False
      Width           =   11250
   End
   Begin VB.Label lblMainState 
      BorderStyle     =   1  'Fixed Single
      Height          =   495
      Left            =   9000
      TabIndex        =   9
      Top             =   960
      UseMnemonic     =   0   'False
      Width           =   2535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmManualRAID1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
  
  Dim iTabShowing As Long
  
  Dim SH As New Shell32.Shell   'reference to shell32.dll class
  Dim ShBFF As Shell32.Folder   'Shell Browse For Folder
  
Private Sub cboChoices_Click()
  If Not bIgnoreClicks Then ShowItemCountsInLists
End Sub

Private Sub ShowItemCountsInLists()

  lblSubState = ""
  Select Case cboChoices.ListIndex
    Case 0
      lstMissing.Visible = True
      lstMasterNew.Visible = False
      lstSlaveNew.Visible = False
      lstSizeVariations.Visible = False
      lstOrphans.Visible = False
      lblSubState = lstMissing.ListItems.Count & " items Missing on Slave"
      cmdSync.Caption = "Sync Missing Files/Folders Master to Slave"
    Case 1
      lstMissing.Visible = False
      lstMasterNew.Visible = True
      lstSlaveNew.Visible = False
      lstSizeVariations.Visible = False
      lstOrphans.Visible = False
      lblSubState = lstMasterNew.ListItems.Count & " items Newer on Master"
      cmdSync.Caption = "Sync Newer Files/Folders Master to Slave"
    Case 2
      lstMissing.Visible = False
      lstMasterNew.Visible = False
      lstSlaveNew.Visible = True
      lstSizeVariations.Visible = False
      lstOrphans.Visible = False
      lblSubState = lstSlaveNew.ListItems.Count & " items Newer on Slave"
      cmdSync.Caption = "Sync Newer Files/Folders Slave to Master"
    Case 3
      lstMissing.Visible = False
      lstMasterNew.Visible = False
      lstSlaveNew.Visible = False
      lstSizeVariations.Visible = True
      lstOrphans.Visible = False
      lblSubState = lstSizeVariations.ListItems.Count & " items have Size Variations"
      cmdSync.Caption = "Not implemented, manually resolve"
    Case 4
      lstMissing.Visible = False
      lstMasterNew.Visible = False
      lstSlaveNew.Visible = False
      lstSizeVariations.Visible = False
      lstOrphans.Visible = True
      lblSubState = lstOrphans.ListItems.Count & " items are on Slave only (orphans)"
      cmdSync.Caption = "Delete Orphan Files/Folders only on Slave"
    Case 5
      lstMissing.Visible = True
      lstMasterNew.Visible = False
      lstSlaveNew.Visible = False
      lstSizeVariations.Visible = False
      lstOrphans.Visible = False
      lblSubState = "Perform " & lstMissing.ListItems.Count + lstMasterNew.ListItems.Count + lstOrphans.ListItems.Count & " Total Actions"
      cmdSync.Caption = "Sync Missing, Newer Master, Delete Orphans"
    
  End Select

End Sub

Private Sub cmdBFFMaster_Click()

  On Error Resume Next
 'set object
  Set ShBFF = SH.BrowseForFolder(hWnd, "Select top level directory for Master drive or folder...", 1, txtMaster)
 'get folder selection
  txtMaster = ShBFF.Items.Item.Path
  
End Sub

Private Sub cmdBFFSlave_Click()

  On Error Resume Next
 'set object
  Set ShBFF = SH.BrowseForFolder(hWnd, "Select top level directory for Slave drive or folder...", 1, txtSlave)
  txtSlave = ShBFF.Items.Item.Path
  
End Sub

Private Sub SyncMissing()
    
  Dim i As Integer
  Dim iTicks As Long
  
  pbAll.Max = lstMissing.ListItems.Count
  pbAll.Value = 0
  pbAll.Visible = True
  Print #iLogFile, Now() & " ..................................Starting Missing Folders/Files update for " & lMissingItems & " items."
 'On Error Resume Next
 'Create folders
  For i = 1 To lstMissing.ListItems.Count  ' UBound(MasterItems)
    pbAll.Value = i
    If lstMissing.ListItems(i).ListSubItems(1) = "Folder" Then
      lblCurrentFolder = "Creating folder: " & txtSlave.Text & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1): DoEvents
      MLMkdir txtSlave & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1)
      Print #iLogFile, Now() & " Create folder: " & txtSlave.Text & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1)
   'End If
 'Next i
    Else
 'Copy file
 'For i = 0 To lstMissing.ListItems.Count  ' UBound(MasterItems)
   'pbAll.Value = i
   'If AllItems(i).ItemType = "File" Then
      lblCurrentFolder.Caption = "Copying " & Format$(FileLen(lstMissing.ListItems(i)), "##,###,###,##0") & " bytes: " & _
                                 lstMissing.ListItems(i) & " to " & txtSlave & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1)
      DoEvents
      On Error GoTo CannotCopy
      FileCopy lstMissing.ListItems(i), txtSlave & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1)
      Print #iLogFile, Now() & " Copy file: " & lstMissing.ListItems(i) & " to " & txtSlave & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1)
AfterBadCopy:
    End If
    If StopRequested Then Exit Sub
  Next i
  
  If StopRequested Then
    lblCurrentFolder = "Missing Files Synchronization Interrupted by User."
  Else
    lblCurrentFolder = "Missing Files Synchronization Complete."
  End If
  pbAll.Visible = False
  Exit Sub
  
CannotCopy:
  Print #iLogFile, Now() & " Error Copying following file: " & Err.Number & ": " & Err.Description
  Print #iLogFile, Now() & " Could not copy file: " & lstMissing.ListItems(i) & " to " & txtSlave & Mid(lstMissing.ListItems(i), Len(txtMaster) + 1)
  Resume AfterBadCopy
  
End Sub

Sub DoComboDependentSync()

  Select Case cboChoices.ListIndex
    Case 0
      lblMainState = "Sync Missing from Slave"
      cboChoices.Enabled = False
      SyncMissing
      cboChoices.Enabled = True
    Case 1
      lblMainState = "Sync Newer In Master than Slave"
      cboChoices.Enabled = False
      SyncNewerMaster
      cboChoices.Enabled = True
    Case 2
      lblMainState = "Sync Newer In Slave than Master"
      cboChoices.Enabled = False
      SyncNewerSlave
      cboChoices.Enabled = True
    Case 3
      
    Case 4
      lblMainState = "Remove Orphans from Slave folder."
      cboChoices.Enabled = False
      RemoveOrphans
      cboChoices.Enabled = True
    
    Case 5
     'This is a exploding call.  It performs three updates:
     '  1. Updates Missing files from Master to Slave
     '  2. Updates Newer files from Master to Slave
     '  3. Deletes Orphans from Slave
     'This is accomplished by forcing cboChoices to take on the thing we want to do, then recursing onto myself.
     cboChoices.Enabled = False  ' Keep idle hands away from it while running.
     cboChoices.ListIndex = 4
     DoComboDependentSync
     cboChoices.ListIndex = 0
     DoComboDependentSync
     cboChoices.ListIndex = 1
     DoComboDependentSync
     cboChoices.Enabled = True
     
    Case Else
      MsgBox "No sync/delete for choice " & cboChoices.ListIndex & " contact support for an update.  Of course, support is you, so feel free!"
  
  End Select
  
End Sub

Sub RemoveOrphans()

  Dim i As Integer
  Dim iAns As Integer
  
  Print #iLogFile, Now() & " ..................................Starting Orphan cleanup on Slave folder for " & lOrphans & " items."
 
 'Remove files first, then orphaned folders.
  On Error GoTo CannotDelete
  For i = 1 To lOrphans  ' UBound(MasterItems)
    If Orphans(i).ItemType = ITEM_FILE Then
      lblCurrentFolder.Caption = "Erasing file " & Orphans(i).ItemName: DoEvents
     'FSO.DeleteFile OrphanItems(i).ItemName, True
      iAns = GetAttr(Orphans(i).ItemName) And (255 - vbReadOnly - vbHidden)
      SetAttr Orphans(i).ItemName, iAns
      On Error GoTo CannotDelete
      Kill Orphans(i).ItemName
      Print #iLogFile, Now() & " Deleted Orphan file: " & Orphans(i).ItemName
BadDeleteReturn:
    End If
  Next i
  
  On Error Resume Next
 'Delete folders
 For i = 1 To lOrphans  ' UBound(MasterItems)
    If Orphans(i).ItemType = ITEM_FOLDER Then
      lblCurrentFolder.Caption = "Removing folder " & Orphans(i).ItemName
      If Not KillFolder(Orphans(i).ItemName) Then
       'MsgBox "Could not delete folder: " & Orphans(i).ItemName
        Print #iLogFile, Now() & " ***> Could not delete folder: " & Orphans(i).ItemName
      Else
        Print #iLogFile, Now() & " Deleted Orphan folder: " & Orphans(i).ItemName
      End If
     'FSO.DeleteFolder OrphanItems(i).ItemName, True
    End If
    If StopRequested Then Exit Sub
  Next i
  
  lblCurrentFolder = "Orphan Cleanup Complete."
  Print #iLogFile, Now() & " Orphan Cleanup finished."
  
  Exit Sub
  
CannotDelete:
  Print #iLogFile, Now() & " Cannot Delete file: " & Orphans(i).ItemName
  Debug.Print " Cannot Delete file: " & Orphans(i).ItemName
  Resume BadDeleteReturn
  
End Sub

Private Sub SyncNewerSlave()

 'Debug.Print "SyncNewerSlave Starting (This is just a stub at the moment, no useful code lives here.)"
  Print #iLogFile, Now() & " ..................................Starting Newer on Slave update (not implemented so it will finish quickly!)"
  If StopRequested Then Exit Sub

End Sub

Private Sub SyncNewerMaster()

  Dim i As Integer
  
  Print #iLogFile, Now() & " ..................................Starting Newer on Master update for " & lNewerMaster & " items."
  On Error Resume Next
  
 'Copy files which are newer on Master than Slave
  For i = 1 To lNewerMaster
    If NewerMaster(i).ItemType = ITEM_FILE Then
      lblCurrentFolder.Caption = "Copying newer " & NewerMaster(i).ItemName & " to " & txtSlave & Mid(NewerMaster(i).ItemName, Len(txtMaster) + 1)
      FileCopy NewerMaster(i).ItemName, txtSlave & Mid(NewerMaster(i).ItemName, Len(txtMaster) + 1)
      Print #iLogFile, Now() & " Copy file: " & NewerMaster(i).ItemName & " to " & txtSlave & Mid(NewerMaster(i).ItemName, Len(txtMaster) + 1)
    End If
    If StopRequested Then Exit Sub
  Next i
  
  lblCurrentFolder = "Newer Files Synchronization Complete."
  
End Sub
Private Sub cmdExit_Click()
  bCancelled = True 'cancel Dir() searching
  DoEvents
  Close
  Unload Me
 'End
End Sub

Private Sub cmdStop_Click()
  bCancelled = True 'cancel Dir() searching
End Sub

Private Sub Cmdpreview_Click()
  DrivePreview
End Sub

Sub DrivePreview()

  Dim varItem As Variant
  Dim lTime As Long
  Dim sFile As String  ' For Dir$ usage
  Dim i As Long
  Dim iMasterSize As Long
  Dim iSlaveSize As Long
  
  If txtMaster <> "" And txtSlave <> "" Then
  
    bCancelled = False
    lTime = GetTickCount
    
    lblMainState = "Initializing...": DoEvents
    Print #iLogFile, Now() & " ..................................Initializing Preview."
    InitializeForSearch  ' Disables a bunch of stuff
  
    lblCurrentFolder = txtMaster
    lblMainState = "Scanning Master folder files.": DoEvents
    Print #iLogFile, Now() & " Searching for Files/Folders on: " & txtMaster
    iScanPass = 0
    
    lblFoldQ.Visible = True
    lblFoldersQ.Visible = True
    
    Debug.Print "================== Scanning " & txtMaster
    Debug.Print Now()
    FindFilesAndFolders txtMaster  ' Gets all folders and files including subfolders with files and folders are seperately saved.
    Debug.Print Now()
    
    lblFoldQ.Visible = False
    lblFoldersQ.Visible = False
    Print #iLogFile, Now() & " Found " & Format$(lTotalFilesFound, "##,###,##0") & " files."
    Print #iLogFile, Now() & " Found " & Format$(lTotalFoldersFound, "##,###,##0") & " folders."
    
    frmManualRAID1.lblTotalFiles(iScanPass) = Format$(lTotalFilesFound, "##,###,##0")
    frmManualRAID1.lblTotalFolders(iScanPass) = Format$(lTotalFoldersFound, "##,###,##0")
    
   'All Master items have been found and stored.  Now check on missing and newer status and add to appropriate listview
    
    lblMainState = "Sorting Master folder files.": lblCurrentFolder = "": DoEvents
    
    cboChoices.Enabled = False
    lstMissing.Visible = False
    lstMasterNew.Visible = False
    lstSlaveNew.Visible = False
    lstSizeVariations.Visible = False
    lstOrphans.Visible = False
    Debug.Print "================== Sorting " & txtMaster
    Debug.Print Now()
    SortOnMissingDateTimeSize
    Debug.Print Now()
    
    Print #iLogFile, Now() & " Found " & Format$(lMissingItems, "##,###,##0") & " files/folders Missing on Slave."
    Print #iLogFile, Now() & " Found " & Format$(lNewerMaster, "##,###,##0") & " files/folders Newer on Master."
    Print #iLogFile, Now() & " Found " & Format$(lNewerSlave, "##,###,##0") & " files/folders Newer on Slave."
    If bCancelled Then
      InitializeForSearch
      ResetPgm
      Exit Sub
    End If
    
    Print #iLogFile, Now() & " Searching for Files/Folders on: " & txtSlave
    Debug.Print "================== Sorting " & txtSlave
    Debug.Print Now()
    SortOrphans  ' Non-recursively gets the full list from Slave and checks if they are on Master, all in this routine.
    Debug.Print Now()
    Print #iLogFile, Now() & " Found " & Format$(lOrphans, "##,###,##0") & " files/folders Orphaned on Slave (Not on Master)."
    cboChoices.Enabled = True
    cboChoices.ListIndex = 0
    If bCancelled Then
      InitializeForSearch
      ResetPgm
      Exit Sub
    End If
    
    frmManualRAID1.lblTotalFiles(iScanPass) = Format$(lTotalFilesFound, "##,###,##0")
    frmManualRAID1.lblTotalFolders(iScanPass) = Format$(lTotalFoldersFound, "##,###,##0")
    
    pbAll.Visible = False
    txtMaster.Enabled = True
    txtSlave.Enabled = True
    cmdPreview.Enabled = True
    cmdBFFMaster.Enabled = True
    cmdBFFSlave.Enabled = True
    cmdSync.Enabled = True
    
    lTime = (GetTickCount - lTime) / 1000
    
    lblMainState = "Scan Complete after " & TimeString(lTime, True) & "."
    
    ShowItemCountsInLists  ' Get the size of the selected LV shown.
    ReSizeLVs  ' Resize columns of all LVs.
    Debug.Print Now() & " All Done."
  End If
End Sub
Sub ResetPgm()

  cmdPreview.Enabled = True
  cmdBFFMaster.Enabled = True
  cmdBFFSlave.Enabled = True
  cmdSync.Enabled = False
  lblTotalFiles(0) = 0
  lblTotalFiles(1) = 0
  lblTotalFolders(0) = 0
  lblTotalFolders(0) = 0
  lblMainState = ""
  lblSubState = ""
  pbAll.Visible = False
  
End Sub
Sub SortOnMissingDateTimeSize()
  
  Dim itxtlength As Long
  Dim i As Long
  Dim iTempFiles As Long
  Dim iTempFolders As Long
  Dim sTemp As String
  Dim sFile As String
  Dim iMasterSize As Long
  Dim iSlaveSize As Long
  Dim iUpdateSeconds As Long
  
  timeBeginPeriod 1
  iUpdateSeconds = timeGetTime
  
  pbAll.Visible = True
  pbAll.Max = lAllItems
  
 'LockWindowUpdate True
  itxtlength = Len(txtMaster) + 1
  For i = 1 To lAllItems
   'Debug.Print AllItems(i).ItemName
    pbAll.Value = i
    If timeGetTime - iUpdateSeconds > MAX_UPDT_INT Then
      iUpdateSeconds = timeGetTime
      lblCurrentFolder = AllItems(i).ItemName
      lblSubState = Format$(i, "##,###,##0") & " items processed.   (" & Format(i / lAllItems, "percent") & ")"
      DoEvents
    End If
    If AllItems(i).ItemType = ITEM_FILE Then
      iTempFiles = iTempFiles + 1
    Else
      iTempFolders = iTempFolders + 1
    End If
    If AllItems(i).ItemType = ITEM_FOLDER Then
      sFile = Dir$(txtSlave & Mid$(AllItems(i).ItemName, itxtlength) & "\*.*", vbDirectory)
      If sFile = "" Then
        Set LstItem = lstMissing.ListItems.Add(, , AllItems(i).ItemName)
        LstItem.ListSubItems.Add , , "Folder"
        lMissingItems = lMissingItems + 1
        If lMissingItems > UBound(MissingItems) Then ReDim Preserve MissingItems(lMissingItems + ARRAY_EXPANSION)
        MissingItems(lMissingItems).ItemName = AllItems(i).ItemName
        MissingItems(lMissingItems).ItemType = ITEM_FOLDER
        lblCurrentFolder = MissingItems(lMissingItems).ItemName
      End If
    Else
      On Error GoTo BadFN
      sFile = Dir$(txtSlave & Mid$(AllItems(i).ItemName, itxtlength), vbNormal + vbHidden + vbReadOnly)
      If sFile = "" Then  ' If missing on Slave folder...
       'Debug.Print "File " & AllItems(i).ItemName & " does not exist on slave."
        Set LstItem = lstMissing.ListItems.Add(, , AllItems(i).ItemName)
        LstItem.ListSubItems.Add , , "File"
        lMissingItems = lMissingItems + 1
        If lMissingItems > UBound(MissingItems) Then ReDim Preserve MissingItems(lMissingItems + ARRAY_EXPANSION)
        MissingItems(lMissingItems).ItemName = AllItems(i).ItemName
        MissingItems(lMissingItems).ItemType = ITEM_FILE
      Else  ' File exists in both places.  Get info for each and sort that info into right buckets.
        Select Case DateDiff("s", FileDateTime(txtSlave & Mid$(AllItems(i).ItemName, itxtlength)), FileDateTime(AllItems(i).ItemName))
          Case Is > 0
            lNewerMaster = lNewerMaster + 1
            If lNewerMaster > UBound(NewerMaster) Then ReDim Preserve NewerMaster(lNewerMaster + ARRAY_EXPANSION)
            NewerMaster(lNewerMaster).ItemName = AllItems(i).ItemName
            NewerMaster(lNewerMaster).ItemType = ITEM_FILE
            Set LstItem = lstMasterNew.ListItems.Add(, , AllItems(i).ItemName)
            LstItem.ListSubItems.Add , , FileDateTime(AllItems(i).ItemName)
            LstItem.ListSubItems.Add , , FileDateTime(txtSlave & Mid$(AllItems(i).ItemName, itxtlength))
          Case Is < 0
            lNewerSlave = lNewerSlave + 1
            If lNewerSlave > UBound(NewerSlave) Then ReDim Preserve NewerSlave(lNewerSlave + ARRAY_EXPANSION)
            NewerSlave(lNewerSlave).ItemName = txtSlave & Mid$(AllItems(i).ItemName, Len(txtMaster))
            NewerSlave(lNewerSlave).ItemType = ITEM_FILE
            Set LstItem = lstSlaveNew.ListItems.Add(, , NewerSlave(lNewerSlave).ItemName)
            LstItem.ListSubItems.Add , , FileDateTime(AllItems(i).ItemName)
            LstItem.ListSubItems.Add , , FileDateTime(txtSlave & Mid$(AllItems(i).ItemName, itxtlength))
          Case Else
           'The files must be the same date/time.
        End Select
        
        iMasterSize = FileLen(AllItems(i).ItemName)
        iSlaveSize = FileLen(txtSlave & Mid$(AllItems(i).ItemName, itxtlength))
        If iMasterSize <> iSlaveSize Then
          Set LstItem = lstSizeVariations.ListItems.Add(, , AllItems(i).ItemName)
          LstItem.ListSubItems.Add , , Format(iMasterSize, "##,###,###,##0")
          LstItem.ListSubItems.Add , , Format(iSlaveSize, "##,###,###,##0")
        End If
      
      End If
    End If
    If bCancelled Then Exit Sub
SkipIt:
  Next
  timeEndPeriod 1
  Exit Sub
  
BadFN:
  Debug.Print "Bad Filename: " & txtSlave & Mid$(AllItems(i).ItemName, itxtlength)
  sFile = ""
  Resume SkipIt
  
End Sub

Sub SortOrphans()

  Dim i As Long
  Dim itxtlength As Long
  Dim sFile As String
  Dim iUpdateSeconds As Long
  
  pbAll.Visible = False
  
  lAllItems = 0  ' Let's start this all over again with some new information
  lTotalFilesFound = 0
  lTotalFoldersFound = 0
  lOrphans = 0
  lblSubState = ""
  
  If txtMaster <> "" And txtSlave <> "" Then
  
    lblMainState = "Finding Orphan Files/Folders in " & txtSlave: DoEvents
    
    lblCurrentFolder = txtSlave
    lblMainState = "Scanning Slave for Files/Folders.": DoEvents
    iScanPass = 1
    FindFilesAndFolders txtSlave  ' Gets all folders and files including sub\ but files and folders are seperate
   
  End If
  
  lblMainState = "Matching Slave files to Master to find Orphans.": DoEvents
  lblCurrentFolder = "": DoEvents
  
  If lAllItems = 0 Then Exit Sub
  
  pbAll.Max = lAllItems
  pbAll.Visible = True
  itxtlength = Len(txtSlave) + 1
  On Error GoTo BadFN
  timeBeginPeriod 1
  iUpdateSeconds = timeGetTime

  For i = 1 To lAllItems
    pbAll.Value = i
    If timeGetTime - iUpdateSeconds > MAX_UPDT_INT Then
      iUpdateSeconds = timeGetTime
      lblCurrentFolder = AllItems(i).ItemName
      lblSubState = Format$(i, "##,###,##0") & " items processed.   (" & Format(i / lAllItems, "percent") & ")"
      DoEvents
    End If
    If AllItems(i).ItemType = ITEM_FOLDER Then
      sFile = Dir$(txtMaster & Mid$(AllItems(i).ItemName, itxtlength) & "\*.*", vbDirectory)
      If sFile = "" Then  ' If null, file is not on master so this is an orphan to remove.
        lOrphans = lOrphans + 1
        If lOrphans > UBound(Orphans) Then ReDim Preserve Orphans(lOrphans + ARRAY_EXPANSION)
        Orphans(lOrphans).ItemName = AllItems(i).ItemName
        If bDoDebugPrints Then Debug.Print Orphans(lOrphans).ItemName
        Orphans(lOrphans).ItemType = ITEM_FOLDER
        Set LstItem = lstOrphans.ListItems.Add(, , AllItems(i).ItemName)
        LstItem.ListSubItems.Add , , "Folder"
      End If
    Else  ' Must be a file...
      sFile = Dir$(txtMaster & Mid$(AllItems(i).ItemName, itxtlength), vbNormal + vbHidden + vbReadOnly)
      If sFile = "" Then  ' If null, file is not on master so this is an orphan to remove.
        lOrphans = lOrphans + 1
        If lOrphans > UBound(Orphans) Then ReDim Preserve Orphans(lOrphans + ARRAY_EXPANSION)
        Orphans(lOrphans).ItemName = AllItems(i).ItemName
        Orphans(lOrphans).ItemType = ITEM_FILE
        Set LstItem = lstOrphans.ListItems.Add(, , AllItems(i).ItemName)
        LstItem.ListSubItems.Add , , "File"
      End If
    End If
    If bCancelled Then Exit Sub
SkipIt:
  Next
  pbAll.Visible = False
  timeEndPeriod 1
  Exit Sub
  
BadFN:
  Print #iLogFile, Now() & " Bad filename: " & txtMaster & Mid$(AllItems(i).ItemName, itxtlength)
  Resume SkipIt
  
End Sub
Sub InitializeForSearch()

  lblCurrentFolder = ""
  
  If Right(txtSlave.Text, 1) <> "\" Then txtSlave.Text = txtSlave.Text & "\"
  If Right(txtMaster.Text, 1) <> "\" Then txtMaster.Text = txtMaster.Text & "\"

  bCancelled = False
  
 'lblCurrentFolder.ForeColor = vbBlack
 'lblCurrentFolder.FontBold = False
 'lblCurrentFolder.Visible = True
  cmdPreview.Enabled = False
  cmdBFFMaster.Enabled = False
  cmdBFFSlave.Enabled = False
  cmdSync.Enabled = False
 
 'Let's start with something reasonable, aka TINY.
  ReDim AllItems(1) As ItemInfo
  ReDim MissingItems(1) As ItemInfo
  ReDim NewerMaster(1) As ItemInfo
  ReDim NewerSlave(1) As ItemInfo
  ReDim Orphans(1) As ItemInfo
  
  lAllItems = 0 ' Total of all files and folders.  They will be treated individually later on.
  lNewerSlave = 0
  lNewerMaster = 0   ' How many items in the NewerMaster array.
  lNewerSlave = 0   ' How many items in the NewerSlave array.
  
  lstMissing.ListItems.Clear
  lstMasterNew.ListItems.Clear
  lstSlaveNew.ListItems.Clear
  lstSizeVariations.ListItems.Clear
  lstOrphans.ListItems.Clear
    
  lblTotalFiles(0) = 0
  lblTotalFiles(1) = 0
  lblTotalFolders(0) = 0
  lblTotalFolders(1) = 0
  
  lstMissing.Visible = True
  cboChoices.Enabled = True
  
End Sub

Private Sub cmdSync_Click()
  DoComboDependentSync
End Sub

Private Sub Form_Load()
  
 'Test parm: -mpath c:\temp master -spath c:\temp slave -cmissing -cnewer -dorphans -sync
 
  Me.Caption = Me.Caption & " -- Vers. " & App.Major & "." & App.Minor & "." & App.Revision
  
  bDoDebugPrints = False
  
  Me.Left = GetSetting(App.EXEName, "Form", "Left", 0)
  Me.Top = GetSetting(App.EXEName, "Form", "Top", 0)
  Me.Width = GetSetting(App.EXEName, "Form", "Width", 10525)
  Me.Height = GetSetting(App.EXEName, "Form", "Height", 8040)
  If Me.Left < 0 Then Me.Left = 0
  If Me.Top < 0 Then Me.Top = 0
  If Me.Top + Me.Height > Screen.Height Then Me.Top = Screen.Height - Me.Height
  If Me.Left + Me.Width > Screen.Width Then Me.Left = Screen.Width - Me.Width
  
  Me.Show
  
  cboChoices.AddItem "Missing on Slave"
  cboChoices.AddItem "Newer on Master"
  cboChoices.AddItem "Newer on Slave"
  cboChoices.AddItem "Size Variations"
  cboChoices.AddItem "Orphans (on Slave only)"
  cboChoices.AddItem "Sync Missing, Newer on Master & Delete Orphans"
  bIgnoreClicks = True
  cboChoices.ListIndex = 0
  bIgnoreClicks = False
  
  bCancelled = False
  
  txtMaster = GetSetting(App.EXEName, "Parms", "Master", "C:\")
  txtSlave = GetSetting(App.EXEName, "Parms", "Slave", "C:\")
  
  iLogFile = FreeFile()
  Open App.Path & "\MR1Log " & Year(Now()) & "-" & Right("0" & Month(Now()), 2) & ".txt" For Append Access Write As #iLogFile
    
  If Not ParseCommandline() Then
    Unload Me
    Exit Sub
  End If
  
  If bAutomaticRun Then
    
    DrivePreview
    If bCopyMissing Then
      cboChoices.ListIndex = 0
      DoComboDependentSync
    End If
    
    If bCopyNewer Then
      cboChoices.ListIndex = 1
      DoComboDependentSync
    End If
    
    If bRemoveOrphans Then
      cboChoices.ListIndex = 4
      DoComboDependentSync
    End If
    
    Unload Me
    Exit Sub
  End If
  
  If txtMaster <> "" Then
    txtMaster.SelStart = 0
    txtMaster.SelLength = Len(txtMaster)
  End If
  
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

  SaveSetting App.EXEName, "Form", "Left", Me.Left
  SaveSetting App.EXEName, "Form", "Top", Me.Top
  SaveSetting App.EXEName, "Form", "Width", Me.Width
  SaveSetting App.EXEName, "Form", "Height", Me.Height
  SaveSetting App.EXEName, "Parms", "Master", txtMaster
  SaveSetting App.EXEName, "Parms", "Slave", txtSlave
  
End Sub

Private Sub Form_Resize()

  If Me.WindowState = vbMinimized Then Exit Sub
  
  If Me.Width < 11850 Then Me.Width = 11850
  If Me.Height < 6000 Then Me.Height = 6000
  
  lblCurrentFolder.Width = Me.Width - 400
  lblCurrentFolder.Top = Me.Height - 830
  
  cmdBFFMaster.Left = Me.Width - cmdBFFMaster.Width - 250
  cmdBFFSlave.Left = cmdBFFMaster.Left
  
  txtMaster.Width = cmdBFFMaster.Left - txtMaster.Left - 150
  txtSlave.Width = txtMaster.Width
  
  lblMainState.Left = Me.Width - lblMainState.Width - 150
  lblSubState.Left = lblMainState.Left
  
  pbAll.Width = lblMainState.Left - 150
  
  lstMissing.Left = 120
  lstMasterNew.Left = lstMissing.Left
  lstSlaveNew.Left = lstMissing.Left
  lstSizeVariations.Left = lstMissing.Left
  lstOrphans.Left = lstMissing.Left
  
  lstMissing.Top = 1970
  lstMasterNew.Top = lstMissing.Top
  lstSlaveNew.Top = lstMissing.Top
  lstSizeVariations.Top = lstMissing.Top
  lstOrphans.Top = lstMissing.Top
  
  lstMissing.Height = lblCurrentFolder.Top - lstMissing.Top - 40
  lstMasterNew.Height = lstMissing.Height
  lstSlaveNew.Height = lstMissing.Height
  lstSizeVariations.Height = lstMissing.Height
  lstOrphans.Height = lstMissing.Height
  
  lstMissing.Width = Me.Width - 270
  lstMasterNew.Width = lstMissing.Width
  lstSlaveNew.Width = lstMissing.Width
  lstSizeVariations.Width = lstMissing.Width
  lstOrphans.Width = lstMissing.Width
  
  ReSizeLVs
  
  txtSlave.Width = txtMaster.Width
  
End Sub
Sub ReSizeLVs()

  If lstMissing.ListItems.Count = 0 Then
    lstMissing.ColumnHeaders(1).Width = (lstMissing.Width - 180) * 0.94
    lstMissing.ColumnHeaders(2).Width = (lstMissing.Width - 180) * 0.04
  Else
    LV_AutoSizeColumn lstMissing
  End If
  
  If lstMasterNew.ListItems.Count = 0 Then
    lstMasterNew.ColumnHeaders(1).Width = lstMissing.Width * 0.75
    lstMasterNew.ColumnHeaders(2).Width = lstMissing.Width * 0.115
    lstMasterNew.ColumnHeaders(3).Width = lstMissing.Width * 0.115
  Else
    LV_AutoSizeColumn lstMasterNew
  End If

  If lstSlaveNew.ListItems.Count = 0 Then
    lstSlaveNew.ColumnHeaders(1).Width = lstMissing.Width * 0.75
    lstSlaveNew.ColumnHeaders(2).Width = lstMissing.Width * 0.115
    lstSlaveNew.ColumnHeaders(3).Width = lstMissing.Width * 0.115
  Else
    LV_AutoSizeColumn lstSlaveNew
  End If
   
  If lstSizeVariations.ListItems.Count = 0 Then
    lstSizeVariations.ColumnHeaders(1).Width = lstMissing.Width * 0.795
    lstSizeVariations.ColumnHeaders(2).Width = lstMissing.Width * 0.092
    lstSizeVariations.ColumnHeaders(3).Width = lstMissing.Width * 0.092
  Else
    LV_AutoSizeColumn lstSizeVariations
  End If

  If lstOrphans.ListItems.Count = 0 Then
    lstOrphans.ColumnHeaders(1).Width = lstOrphans.Width * 0.915
    lstOrphans.ColumnHeaders(2).Width = lstOrphans.Width * 0.06
  Else
    LV_AutoSizeColumn lstOrphans
  End If

End Sub
Function ParseCommandline() As Boolean

 Dim sCommand As String
 Dim aCommand() As String
 Dim i As Long
 Dim sTemp As String
 Dim iSpace As Long
 
 ParseCommandline = True
 
 'Command driven - here we go
 'Parms:
 '       -mp "complete Master path"  ' Complete path to Master directory (no quotes, please)
 '       -sp "complete Slave  path"  ' Complete path to Slave directory (no quotes, please)
 '       -cm(issing)                 ' Copy Missing files from Master to Slave
 '       -cn(ewer)                   ' Copy Newer in Master from Master to Slave
 '       -cb(ack)                    ' Copy Back newer in Slave to Master (Danger, Will Robinson!)
 '       -do(rphans)                 ' Finish Sync by Deleting Orphan files on Slave no longer on Master
 '       -sy(ncronize)               ' Do: (1)Copy Missing Master to Slave, (2)Copy Newer Master to Slave & (3)Remove Orphans on Slave
 '
 'Note:  There can be more than one of each of these.  The rightmost folder path of each kind wins.
 '       There should be at least one -cx and can be up to 4 for the four types of run.
 '
 '       The -mp and -sp will override what is automatically saved in the two text boxes.
 '
 'Sample input: <path>\FolderSyncronize -sf Working Directory -if Working Backup -cm -cn -do
 'Sample input: <path>\FolderSyncronize -sf Working Directory -if Working Backup -synronize
 '
 'Decoded: This example would Copy Missing T to I, Copy Newer on T to I and then Delete any files on I not still on T
 '
  sCommand = UCase$(Command$)
  If sCommand = "" Then Exit Function ' Manual operation this time.
  Print #iLogFile, Now() & "  Parsing command line options: " & sCommand
  If bDoDebugPrints Then Debug.Print "Parsing command line options: " & sCommand
  aCommand = Split(sCommand, "-")  ' Split input parms on dash (thereby removing the dash)
 'Set defaults to off.
  bCopyMissing = False
  bCopyNewer = False
  bCopyBack = False
  bRemoveOrphans = False
  bAutomaticRun = False
  
  For i = 1 To UBound(aCommand)
    Select Case Left(aCommand(i), 2)
      Case "MP"
        sMasterPath = Trim$(aCommand(i))
        iSpace = InStr(1, sMasterPath, " ")
        If iSpace = 0 Then GoTo ErrorParm
        sMasterPath = Mid$(sMasterPath, iSpace + 1)
        sTemp = Dir$(sMasterPath, vbDirectory)
        If sTemp = "" Then
          Print #iLogFile, Now(), " Master Path not found: " & sMasterPath
          ParseCommandline = False
          Exit Function
        End If
        txtMaster = sMasterPath
        bAutomaticRun = True
      Case "SP"
        sSlavePath = Trim$(aCommand(i))
        iSpace = InStr(1, sSlavePath, " ")
        If iSpace = 0 Then GoTo ErrorParm
        sSlavePath = Mid$(sSlavePath, iSpace + 1)
        sTemp = Dir$(sSlavePath, vbDirectory)
        If sTemp = "" Then
          Print #iLogFile, Now(), " Slave Path not found: " & sSlavePath
          ParseCommandline = False
          Exit Function
        End If
        txtSlave = sSlavePath
        bAutomaticRun = True
      Case "CM"
        bCopyMissing = True
        bAutomaticRun = True
      Case "CN"
        bCopyNewer = True
        bAutomaticRun = True
      Case "CB"
        bCopyBack = True
        bAutomaticRun = True
      Case "DO"
        bRemoveOrphans = True
        bAutomaticRun = True
      Case "SY"
        bCopyMissing = True
        bCopyNewer = True
        bRemoveOrphans = True
        bAutomaticRun = True
      Case Else
        GoTo ErrorParm
    End Select
  Next
  
  Exit Function
  
ErrorParm:
  Print #iLogFile, Now(), " There is a format error in the input command line.  Please revise and rerun."
  Print #iLogFile, Now(), " Command line with error: " & sCommand
  Unload Me
  ParseCommandline = False
  Exit Function
  End
  
End Function

Private Sub lstMasterNew_Click()

  Dim i As Long
  
  For i = 1 To lstMasterNew.ListItems.Count
    If lstMasterNew.ListItems(i).Selected Then
      lblCurrentFolder = "(Newer on Master) " & lstMasterNew.ListItems(i)
      Exit For
    End If
  Next

End Sub

Private Sub lstMissing_Click()
  
  Dim i As Long
  
  For i = 1 To lstMissing.ListItems.Count
    If lstMissing.ListItems(i).Selected Then
      lblCurrentFolder = "(" & lstMissing.ListItems(i).SubItems(1) & ") " & lstMissing.ListItems(i)
      Exit For
    End If
  Next

End Sub

Private Sub lstOrphans_Click()

  Dim i As Long
  
  For i = 1 To lstOrphans.ListItems.Count
    If lstOrphans.ListItems(i).Selected Then
      lblCurrentFolder = "(Only on Slave) " & lstOrphans.ListItems(i)
      Exit For
    End If
  Next

End Sub

Private Sub txtMaster_GotFocus()

  If txtMaster <> "" Then
    txtMaster.SelStart = 0
    txtMaster.SelLength = Len(txtMaster)
  End If

End Sub

Private Sub txtSlave_GotFocus()

  If txtSlave <> "" Then
    txtSlave.SelStart = 0
    txtSlave.SelLength = Len(txtSlave)
  End If

End Sub

Public Sub LV_AutoSizeColumn(LV As ListView, Optional Column As ColumnHeader = Nothing)
 
 Dim C As ColumnHeader
 
 If Column Is Nothing Then
   For Each C In LV.ColumnHeaders
     SendMessage LV.hWnd, LVM_FIRST + 30, C.Index - 1, -1
   Next
 Else
   SendMessage LV.hWnd, LVM_FIRST + 30, Column.Index - 1, -1
 End If
 
 LV.Refresh

End Sub

