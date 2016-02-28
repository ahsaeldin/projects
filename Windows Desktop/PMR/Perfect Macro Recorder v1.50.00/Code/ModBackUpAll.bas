Attribute VB_Name = "ModBackUpAll"
Option Explicit

'//ShowSelectFolder function Support''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'//cprin 26Jul2006: added next line in support of ShowSelectFolder function
Const MAX_PATH = 260

'//cprin 26Jul2006: added next type in support of ShowSelectFolder function
Private Type BrowseInfo
    hWndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

'//cprin 26Jul2006: added next line in support of ShowSelectFolder function
Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
'//cprin 26Jul2006: added next line in support of ShowSelectFolder function
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
'//cprin 26Jul2006: added next line in support of ShowSelectFolder function
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Function ShowSelectFolder(Owner As Long, Title As String, sPath As String) As Boolean
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Dim iNull As Integer
    Dim lpIDList As Long
    Dim udtBI As BrowseInfo
    
    With udtBI
        '//cprin 26Jul2006: Set the owner window
        .hWndOwner = Owner
        '//cprin 26Jul2006: lstrcat appends the two strings and returns the memory address
        .lpszTitle = lstrcat(Title, "")
    End With
    
    '//cprin 26Jul2006: Show the Browse for folder dialog
    lpIDList = SHBrowseForFolder(udtBI)
    If lpIDList Then
        sPath = String$(MAX_PATH, 0)
        '//cprin 26Jul2006: Get the path from the IDList
        SHGetPathFromIDList lpIDList, sPath
        '//cprin 26Jul2006: free the block of memory
        CoTaskMemFree lpIDList
        
        iNull = InStr(sPath, vbNullChar)
        If iNull Then
            sPath = Left$(sPath, iNull - 1)
        End If
        '//cprin 26Jul2006: //Return true if folder selected
        ShowSelectFolder = True
    Else
        '//cprin 26Jul2006: //Return false if no folder selected
        ShowSelectFolder = False
    End If
    
End Function

'//cprin 26Jul2006: added next function in support of backupall feature
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub BackupMacro(index As Integer, ByVal FolderToSave As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

On Error GoTo quitme:

Dim OF As OFSTRUCT
Dim IsThere As Long
Dim FileData As String
Dim MacFileData As String
Dim BackupFilePath As String
Dim SelectedMacroName As String
Dim SelectedDecMacroName As String

If frmJMacro.CobMacros.ListCount = 0 Then
   MsgBox "There is no macro to backup."
   Exit Sub
End If

SelectedMacroName = AppPath & "Data\" & frmJMacro.CobMacros.List(index) & ".pmf"
    
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'IsAlreadyChoosed = OpenFile(SelectedMacroName, OF, OF_READ)
          
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'CloseHandle IsAlreadyChoosed
            
'//cprin 29Jul2006: changed next line because if file name is very large, OpenFile return -1
'If IsAlreadyChoosed <> -1 Then
'//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
'//If PathFileExists(SelectedMacroName) Then
If FileExists(SelectedMacroName) Then
   '/cprin 17July2006: changed next MsgBox function to display a well form messagebox.
   'MsgBox "Data of [" & CobMacros.List(CobMacros.ListIndex) & "] macro missed or manually deleted by the user."
    'MsgBox "Can't Found the data file of " & CobMacros.List(CobMacros.ListIndex) & " macro" & Chr(13) & Chr(10) _
    '& "(" & SelectedMacroName & ")", vbCritical, App.Title
   'Exit Sub
End If

IsThere = 0

SelectedDecMacroName = AppPath & "Data\" & frmJMacro.CobMacros.List(index) & ".ptf"
          
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'IsAlreadyChoosed = OpenFile(SelectedMacroName, OF, OF_READ)
          
'//cprin 29Jul2006: comment next line because if file name is very large, OpenFile return -1
'CloseHandle IsAlreadyChoosed
            
'//cprin 29Jul2006: changed next line because if file name is very large, OpenFile return -1
'If IsAlreadyChoosed <> -1 Then
'//cprin 9Feb2007: changed next line because PathFileExists fails due to Craig=>operationivyflag@gmail.com report
'//If PathFileExists(SelectedMacroName) Then
If FileExists(SelectedMacroName) Then
     '/cprin 17July2006: changed next MsgBox function to display a well form messagebox.
    'MsgBox "Data of [" & CobMacros.List(CobMacros.ListIndex) & "] macro missed or manually deleted by the user."
    'MsgBox "Can't Found the data file of " & CobMacros.List(CobMacros.ListIndex) & " macro" & Chr(13) & Chr(10) _
    '& "(" & SelectedDecMacroName & ")", vbCritical, App.Title
    'Exit Sub
End If

If (right(FolderToSave, 1) <> "\") Then FolderToSave = FolderToSave & "\"

FolderToSave = FolderToSave & frmJMacro.CobMacros.List(index) & ".pbak"

BackupFilePath = FolderToSave

Open SelectedDecMacroName For Binary As #1
        
     FileData = Input(LOF(1), 1)
        
Close #1

FileData = FileData & "cprinsosonono" & Len(frmJMacro.CobMacros.List(index)) & frmJMacro.CobMacros.List(index)

Open SelectedMacroName For Binary As #2
        
    MacFileData = Input(LOF(2), 2)
        
Close #2
 
FileData = FileData & MacFileData
 
Open BackupFilePath For Binary As #3
        
    Put #3, , FileData
        
Close #3
 
Exit Sub
quitme:

End Sub

'//cprin 26Jul2006: added next function to backupall macros
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub BackupAllMacros(FolderToSave As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Dim i As Integer
    '//cprin 26Jul2006: loop thourgh CobMacros to BackupAllMacros
    For i = 0 To frmJMacro.CobMacros.ListCount - 1
        '//Backup Macro number i
        BackupMacro i, FolderToSave
    Next i
    
End Sub
'//cprin 27Jul2006: added next function to bandle all macros in one file
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub BandleAllMacros()
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
    Dim i As Integer
    Dim TempPath As String
    Dim TempMacPath As String
    Dim BandleFilePath As String
    Dim TempBackUpData As String
    
    '//cprin 27Jul2006: Get the path of the Temp Folder
    TempPath = GetTempPathName
         
    On Error Resume Next
    '//cprin 27Jul2006: Create erqwe54sdf54dd4534 if it doesn't exists yet.
    MkDir TempPath & "erqwe54sdf54dd4534"
    '//cprin 27Jul2006: delete all pervious macros if any.
    Kill TempPath & "erqwe54sdf54dd4534\*.*"  'erqwe54sdf54dd4534 is the folder name in which i choosed
    On Error GoTo 0
        
    '//cprin 27Jul2006: backup all macros to erqwe54sdf54dd4534 in order to start bundle
    BackupAllMacros TempPath & "erqwe54sdf54dd4534\"

    '//cprin 27Jul2006: Get the path in which the user want to save the bundle to
    With frmJMacro.CommonDialog1
        '//Default file name
        .FileName = "All " & SoftwareName & " Macros"
        '//Set Filter
        .Filter = SoftwareName & " backup File (*.pbak)|*.pbak"
        On Error GoTo quit:
        '//Set Cancel Error of the showsave
        .CancelError = True
        '//Can't Overwrite and can't enter invalid file name
        .flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
        '//Set the title of ShowSave
        .DialogTitle = "Backup All Macros in " & SoftwareName
        '//Show Save
        .ShowSave
    End With
    
    '//cprin 27Jul2006: Store thte path of the bandle file in a buffer.
    BandleFilePath = frmJMacro.CommonDialog1.FileName
    
    '//cprin 27Jul2006: Open the bandle file in order to write all macros files to.
    Open BandleFilePath For Binary As #1
        '// "cpringold79#%^&***&(@#$%ewrwear" is the header of the bundle followed by the number of the macros
        '// we use "pbak53333" as marker to get the number of macros how??
        '// for example the number of macros e.g. 10 will in between "cpringold79#%^&***&(@#$%ewrwear" and
        '// "pbak53333" hence we can easily get the number of macros.
        Put #1, , "cpringold79#%^&***&(@#$%ewrwear" & frmJMacro.CobMacros.ListCount & "pbak53333"
        For i = 0 To frmJMacro.CobMacros.ListCount - 1
            TempMacPath = TempPath & "erqwe54sdf54dd4534\" & frmJMacro.CobMacros.List(i) & ".pbak"
            '//cprin 26Jul2006: Load the data from the backup file into a buffer
            Open TempMacPath For Binary As #2
                '//cprin 26Jul2006: Load the data from the backup file into a buffer
                TempBackUpData = Input(LOF(2), 2)
            Close #2
            '//write the data to the bandle
            Put #1, , TempBackUpData
            '//"cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyy" & i is a marker "means this is the end of this macro"
            '//in other words after the end of  every file I will add this cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyyn   where n is the file number from the top
            Put #1, , "cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyy" & i
        Next i
    '//cprin 27Jul2006: Close the bandle file
    Close #1
    
quit:
    
End Sub
'//cprin 27Jul2006: added next function to import a bandle to PMR
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ImportBandleToPMR(BandleData As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Dim i As Integer
Dim FirstBound As Long
Dim NumbersCount As Long
Dim Wherepbak53333 As Long
Dim TempBackUpData As String
Dim WhereFirstMarker As Long
Dim NumberOfMacrosPos1 As Long
Dim MarkerBetweenFiles As String
Dim WhereMarkerBetweenFiles As Long
Dim OldMarkerBetweenFiles As String
Dim NumberOfMacrosInTheBandle As Long
Dim OldWhereMarkerBetweenFiles As Long
  
'//cprin 27Jul2006: First we will check if this is a bundle file that contains more than one macro
WhereFirstMarker = InStr(1, BandleData, "cpringold79#%^&***&(@#$%ewrwear")
'//Get the position of the seceond marker that bounds the number of the macros
Wherepbak53333 = InStr(1, BandleData, "pbak53333") - 1
'//Find out where is the first marker that bounds the number of macros from the left
NumberOfMacrosPos1 = Len("cpringold79#%^&***&(@#$%ewrwear")
'//Find out the number of digits in Macros number for ex: if 10 means == 2 digits
NumbersCount = Wherepbak53333 - NumberOfMacrosPos1
'//Get the Number of the macros in the bandle "at last heeeeee"
'Note if number of Macros 7 that's means we have 0,1,2,3,4,5,6 "it starts from 0 not 1"
NumberOfMacrosInTheBandle = Mid(BandleData, NumberOfMacrosPos1 + 1, NumbersCount)

'//Loop through the bandle to parse and load the data in order to import
For i = 0 To NumberOfMacrosInTheBandle - 1
    '//Create the Marker buffer that will will search for
    MarkerBetweenFiles = "cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyy" & i
    '//Search where is this marker in the bandle
    WhereMarkerBetweenFiles = InStr(1, BandleData, MarkerBetweenFiles)
    '//Speacial case if this is the first macro because this
    '//the bounds of this macro is the end of pbak53333 and "cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyy0"
    If i = 0 Then
       '//this is a plain logic, we need to parse the first macro that bounds between
       '//the end of pbak53333 and "cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyy0"
       TempBackUpData = Mid(BandleData, Wherepbak53333 + Len("pbak53333") + 1, WhereMarkerBetweenFiles - 1 - (Wherepbak53333 + Len("bak53333") + 1))
       '//Import this macro to PMR
       ImportSingleMacro (TempBackUpData)
    '//if any case rather than 0
    Else
       OldMarkerBetweenFiles = "cpringold&*(#$%^@#%^&*%(((%^&*yrvfyyy" & (i - 1)
       OldWhereMarkerBetweenFiles = InStr(1, BandleData, OldMarkerBetweenFiles)
       FirstBound = OldWhereMarkerBetweenFiles + Len(OldMarkerBetweenFiles) + 1
       TempBackUpData = Mid(BandleData, FirstBound, WhereMarkerBetweenFiles - FirstBound)
       '//Import this macro to PMR
       ImportSingleMacro (TempBackUpData)
    End If
Next i

End Sub
'//cprin 27Jul2006:* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Public Sub ImportSingleMacro(MacroData As String)
'//* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Dim i As Integer
Dim Res As Variant
Dim marker As Variant
Dim MacData As String
Dim DecData As String
Dim NameLen As Integer
Dim MacroName As String
Dim BackUpData As String
Dim DecFilePath As String
Dim MacFilePath As String
Dim Iscprinsosonono As String
 
BackUpData = MacroData

marker = InStr(1, BackUpData, "cprinsosonono")

If marker = 0 Then
   '//cprin 17July2006: changed next line to display a well form for MsgBox
   'MsgBox "invaild backup file."
   MsgBox "Invaild backup file.", vbCritical, SoftwareName
   Exit Sub
End If

On Error Resume Next
''''''''''''''''''''''''''''''''''''''Fix of NameLen exceeds one digits bug''''''''
'//cprin 29Jul2006: added next code block to fix the bug that appears if the namelen
'//is more than 1 digits for ex 10 or 100
Dim NextChar As String

NameLen = Mid(BackUpData, marker + 13, 1)

'//Since we doesn't allow the macro name to start with numbers hence
'//we will check if the next char is a number thus if it is a number then
'//name len exceeds 9
If IsNumeric(Mid(BackUpData, marker + 14, 1)) Then
     Dim NameLen2 As Integer
     NameLen2 = Mid(BackUpData, marker + 13, 2)
     '//Since we doesn't allow the macro name to start with numbers hence
     '//we will check if the next char is a number thus if it is a number then
     '//name len exceeds 99
     If IsNumeric(Mid(BackUpData, marker + 15, 1)) Then
        Dim NameLen3 As Integer
        NameLen3 = Mid(BackUpData, marker + 13, 3)
    End If
End If

If NameLen2 > NameLen Then
     NameLen = NameLen2
     marker = marker + 1   '//shift the marker one digit because we have another digit
End If

If NameLen3 > NameLen2 Then
     NameLen = NameLen3
     marker = marker + 1   '//shift the marker one digit because we have another digit
End If
'''''''''''''''''''''''''''''''''''''''''End of the fix'''''''''''''''''''''''''''''''
MacroName = Mid(BackUpData, marker + 14, NameLen)

For i = 0 To frmJMacro.CobMacros.ListCount - 1
   
       If MacroName = frmJMacro.CobMacros.List(i) Then
          '/cprin 17July2006: changed next line to display a well form messagebox
          'Res = MsgBox(MacroName & " macro already exists. do you want to overwrite it?", vbYesNo)
          Res = MsgBox("A macro with the name " & frmJMacro.CobMacros.List(i) & _
          " Already Exist, Do you want to overwrite it?", vbYesNo, App.Title)
          If Res = vbNo Then
             Exit Sub
          End If
       End If
   
Next i

MacFilePath = AppPath & "Data\" & MacroName & ".pmf"

DecFilePath = AppPath & "Data\" & MacroName & ".ptf"
   
On Error Resume Next
Kill MacFilePath
Kill DecFilePath
On Error GoTo 0

MacData = Mid(BackUpData, marker + 13 + NameLen + 1, Len(BackUpData) - marker + 13 + NameLen)

Open MacFilePath For Binary As #4
     
     Put #4, 1, MacData
     
Close #4
   
DecData = Left(BackUpData, marker - 1)


Open DecFilePath For Binary As #4
     
     Put #4, 1, DecData
     
Close #4
   
SaveSetting App.EXEName, "sm", MacroName, ""

frmJMacro.FillComboBox

End Sub


