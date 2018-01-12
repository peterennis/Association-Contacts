''''''''''''''''''''''''''''''''''''''''
' NAME: basFileSys
' VERSION: v1.0
'''''
' DEPENDENCIES
'   LIB PACKAGE: None
'   STANDALONE: None
'''''
' SOURCE: Jack D. Leach (www.dymeng.com)
'''''
' INCLUDED IN: Core 1.0
''''''''''''''''''''''''''''''''''''''''
Option Compare Database
Option Explicit

' V1.0
' V0.5 (added GetTempFile and TempDir)
' V0.4 (added CreateDirectory)
' V0.3 (added prelims of StringToFile and FileToString)
' V0.2 (added prelims of WebpageToString and DownloadBinary)
' V0.1

' GetOpenFile
' GetSaveFile
' BackupFolder
' CreateTempFile
' FileToString
' StringToFile
' GetOfficeExe
' GetFileTimes
' GetMappedURL
' GetDirContents
' IsReadOnly
' IsSystemFile
' IsHiddenFile

' CreateDirectory
' StringToFile (prelim)
' FileToString (prelim)
' WebpageToString (prelim)
' DownloadBinary (prelim)
' GetDiskSpace
' ConvertBytes
' IsDriveReady
' GetDriveType
' GetFileProperty
' BrowseFolder
' GetSpecialFolder
' FileLocked
' RemoveExtension
' ChangeExtension
' FileExists
' DirectoryExists
' GetFilename    'returns the filename from a full path
' GetPath        'returns the directory from a full path
' GetExtension   'returns the extension of a file
' TrailingSlash  'adds the trailing slash to a directory if required
' LeadingDot     'adds the prefix "." to an extension if required
' RemoveTrailingSlash  'Removes the trailing slash of a directory if required
' pfTruncateBytes
' pfConvertToBytes
' pfGetDiskFreeSpaceStruct 'returns a DISKFREESPACE structure (private)

' GetSpecialFolder functionality by Dev Ashish
' (see comments at the end of the module)

' Used by:
' pfGetDiskFreeSpaceStruct
' GetDiskFreeSpace
Private Type DISKFREESPACE
    lSectorsPerCluster As Long
    lBytesPerSector As Long
    lNumberOfFreeClusters As Long
    lTotalNumberOfClusters As Long
End Type

' Used by:
' ConvertBytes
' GetDiskFreeSpace
' pfBytesFromDFSS
Public Enum ByteUM
    umBytes = 0
    umKilobytes = 1 '=Bytes / 2^10   (ref: 2^10 = 1,024)
    umMegabytes = 2 '=Bytes / 2^20   (ref: 2^20 = 1,048,576)
    umGigabytes = 3 '=Bytes / 2^30   (ref: 2^30 = 1,073,741,824)
    umTerabytes = 4 '=Bytes / 2^40   (ref: 2^40 = 1,099,511,627,776)
End Enum

' Used by:
' GetDiskSpace
Public Enum DiskSpaceType
    DiskSpaceTotal = 0
    DiskSpaceUsed = 1
    DiskSpaceFree = 2
End Enum

' Used by:
' GetDriveType
Public Enum DriveType
    DriveTypeUnkown = 0
    DriveTypeNoRootDir = 1
    DriveTypeRemovable = 2
    DriveTypeFixed = 3
    DriveTypeRemote = 4
    DriveTypeCDROM = 5
    DriveTypeRAMDisk = 6
End Enum

Public Enum FileProperty
    FilePropHeaderInfo = -1
    FilePropName = 0
    FilePropSize = 1
    FilePropItemType = 2
    FilePropDateModified = 3
    FilePropDateCreated = 4
    FilePropDateAccessed = 5
    FilePropAttributes = 6
    FilePropPerceivedType = 9
    FilePropOwner = 10
    FilePropKind = 11
    FilePropDateTaken = 12
    FilePropRating = 19
    FilePropLength = 27
    FilePropBitRate = 28
    FilePropProtected = 29
    FilePropDimensions = 31
End Enum

Public Enum CSIDL
    CSIDL_ADMINTOOLS = &H30
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_APPDATA = &H1A
    CSIDL_BITBUCKET = &HA
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_COMMON_DESKTOPDIRECOTRY = &H19
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARMENU = &H16
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_CONNECTION = &H31
    CSIDL_CONTROLS = &H3
    CSIDL_COOKIES = &H21
    CSIDL_DESKTOP = &H0
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_FAVORITES = &H6
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_MASK = &HFF00&
    CSIDL_FLAT_PFTI_TRACKTARGET = CSIDL_FLAG_DONT_VERIFY
    CSIDL_FONTS = &H14
    CSIDL_HISTORY = &H22
    CSIDL_INTERNET = &H1
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_MYPICTURES = &H27
    CSIDL_NETHOOD = &H13
    CSIDL_NETWORK = &H12
    CSIDL_PERSONAL = &H5  ' My Documents
    CSIDL_MY_DOCUMENTS = &H5
    CSIDL_PRINTERS = &H4
    CSIDL_PRINTHOOD = &H1B
    CSIDL_PROFILE = &H28
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAMS = &H2
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_STARTMENU = &HB
    CSIDL_STARTUP = &H7
    CSIDL_SYSTEM = &H25
    CSIDL_SYSTEMX86 = &H29
    CSIDL_TEMPLATES = &H15
    CSIDL_WINDOWS = &H24
End Enum

' Used by BrowseFolder()
Public Enum BrowseFolderFlags
    BIF_RETURNONLYFSDIRS = &H1   'returns only filesystem directories
    BIF_DONTGOBELOWDOMAIN = &H2  'don't show network folders below the domain level
    BIF_RETURNFSANCESTORS = &H8  'return only ancestors of root folder
    BIF_BRROWSEFORCOMPUTER = &H1000
    BIF_BROWSEFORPRINTER = &H2000
End Enum

Private Const MAX_PATH = 260

' ------ The following declarations used by GetSpecialFolder()
'   Retrieves a pointer to the ITEMIDLIST structure of a special folder.
Private Declare PtrSafe Function apiSHGetSpecialFolderLocation Lib "shell32" _
    Alias "SHGetSpecialFolderLocation" _
    (ByVal hwndOwner As Long, _
    ByVal nFolder As Long, _
    ppidl As Long) _
    As Long
'   Converts an item identifier list to a file system path.
Private Declare PtrSafe Function apiSHGetPathFromIDList Lib "shell32" _
    Alias "SHGetPathFromIDList" _
    (pidl As Long, _
    ByVal pszPath As String) _
    As Long
'   Frees a block of task memory previously allocated through a call to
'   the CoTaskMemAlloc or CoTaskMemRealloc function.
Private Declare PtrSafe Sub sapiCoTaskMemFree Lib "ole32" _
    Alias "CoTaskMemFree" _
    (ByVal pv As Long)

' Used by:
' GetDriveType
Private Declare PtrSafe Function apiGetDriveType _
    Lib "kernel32" Alias "GetDriveTypeA" _
    (ByVal lpRootPathName As String) As Long

' Used by:
'   pfGetDiskFreeSpaceStruct
'   GetDiskSpace (indirect)
'   IsDriveReady (indirect)
Private Declare PtrSafe Function apiGetDiskFreeSpace _
    Lib "kernel32" Alias "GetDiskFreeSpaceA" _
    (ByVal lpRootPathName As String, _
    lpSectorsPerCluster As Long, _
    lpBytesPerSector As Long, _
    lpNumberOfFreeClusters As Long, _
    lpTotalNumberOfClusters As Long _
    ) As Long

' Used by:
' DownloadBinary
' WebpageToString
Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

' Used by TempDir() and GetTempFile()
' GetTempFileName
' GetTempPath
Private Declare PtrSafe Function apiGetTempPath _
    Lib "kernel32" Alias "GetTempPathA" _
    (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String _
    ) As Long
     
Private Declare PtrSafe Function apiGetTempFileName _
    Lib "kernel32" Alias "GetTempFileNameA" _
    (ByVal lpszPath As String, _
    ByVal lpPrefixString As String, _
    ByVal wUnique As Long, _
    ByVal lpTempFileName As String _
    ) As Long

' ************************************************************************************
'
'Public Function ExportToExcel(queryName As String) As Boolean
'On Error GoTo Err_Proc
''=========================
'  Dim OutputPath As String
'  Dim fd As FileOpenDialog
'  Dim ret As Boolean
''=========================
'
'  Set fd = New FileOpenDialog
'
'  With fd
'    .AllowCreation = True
'    .PromptOnCreate = False
'    .AddFilterItem "Excel Files (*.xls)", "*.xls"
'    .PathMustExist = True
'    .MultiSelect = False
'    .title = "Export to Excel"
'    OutputPath = .GetFile()
'  End With
'
'  If OutputPath = "" Then GoTo Exit_Proc
'  If FileSys.GetFilename(OutputPath) = "" Then OutputPath = OutputPath & ".xls"
'  If FileSys.GetExtension(OutputPath) <> ".xls" Then OutputPath = FileSys.ChangeExtension(OutputPath, ".xls")
'
'  DoCmd.OutputTo acOutputQuery, queryName, acFormatXLS, OutputPath, True, , , acExportQualityScreen
'
'  ret = True
'
''=========================
'Exit_Proc:
'  ExportToExcel = ret
'  Exit Function
'Err_Proc:
'  ret = False
'  Err.Source = "FileSys.ExportToExcel"
'  Select Case Err.Number
'    Case Else
'      HandleError
'  End Select
'  Resume Exit_Proc
'  Resume
'End Function

Public Function ArrayToFile(outputarray As Variant, Filename As String) As Boolean

    Dim s As String
    s = Join(outputarray, vbCrLf)
    basFileSys.StringToFile s, Filename

End Function

Public Function GetTempDirectory() As String

    Dim lngRet As Long
    Dim strTempDir As String
    Dim lngBuf As Long
  
    strTempDir = String$(255, 0)
    lngBuf = Len(strTempDir)
    lngRet = apiGetTempPath(lngBuf, strTempDir)
  
    If lngRet > lngBuf Then
        strTempDir = String$(lngRet, 0)
        lngBuf = Len(strTempDir)
        lngRet = apiGetTempPath(lngBuf, strTempDir)
    End If
  
    GetTempDirectory = Left(strTempDir, lngRet)

End Function

Public Function GetTempFile( _
    Optional CreateFile As Boolean = False, _
    Optional FilePrefix As String = "tmp", _
    Optional ByVal FilePath As String = "" _
    ) As String
    
    Dim lpTempFileName As String * 255
    Dim strTemp As String
    Dim lngRet As Long
  
    If Not basFileSys.DirectoryExists(FilePath) Then FilePath = basFileSys.GetTempDirectory()
  
    lngRet = apiGetTempFileName(FilePath, FilePrefix, 0, lpTempFileName)
  
    strTemp = lpTempFileName
  
    lngRet = InStr(lpTempFileName, Chr$(0))
    strTemp = Left(lpTempFileName, lngRet - 1)
  
    If Not CreateFile Then
        Kill strTemp
        Do Until Dir(strTemp) = "": DoEvents: Loop
    End If
  
    GetTempFile = strTemp
    
End Function

Public Function CreateDirectory( _
    ByVal Directory As String, _
    Optional CreateTree As Boolean = True _
    ) As Long
    On Error GoTo Err_Proc
    Dim ret As Long

    'creates a directory, optionally creating the tree as required
    'returns -1 on success, 0 on failure
  
    Dim v As Variant
    Dim i As Integer
    Dim s As String
  
    'replace all "/" with "\"
    Directory = Replace(Directory, "/", "\")
  
    'if we're on a server share replace the first two "\\" with pipes so it splits correctly
    If Left(Directory, 2) = "\\" Then Directory = "||" & Mid(Directory, 3)
  
    v = Split(Directory, "\")
  
    For i = 0 To UBound(v)
  
        'replace || with \\ if we're on a server
        If i = 0 Then If Left(Trim(CStr(v(i))), 2) = "||" Then v(i) = Replace(Trim(CStr(v(i))), "||", "\\")
  
        'build the string
        s = s & v(i) & "\"
    
        If Not CBool(basFileSys.DirectoryExists(s)) Then
            'dir doesn't exist, check if we're on the last entry
            If i < UBound(v) Then
                'we're in the middle of the tree... check if we have authorization to create the tree
                If Not CreateTree Then
                    ret = 0
                    GoTo Exit_Proc
                End If
            End If
      
            'craete the directory
            MkDir s
        End If
  
    Next i

    ret = -1

Exit_Proc:
    CreateDirectory = ret
    Exit Function
Err_Proc:
    ret = 0
    Resume Exit_Proc
End Function

Public Function StringToFile( _
    SourceString As String, _
    DestFile As String _
    )
    'prelim, no validations yet
    Dim i As Integer
  
    i = FreeFile
  
    Open DestFile For Append As #i
    Print #i, SourceString
    Close #i

End Function
    
Public Function FileToString(SourceFile As String) As String
    'prelim, no validations yet
    Dim i As Integer
    Dim s As String
  
    i = FreeFile
    Open SourceFile For Input As #i
    s = Input(LOF(i), i)
    Close #i
  
    FileToString = s
End Function

Public Function WebpageToString(ByVal srcURL As String) As String
    'return a string of the specified URL source
    'prelim version, no validations
    On Error Resume Next
    Dim xhr As Object
    Set xhr = CreateObject("Microsoft.XMLHTTP")
    xhr.Open "GET", srcURL, False
    xhr.send
    WebpageToString = xhr.responseText
    Set xhr = Nothing
    If Err.Number <> 0 Then
        WebpageToString = ""
    End If
End Function

Public Function DownloadBinary( _
    src As String, _
    dest As String, _
    Optional TimeoutMS As Long = 45000, _
    Optional ByRef Header As String _
    ) As Long
    ' Currently provides no validation for src or dest
    ' prelim version
    ' returns 0 on success
    ' returns -1 on timeout
    ' returns httpRequestStatus on other errors

    Const HTTPREQ_TIMEOUT_CHECK = 50

    Dim req As Object
    Dim lTimer As Long
    Dim bFlag As Boolean
    Dim bTimeout As Boolean
  
    Dim vBytes As Variant
    Dim bBytes() As Byte
  
    Dim iFile As Integer
  
    Set req = CreateObject("MSXML2.XMLHTTP.3.0")
    req.Open "GET", src, True
    req.send
  
    'timeout
    While bFlag = False
        DoEvents: DoEvents: DoEvents
        If req.ReadyState <> 4 Then
            'not done
            If lTimer >= TimeoutMS Then
                bFlag = True
                bTimeout = True
            End If
        Else
            bFlag = True
        End If
        Sleep HTTPREQ_TIMEOUT_CHECK
        lTimer = lTimer + HTTPREQ_TIMEOUT_CHECK
    Wend

    If bTimeout Then
        DownloadBinary = -1 'timeout
    Else
        If req.status = 200 Then
      
            Header = req.getAllResponseHeaders()
    
            vBytes = req.responseBody

            ReDim bBytes(0 To UBound(vBytes))
            bBytes = vBytes
      
            iFile = FreeFile()
            Open dest For Binary Access Write As #iFile
            Put #iFile, , bBytes
            Close #iFile
      
            DownloadBinary = 0
      
        Else
            DownloadBinary = req.status
        End If
    End If
  
    Set req = Nothing
  
End Function

Public Function GetDiskSpace( _
    ByVal Drive As String, _
    Optional SpaceType As DiskSpaceType = DiskSpaceTotal, _
    Optional UM As ByteUM = umBytes, _
    Optional DecimalPlaces As Integer = -1 _
    ) As Double

    Dim ret As Double
    Dim fs As DISKFREESPACE
  
    Dim Total As Double
    Dim Free As Double
    Dim Used As Double
  
    Drive = basFileSys.TrailingSlash(Drive)
  
    If basFileSys.IsDriveReady(Drive) Then
    
        fs = pfGetDriveFreeSpaceStruct(Drive)
    
        'build in two steps, avoids overflow
        Total = CDbl(fs.lBytesPerSector * fs.lSectorsPerCluster)
        Total = CDbl(Total * fs.lTotalNumberOfClusters)
    
        Free = CDbl(fs.lBytesPerSector * fs.lSectorsPerCluster)
        Free = CDbl(Free * fs.lNumberOfFreeClusters)
    
        Used = Total - Free
    
        Select Case SpaceType
    
            Case DiskSpaceTotal
                ret = ConvertBytes(Total, umBytes, UM, DecimalPlaces)
        
            Case DiskSpaceUsed
                ret = ConvertBytes(Used, umBytes, UM, DecimalPlaces)
        
            Case DiskSpaceFree
                ret = ConvertBytes(Free, umBytes, UM, DecimalPlaces)
        
        End Select
  
    End If
  
    GetDiskSpace = ret

End Function

Public Function ConvertBytes( _
    ValIn As Double, _
    UMIn As ByteUM, _
    UMOut As ByteUM, _
    Optional TruncateDecimalPlaces As Integer = -1 _
    ) As Double

    ' TruncateDecimalPlaces at -1 performs no truncation

    Dim ret As Double
    Dim Bytes As Double

    Bytes = pfConvertToBytes(ValIn, UMIn)

    Select Case UMOut
        Case umBytes
            ret = pfTruncateBytes(Bytes, TruncateDecimalPlaces)
        Case umKilobytes
            ret = pfTruncateBytes(Bytes / 2 ^ 10, TruncateDecimalPlaces)
      
        Case umMegabytes
            ret = pfTruncateBytes(Bytes / 2 ^ 20, TruncateDecimalPlaces)
      
        Case umGigabytes
            ret = pfTruncateBytes(Bytes / 2 ^ 30, TruncateDecimalPlaces)
      
        Case umTerabytes
            ret = pfTruncateBytes(Bytes / 2 ^ 40, TruncateDecimalPlaces)
      
    End Select

    ConvertBytes = ret

End Function

Public Function IsDriveReady(ByVal Drive As String) As Boolean
    'determines if a drive is readable, etc by attempting to
    'determine it's free space
    Dim apiret As Boolean
    
    Drive = basFileSys.TrailingSlash(Drive)
  
    apiret = apiGetDiskFreeSpace(Drive, 0&, 0&, 0&, 0&)
  
    IsDriveReady = CBool(apiret)
  
End Function

Public Function GetDriveType(ByVal Drive As String) As DriveType
    Dim apiret As Long
  
    Drive = basFileSys.TrailingSlash(Drive)
  
    apiret = apiGetDriveType(Drive)
  
    GetDriveType = apiret
  
End Function

Public Function GetFileProperty(FilePath As String, prop As FileProperty) As Variant
    'returns the specified property, or Null if no value
    Dim s As String
    With CreateObject("Shell.Application").Namespace(basFileSys.GetPath(FilePath))
        s = .GetDetailsOf(.ParseName(basFileSys.GetFilename(FilePath)), prop)
    End With
    If Len(s) = 0 Then
        GetFileProperty = Null
    Else
        GetFileProperty = s
    End If
End Function

Public Function GetSpecialFolder(ByVal ID As CSIDL) As String
    '   ********** Code Start **********
    ' This code was originally written by Dev Ashish.
    ' It is not to be altered or distributed,
    ' except as part of an application.
    ' You are free to use it in any application,
    ' provided the copyright notice is left unchanged.
    '
    ' Code Courtesy of
    ' Dev Ashish
    '
    '   Returns path to a special folder on the machine
    '   without a trailing backslash.
    '
    '   Refer to the comments in declarations for OS and
    '   IE dependent CSIDL values.
    '
    Dim lngRet As Long
    Dim strLocation As String
    Dim pidl As Long
    Const NOERROR = 0

    ' retrieve a PIDL for the specified location
    lngRet = apiSHGetSpecialFolderLocation(hWndAccessApp, ID, pidl)
    If lngRet = NOERROR Then
        strLocation = Space$(MAX_PATH)
        ' convert the pidl to a physical path
        lngRet = apiSHGetPathFromIDList(ByVal pidl, strLocation)
        If Not lngRet = 0 Then
            ' if successful, return the location
            GetSpecialFolder = Left$(strLocation, _
                InStr(strLocation, vbNullChar) - 1)
        End If
        ' calling application is responsible for freeing the allocated memory
        ' for pidl when calling SHGetSpecialFolderLocation. We have to
        ' call IMalloc::Release, but to get to IMalloc, a tlb is required.
        '
        ' According to Kraig Brockschmidt in Inside OLE,   CoTaskMemAlloc,
        ' CoTaskMemFree, and CoTaskMemRealloc take the same parameters
        ' as the interface functions and internally call CoGetMalloc, the
        ' appropriate IMalloc function, and then IMalloc::Release.
        Call sapiCoTaskMemFree(pidl)
    End If
End Function

Public Function BrowseFolder( _
    Optional StartDir As Variant, _
    Optional Caption As String = "Select a Folder", _
    Optional Flags As BrowseFolderFlags = BIF_RETURNONLYFSDIRS _
    ) As Variant
    ' 2011-08IR jl
    ' Returns Null on cancel

    Dim o As Object
    Dim Res As String
  
    Set o = CreateObject("Shell.Application").BrowseForFolder( _
        0, Caption, Flags, StartDir)
  
    On Error Resume Next
    Res = o.self.Path
  
    If Err.Number <> 0 Then
        BrowseFolder = Null
    Else
        BrowseFolder = Res
    End If
  
End Function

Public Function FileLocked(arg As Variant) As Variant
    ' 2011-08IR jl
    ' If arg is Null, Null is returned
    ' Returns -1 if the file cannot be opened
    ' Returns 1 if the file was not found
    ' Returns 0 otherwise
    Dim ret As Variant
    Dim s As String
    Dim i As Integer

    If IsNull(arg) Then
        ret = Null
        GoTo Exit_Proc
    End If

    If basFileSys.FileExists(arg, vbArchive + vbHidden + vbReadOnly + vbSystem) = False Then
        ret = CInt(1)
        GoTo Exit_Proc
    End If
  
    s = Trim(CStr(arg))
    i = FreeFile()
    ' Open parameters from Microsoft KB - Article Unknown
    On Error Resume Next
    Open s For Binary Access Read Write Lock Read Write As i
    Close #i
    If Err.Number = 0 Then
        ret = 0
    Else
        ret = -1
    End If
Exit_Proc:
    FileLocked = ret
End Function

Public Function ChangeExtension(arg As Variant, ext As Variant) As Variant
    ' 2011-08IR jl
    ' Changes the file's extension
    ' If arg is Null, return is Null
    ' If ext is Null no change
    ' If arg is empty string, ZLS returned regardless of ext
    ' If ext is empty string, arg extension
    ' is removed. If arg has no extension
    ' ext is added
    Dim ret As Variant
    Dim s As String
    Dim sOrigExt As String
  
    If IsNull(arg) Then
        ret = Null
        GoTo Exit_Proc
    End If
    If IsNull(ext) Then
        ret = CStr(arg)
        GoTo Exit_Proc
    End If
  
    s = Trim(CStr(arg))
  
    If Len(s) = 0 Then
        ret = ""
        GoTo Exit_Proc
    End If
  
    s = basFileSys.RemoveExtension(s)
    s = s & basFileSys.LeadingDot(ext)
    ret = s
  
Exit_Proc:
    ChangeExtension = ret
End Function

Public Function RemoveExtension(arg As Variant) As Variant
    ' 2011-08IR jl
    ' Removes the extension from a filename
    ' If arg is Null return is Null, otherwise return
    ' is all characters up to the last "." occurrence
    Dim ret As Variant
    Dim s As String
    Dim i As Integer
  
    If IsNull(arg) Then
        ret = Null
    Else
        s = Trim(CStr(arg))
        i = InStrRev(s, ".")
        If i = 0 Then
            ret = s
        Else
            ret = Left(s, i - 1)
        End If
    End If
    RemoveExtension = ret
End Function

Public Function FileExists(arg As Variant, Optional FileAttr As VbFileAttribute = VbFileAttribute.vbNormal) As Variant
    ' 2011-08IR jl
    ' Returns True if the arg is a valid file
    ' per the specified attribute.  If arg is
    ' Null return is Null.  If arg is empty string
    ' return is False.  FileAttribe usage examples:
    ' (hidden file)
    ' FileExists(HiddenFile, vbNormal) = False
    ' FileExists(HiddenFile, vbHidden) = True
    ' FileExists(NormalFile, vbNormal) = True
    ' FileExists(NormalFile, vbHidden) = True
    Dim ret As Variant
    Dim s As String
  
    If IsNull(arg) Then
        ret = Null
    Else
        s = RemoveTrailingSlash(Trim(CStr(arg)))
        If Len(s) = 0 Then
            ret = CBool(0)
        Else
            If Len(Dir(s, FileAttr)) = 0 Then
                ret = CBool(0)
            Else
                ret = CBool(-1)
            End If
        End If
    End If
    FileExists = ret
End Function

Public Function DirectoryExists(arg As Variant) As Variant
    ' 2011-08IR jl
    ' Returns True if arg is a valid directory
    ' If arg is Null, return is Null.  If arg
    ' is empty string, return is false.
    Dim ret As Variant
    Dim s As String
  
    If IsNull(arg) Then
        ret = Null
    Else
        s = RemoveTrailingSlash(Trim(CStr(arg)))
        If Len(s) = 0 Then
            ret = CBool(0)
        Else
            If Len(Dir(s, vbDirectory)) = 0 Then
                ret = CBool(0)
            Else
                ret = CBool(-1)
            End If
        End If
    End If
    DirectoryExists = ret
End Function

Public Function RemoveTrailingSlash(arg As Variant) As Variant
    ' 2011-08IR jl
    ' Returns the path without the trailing slash
    ' If arg is Null, return is Null.  If arg is
    ' and empty string, return is ZLS.
    Dim ret As Variant
    Dim s As String
  
    If IsNull(arg) Then
        ret = Null
    Else
        s = Trim(CStr(arg))
        If Len(s) = 0 Then
            ret = ""
        Else
            If Right(s, 1) = "\" Then
                ret = Left(s, Len(s) - 1)
            Else
                ret = s
            End If
        End If
    End If
    RemoveTrailingSlash = ret
End Function

Public Function GetFilename(arg As Variant) As Variant
    ' 2011-08IR jl
    ' 2012-02R1 jl
    ' Fixed function to work on both "/" and "\" (url handling)
    ' Converts all instances of "/" to "\"
    '
    ' Returns the Filename from a complete path.
    ' Returns variant (string) of all characters
    ' after the last occurring "\".  If no "\"
    ' is found, the entire string is returned.
    ' If arg is Null, Null is returned.  If arg
    ' is empty string, ZLS is returned.
    Dim ret As Variant
    Dim s As String
    If IsNull(arg) Then
        ret = Null
    Else
        s = Trim(CStr(arg))
        s = Replace(s, "/", "\")
        If Len(s) = 0 Then
            ret = ""
        Else
            If InStr(1, s, "\") = 0 Then
                ret = s
            Else
                ret = Mid(s, InStrRev(s, "\") + 1)
            End If
        End If
    End If
    GetFilename = ret
End Function

Public Function GetPath(arg As Variant) As Variant
    ' 2011-08IR jl
    ' Returns a variant (string) of all characters up
    ' to the last occurence of "\".  If arg is Null,
    ' Null is returned.  If arg is empty string, ZLS
    ' is returned.  If "\" is not found, ZLS is returned
    ' If Extension is not found but "\" is found, the
    ' entire string is returned.
    ' Ex: ?GetPath("C:\asdfa\asdfa")
    '     C:\asdfa\asdfa
    Dim ret As Variant
    Dim s As String
  
    If IsNull(arg) Then
        ret = Null
        GoTo Exit_Proc
    End If
  
    s = Trim(CStr(arg))
  
    If Len(s) = 0 Then
        ret = ""
        GoTo Exit_Proc
    End If
  
    If InStr(1, s, "\") = 0 Then
        ret = ""
        GoTo Exit_Proc
    End If
  
    If basFileSys.GetExtension(s) = "" Then
        ret = s
    Else
        ret = Left(s, InStrRev(s, "\") - 1)
    End If
  
Exit_Proc:
    GetPath = ret
End Function

Public Function GetExtension(arg As Variant) As Variant
    ' 2011-08IR jl
    ' returns a variant (string) of the last "." and characters
    ' if not "." is found or if "." found preceeding a "\" then
    ' ZLS is returned.  If arg is Null, Null is returned.  If
    ' arg is empty string, ZLS is returned.
    Dim ret As Variant
    Dim s As String
    If IsNull(arg) Then
        ret = Null
    Else
        s = Trim(CStr(arg))
        If Len(s) = 0 Then
            ret = ""
        Else
            If InStrRev(s, "\") > InStrRev(s, ".") Then
                ret = ""
            Else
                If InStrRev(s, ".") = 0 Then
                    ret = ""
                Else
                    ret = Mid(s, InStrRev(s, "."))
                End If
            End If
        End If
    End If
    GetExtension = ret
End Function

Public Function TrailingSlash(arg As Variant) As Variant
    ' 2011-08IR jl
    ' returns a variant (string) with the trailing slash
    ' added to the filepath.  If arg is Null, return is Null.
    ' If arg is empty string, ZLS is returned.
    Dim ret As Variant
    Dim s As String
    If IsNull(arg) Then
        ret = Null
    Else
        s = Trim(CStr(arg))
        If Len(s) = 0 Then
            ret = ""
        Else
            If Right(s, 1) <> "\" Then
                ret = s & "\"
            Else
                ret = s
            End If
        End If
    End If
    TrailingSlash = ret
End Function

Public Function LeadingDot(arg As Variant) As Variant
    ' 2011-08IR jl
    ' returns a variant (string) with the leading dot on the
    ' filepath.  If arg is Null, return is Null.  If arg
    ' is an empty string, return is ZLS.
    Dim ret As Variant
    Dim s As String
    If IsNull(arg) Then
        ret = Null
    Else
        s = Trim(CStr(arg))
        If Len(s) = 0 Then
            ret = ""
        Else
            If Left(s, 1) <> "." Then
                ret = "." & s
            Else
                ret = s
            End If
        End If
    End If
    LeadingDot = ret
End Function

Private Function pfTruncateBytes( _
    ByVal Val As Double, DecimalPlaces As Integer _
    ) As Double

    Dim s As String
    Dim ret As Double
    Dim decPos As Integer
  
    If DecimalPlaces = -1 Then
        ret = Val
    Else
    
        s = Trim(str(Val))
  
        decPos = InStr(1, s, ".")
    
        If decPos = 0 Then
            ret = Val
        Else
    
            'decimal place was found, continue
            s = Left(s, decPos) & Mid(s, decPos + 1, DecimalPlaces)
            ret = CDbl(s)
  
        End If
    End If
  
    pfTruncateBytes = ret

End Function

Private Function pfConvertToBytes( _
    dblIn As Double, _
    UMIn As ByteUM _
    ) As Double

    Dim ret As Double

    Select Case UMIn
    
        Case umBytes
            ret = dblIn
    
        Case umKilobytes
            ret = dblIn * 2 ^ 10
    
        Case umMegabytes
            ret = dblIn * 2 ^ 20
      
        Case umGigabytes
            ret = dblIn * 2 ^ 30
    
        Case umTerabytes
            ret = dblIn * 2 ^ 40

    End Select

    pfConvertToBytes = ret

End Function

Private Function pfGetDriveFreeSpaceStruct( _
    ByVal Drive As String _
    ) As DISKFREESPACE

    Dim apiret As Long
    Dim ret As DISKFREESPACE

    Dim lSpC As Long  ' sectors per cluster
    Dim lBpS As Long  ' bytes per sector
    Dim lNFC As Long  ' number of free clusters
    Dim lNTC As Long  ' number of total clusters

    Drive = basFileSys.TrailingSlash(Drive)

    apiret = apiGetDiskFreeSpace( _
        Drive, lSpC, lBpS, lNFC, lNTC)

    If apiret <> 0 Then
        With ret
            .lSectorsPerCluster = lSpC
            .lBytesPerSector = lBpS
            .lNumberOfFreeClusters = lNFC
            .lTotalNumberOfClusters = lNTC
        End With
    End If

    pfGetDriveFreeSpaceStruct = ret

End Function

' SOURCE:
' http://mvps.org/access/api/api0054.htm
' 2009/10/24
'   ********** Code Start **********
' This code was originally written by Dev Ashish.
' It is not to be altered or distributed,
' except as part of an application.
' You are free to use it in any application,
' provided the copyright notice is left unchanged.
'
' Code Courtesy of
' Dev Ashish
'
'   The following table outlines the different DLL versions,
'   and how they were distributed.
'
'   Version     DLL             Distribution Platform
'   4.00          All               Microsoft® Windows® 95/Windows NT® 4.0.
'   4.70          All               Microsoft® Internet Explorer 3.x.
'   4.71          All               Microsoft® Internet Explorer 4.0
'   4.72          All               Microsoft® Internet Explorer 4.01 and Windows® 98
'   5.00          Shlwapi.dll  Microsoft® Internet Explorer 5
'   5.00          Shell32.dll   Microsoft® Windows® 2000.
'   5.80          Comctl32.dll Microsoft® Internet Explorer 5
'   5.81          Comctl32.dll Microsoft® Windows 2000
'

'   © Microsoft. Information copied from Microsoft's
'   Platform SDK Documentation in MSDN
'   (http://msdn.microsoft.com)
'
'   If a special folder does not exist, you can force it to be
'   created by using the following special CSIDL:
'   (Version 5.0)
' Public Const CSIDL_FLAG_CREATE = &H8000

'   Combine this CSIDL with any of the CSIDLs listed below
'   to force the creation of the associated folder.

'   The remaining CSIDLs correspond to either file system or virtual folders.
'   Where the CSIDL identifies a file system folder, a commonly used path
'   is given as an example. Other paths may be used. Some CSIDLs can be
'   mapped to an equivalent %VariableName% environment variable.
'   CSIDLs are much more reliable, however, and should be used if at all possible.

'   File system directory that is used to store administrative tools for an individual user.
'   The Microsoft Management Console will save customized consoles to
'   this directory and it will roam with the user.
'   (Version 5.0)
' Public Const CSIDL_ADMINTOOLS = &H30
'
''   File system directory that corresponds to the user's
''   nonlocalized Startup program group.
' Public Const CSIDL_ALTSTARTUP = &H1D
'
''   File system directory that serves as a common repository for application-specific
''   data. A typical path is C:\Documents and Settings\username\Application Data.
''   This CSIDL is supported by the redistributable ShFolder.dll for systems that do
''   not have the Internet Explorer 4.0 integrated shell installed.
''   (Version 4.71)
' Public Const CSIDL_APPDATA = &H1A
'
''   Virtual folder containing the objects in the user's Recycle Bin.
' Public Const CSIDL_BITBUCKET = &HA
'
''   File system directory containing containing administrative tools
''   for all users of the computer.
''   Version 5
' Public Const CSIDL_COMMON_ADMINTOOLS = &H2F
'
''   File system directory that corresponds to the nonlocalized Startup program
''   group for all users. Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_ALTSTARTUP = &H1E
'
''   Application data for all users. A typical path is
''   C:\Documents and Settings\All Users\Application Data.
''   Version 5
' Public Const CSIDL_COMMON_APPDATA = &H23
'
''   File system directory that contains files and folders that appear on the
''   desktop for all users. A typical path is C:\Documents and Settings\All Users\Desktop.
''   Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_DESKTOPDIRECTORY = &H19
'
''   File system directory that contains documents that are common to all users.
''   A typical path is C:\Documents and Settings\All Users\Documents.
''   Valid for Windows NT® systems and Windows 95 and Windows 98
''   systems with Shfolder.dll installed.
' Public Const CSIDL_COMMON_DOCUMENTS = &H2E
'
''   File system directory that serves as a common repository for all users' favorite items.
''   Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_FAVORITES = &H1F
'
''   File system directory that contains the directories for the common program
''   groups that appear on the Start menu for all users. A typical path is
''   C:\Documents and Settings\All Users\Start Menu\Programs.
''   Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_PROGRAMS = &H17
'
''   File system directory that contains the programs and folders that appear on
''   the Start menu for all users. A typical path is
''   C:\Documents and Settings\All Users\Start Menu.
''   Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_STARTMENU = &H16
'
''   File system directory that contains the programs that appear in the
''   Startup folder for all users. A typical path is
''   C:\Documents and Settings\All Users\Start Menu\Programs\Startup.
''   Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_STARTUP = &H18
'
''   File system directory that contains the templates that are available to all users.
''   A typical path is C:\Documents and Settings\All Users\Templates.
''   Valid only for Windows NT® systems.
' Public Const CSIDL_COMMON_TEMPLATES = &H2D
'
''   Virtual folder containing icons for the Control Panel applications.
' Public Const CSIDL_CONTROLS = &H3
'
''   File system directory that serves as a common repository for Internet cookies.
''   A typical path is C:\Documents and Settings\username\Cookies.
' Public Const CSIDL_COOKIES = &H21
'
''   Windows Desktop—virtual folder that is the root of the namespace..
' Public Const CSIDL_DESKTOP = &H0
'
''   File system directory used to physically store file objects on the desktop
''   (not to be confused with the desktop folder itself).
''   A typical path is C:\Documents and Settings\username\Desktop
' Public Const CSIDL_DESKTOPDIRECTORY = &H10
'
''   My Computer—virtual folder containing everything on the local computer:
''   storage devices, printers, and Control Panel. The folder may
''   also contain mapped network drives.
' Public Const CSIDL_DRIVES = &H11
'
''   File system directory that serves as a common repository for the user's
''   favorite items. A typical path is C:\Documents and Settings\username\Favorites.
' Public Const CSIDL_FAVORITES = &H6
'
''   Virtual folder containing fonts. A typical path is C:\WINNT\Fonts.
' Public Const CSIDL_FONTS = &H14
'
''   File system directory that serves as a common repository for
''   Internet history items.
' Public Const CSIDL_HISTORY = &H22
'
''   Virtual folder representing the Internet.
' Public Const CSIDL_INTERNET = &H1
'
''   File system directory that serves as a common repository for
''   temporary Internet files. A typical path is
''   C:\Documents and Settings\username\Temporary Internet Files.
' Public Const CSIDL_INTERNET_CACHE = &H20
'
''   File system directory that serves as a data repository for local
''   (non-roaming) applications. A typical path is
''   C:\Documents and Settings\username\Local Settings\Application Data.
''   Version 5
' Public Const CSIDL_LOCAL_APPDATA = &H1C
'
''   My Pictures folder. A typical path is
''   C:\Documents and Settings\username\My Documents\My Pictures.
''   Version 5
' Public Const CSIDL_MYPICTURES = &H27
'
''   A file system folder containing the link objects that may exist in the
''   My Network Places virtual folder. It is not the same as CSIDL_NETWORK,
''   which represents the network namespace root. A typical path is
''   C:\Documents and Settings\username\NetHood.
' Public Const CSIDL_NETHOOD = &H13
'
''   Network Neighborhood—virtual folder representing the
''   root of the network namespace hierarchy.
' Public Const CSIDL_NETWORK = &H12
'
''   File system directory that serves as a common repository for documents.
''   A typical path is C:\Documents and Settings\username\My Documents.
' Public Const CSIDL_PERSONAL = &H5
'
''   Virtual folder containing installed printers.
' Public Const CSIDL_PRINTERS = &H4
'
''   File system directory that contains the link objects that may exist in the
''   Printers virtual folder. A typical path is
''   C:\Documents and Settings\username\PrintHood.
' Public Const CSIDL_PRINTHOOD = &H1B
'
''   User's profile folder.
''   Version 5
' Public Const CSIDL_PROFILE = &H28
'
''   Program Files folder. A typical path is C:\Program Files.
''   Version 5
'Public Const CSIDL_PROGRAM_FILES = &H2A
'
''   A folder for components that are shared across applications. A typical path
''   is C:\Program Files\Common.
''   Valid only for Windows NT® and Windows® 2000 systems.
''   Version 5
' Public Const CSIDL_PROGRAM_FILES_COMMON = &H2B
'
''   Program Files folder that is common to all users for x86 applications
''   on RISC systems. A typical path is C:\Program Files (x86)\Common.
''   Version 5
' Public Const CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
'
''   Program Files folder for x86 applications on RISC systems. Corresponds
''   to the %PROGRAMFILES(X86)% environment variable.
''   A typical path is C:\Program Files (x86).
''   Version 5
' Public Const CSIDL_PROGRAM_FILESX86 = &H2A
'
''   File system directory that contains the user's program groups (which are
''   also file system directories). A typical path is
''   C:\Documents and Settings\username\Start Menu\Programs.
' Public Const CSIDL_PROGRAMS = &H2
'
''   File system directory that contains the user's most recently used documents.
''   A typical path is C:\Documents and Settings\username\Recent.
''   To create a shortcut in this folder, use SHAddToRecentDocs. In addition to
''   creating the shortcut, this function updates the shell's list of recent documents
''   and adds the shortcut to the Documents submenu of the Start menu.
'Public Const CSIDL_RECENT = &H8
'
''   File system directory that contains Send To menu items. A typical path is
''   C:\Documents and Settings\username\SendTo.
' Public Const CSIDL_SENDTO = &H9
'
''   File system directory containing Start menu items.
''   A typical path is C:\Documents and Settings\username\Start Menu.
' Public Const CSIDL_STARTMENU = &HB
'
''   File system directory that corresponds to the user's Startup program group.
''   The system starts these programs whenever any user logs onto Windows NT® or
''   starts Windows® 95. A typical path is
''   C:\Documents and Settings\username\Start Menu\Programs\Startup.
' Public Const CSIDL_STARTUP = &H7
'
''   System folder. A typical path is C:\WINNT\SYSTEM32.
''   Version 5
' Public Const CSIDL_SYSTEM = &H25
'
''   System folder for x86 applications on RISC systems.
''   A typical path is C:\WINNT\SYS32X86.
''   Version 5
' Public Const CSIDL_SYSTEMX86 = &H29
'
''   File system directory that serves as a common repository
''   for document templates.
' Public Const CSIDL_TEMPLATES = &H15
'
''   Version 5.0. Windows directory or SYSROOT. This corresponds to the %windir%
''   or %SYSTEMROOT% environment variables. A typical path is C:\WINNT.
' Public Const CSIDL_WINDOWS = &H24