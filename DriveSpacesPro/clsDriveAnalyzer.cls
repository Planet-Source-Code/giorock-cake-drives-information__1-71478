VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsDriveAnalyzer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'********************************************
'* Thanks to Kenneth Ives  kenaso@tx.rr.com *
'*              for Media Type              *
'********************************************
'*            Thanks to NightWolf           *
'*              for Drive Info              *
'********************************************
'*         Assembled by GioRock 2008        *
'********************************************

Private Declare Function GetDriveTypeA Lib "kernel32" (ByVal Drive As String) As Long 'Gets the type of the specified file
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long 'Retrieves all available drives
Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long 'Retrieves Drive Information
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long 'Retrieves Drive Space Information
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long 'Closes an opened handle

Private Declare Function GetLastError Lib "kernel32" () As Long 'Retrieves last error information

Private Const DEF_ATTR As Integer = 55 'vbReadOnly + vbHidden + vbSystem + vbDirectory + vbArchive

Public Enum DriveTypeValues
    DTV_Unknown
    DTV_NotFound
    DTV_Removable
    DTV_Fixed
    DTV_Remote
    DTV_CDROM
    DTV_RAMDisk
End Enum

'Add by GioRock
Private Type OSVERSIONINFO
    OSVSize        As Long
    dwVerMajor     As Long
    dwVerMinor     As Long
    dwBuildNumber  As Long
    PlatformID     As Long
    szCSDVersion   As String * 128
End Type
Private Declare Function GetVersionEx Lib "kernel32" _
        Alias "GetVersionExA" (LpVersionInformation As Any) As Long
Private Const VER_PLATFORM_WIN32_NT             As Long = 2
Private Const GENERIC_READ                      As Long = &H80000000
Private Const GENERIC_WRITE As Long = &H40000000
Private Const FILE_SHARE_READ As Long = &H1
Private Const FILE_SHARE_WRITE As Long = &H2
Private Const OPEN_EXISTING As Long = 3

Private Type DEVICE_MEDIA_INFO
    Cylinders             As Double
    MediaType             As Long
    TracksPerCylinder     As Long
    SectorsPerTrack       As Long
    BytesPerSector        As Long
    NumberMediaSides      As Long
    MediaCharacteristics  As Long
End Type

Private Type GET_MEDIA_TYPES
    DeviceType     As Long
    MediaInfoCount As Long
    MediaInfo(10)  As DEVICE_MEDIA_INFO
End Type

Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
        (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, _
        ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, _
        ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, _
        ByVal hTemplateFile As Long) As Long

Private Const INVALID_HANDLE_VALUE              As Long = -1

Private Declare Function DeviceIoControl Lib "kernel32" _
        (ByVal hDevice As Long, ByVal dwIoControlCode As Long, _
        lpInBuffer As Any, ByVal nInBufferSize As Long, _
        lpOutBuffer As Any, ByVal nOutBufferSize As Long, _
        lplngBytesRet As Long, lpOverlapped As Any) As Long
Private Const IOCTL_STORAGE_GET_MEDIA_TYPES_EX  As Long = &H2D0C04
Private Const FILE_DEVICE_CD_ROM                As Long = &H2
Private Const FILE_DEVICE_DVD                   As Long = &H33

Private m_objLastError As ErrObject 'Holds information regarding the last ocurring error

Public Function UnQualifyPath(ByVal strPath As String) As String

  'removes any trailing slash from the path
   strPath = Trim$(strPath)
   
   If Right$(strPath, 1) = "\" Then
       UnQualifyPath = Left$(strPath, Len(strPath) - 1)
   Else
       UnQualifyPath = strPath
   End If
   
End Function

Private Function IsWin2kPlus() As Boolean

    ' Called by GetDriveInfo()
    '           GetMediaInfo()
    '           IsDeviceReady()
    '           DeviceLock()
    '           GetDiskGeometry()
    
    Dim typOSVI As OSVERSIONINFO
    
    typOSVI.OSVSize = Len(typOSVI)
    
    If GetVersionEx(typOSVI) = 1 Then
    
        IsWin2kPlus = (typOSVI.PlatformID = VER_PLATFORM_WIN32_NT) And _
                      (typOSVI.dwVerMajor = 5) And _
                      (typOSVI.dwVerMinor >= 0)
    End If

End Function
Private Function GetMediaType(ByVal strDrive As String) As String

    Dim hHandle      As Long
    Dim lngStatus    As Long
    Dim lngReturn    As Long
    Dim lngNullValue As Long
    Dim typGMT       As GET_MEDIA_TYPES

    On Error GoTo GetMediaType_Error

    GetMediaType = ""
    
    ' test for operating system of Windows 2K or newer
    If IsWin2kPlus Then
    
        strDrive = UnQualifyPath(strDrive)
        
        ' get a handle to the drive
        hHandle = CreateFile("\\.\" & UCase$(strDrive), _
                           GENERIC_READ Or GENERIC_WRITE, _
                           FILE_SHARE_READ Or FILE_SHARE_WRITE, _
                           lngNullValue, OPEN_EXISTING, 0, lngNullValue)
        
        If hHandle <> INVALID_HANDLE_VALUE Then
        
           lngStatus = DeviceIoControl(hHandle, _
                                       IOCTL_STORAGE_GET_MEDIA_TYPES_EX, _
                                       lngNullValue, 0, typGMT, _
                                       2048, lngReturn, ByVal 0)
        
           If lngStatus <> 0 Then
               GetMediaType = typGMT.DeviceType
           End If
        
        End If
        
    Else
        GetMediaType = "This application is not designed to execute " & _
                "on operating systems earlier than Windows 2000."
    End If

GetMediaType_CleanUp:
    CloseHandle hHandle
    On Error GoTo 0
    Exit Function

GetMediaType_Error:
    MsgBox Err.Description, vbCritical, "FileAnalizer - GetMediaType"
    Resume GetMediaType_CleanUp

End Function

Public Function ParseSize(ByVal Size As Currency, Optional RoundBy As Integer = 2) As String

'Store the size of each unit in bytes to save up calculation time

Const KB_LEN As Long = &H400 'Number of bytes needed to make a KiloByte
Const MB_LEN As Long = &H100000 'Number of bytes needed to make a MegaByte
Const GB_LEN As Long = &H40000000 'Number of bytes needed to make a GigaByte

    If Size < KB_LEN Then 'If the file is smaller than one KB

        ParseSize = Round(Size, RoundBy) & " Bytes" 'Output it's size in bytes

    ElseIf Size < MB_LEN Then 'If the file is smaller than one MB

        ParseSize = Round(Size / KB_LEN, RoundBy) & " KB" 'Output it's size in KBs

    ElseIf Size < GB_LEN Then 'If the file is smaller than one GB

        ParseSize = Round(Size / MB_LEN, RoundBy) & " MB" 'Output it's size in MBs

    Else 'If the file is bigger than than one GB

        ParseSize = Round(Size / GB_LEN, RoundBy) & " GB" 'Output it's size in GBs

    End If

End Function

Public Function GetDriveTypeName(ByVal Drive As String) As String
Dim sCD_DVD As String

    Select Case GetDriveType(Drive) 'Check the type of the specified drive

        Case DTV_Unknown: GetDriveTypeName = "Unknown" 'Drive Type is Unknown
        Case DTV_NotFound: GetDriveTypeName = "Not Found" 'Drive was not found
        Case DTV_Removable: If Left$(Drive, 1) = "A" Or Left$(Drive, 1) = "B" Then GetDriveTypeName = "Floppy" Else GetDriveTypeName = "USB Drive" 'Drive is either Removable or a Floppy Drive
        Case DTV_Fixed: GetDriveTypeName = "Fixed" 'Drive is Fixed
        Case DTV_Remote: GetDriveTypeName = "Remote" 'Drive is Remote
        Case DTV_CDROM
            sCD_DVD = GetMediaType(Drive)
            If sCD_DVD = FILE_DEVICE_CD_ROM Then
                GetDriveTypeName = "CD-ROM" 'Drive is a CD-ROM
            End If
            If sCD_DVD = FILE_DEVICE_DVD Then
                GetDriveTypeName = "DVD-ROM" 'Drive is a CD-ROM
            End If
        Case DTV_RAMDisk: GetDriveTypeName = "RAM Disk" 'Drive is a RAM Disk

    End Select

End Function

Public Function GetDriveType(ByVal Drive As String) As Integer

    GetDriveType = GetDriveTypeA(Drive) 'Retrieves the specified drive's type

End Function

Public Function Exists(ByVal Path As String) As Boolean

    On Error Resume Next

    Exists = (Dir$(Path, DEF_ATTR) <> vbNullString) 'Checks if the file exists

End Function

Public Function GetDrives(Drives() As String) As Integer

Dim strBuffer As String

    strBuffer = String(255, vbNullChar) 'Create space in the buffer

    Call GetLogicalDriveStrings(255, strBuffer) 'Get the available drives in the buffer

    strBuffer = Left$(strBuffer, InStrRev(strBuffer, "\") + 1) 'Strip null terminators

    Drives() = Split(strBuffer, vbNullChar) 'Add the available drives to an array

    ReDim Preserve Drives(UBound(Drives) - 1) 'Strip off last empty item

End Function
Public Function GetDriveInfo(ByVal Drive As String, Optional ByRef Volume As String, Optional ByRef FileSystem As String, Optional ByRef SerialNumber As String) As Integer

    On Error GoTo GetDriveInfo_Err

Dim lngSerial As Long

    Volume = String$(255, vbNullChar): FileSystem = String$(255, vbNullChar) 'Fill the buffers with null characters

    GetDriveInfo = GetVolumeInformation(Drive, Volume, Len(Volume), lngSerial, 0, 0, FileSystem, Len(FileSystem)) 'Retrieve volume information

    If GetDriveInfo = 0 Then GetDriveInfo = GetLastError: Exit Function Else GetDriveInfo = 0 'Check if any errors occurred

    Volume = Left$(Volume, InStr(Volume, vbNullChar) - 1) 'Retrieve the volume name

    FileSystem = Left$(FileSystem, InStr(FileSystem, vbNullChar) - 1) 'Retrieve the drive's file system

    ' by GioRock
    SerialNumber = IIf(lngSerial <> 0, Format$(String$(8 - Len(Hex$(lngSerial)), "0") + Hex(lngSerial), "@@@@-@@@@"), "No SN")   'Retrieve the drive's serial number

    Exit Function

GetDriveInfo_Err:     GetDriveInfo = Err.Number: Set m_objLastError = Err

End Function

Public Function GetDriveSpace(ByVal Drive As String, Optional ByRef TotalSpace As Currency, Optional ByRef FreeSpace As Currency, Optional ByRef UsedSpace As Currency, Optional ByRef UsedSpacePercent As Single) As Integer

    On Error GoTo GetDriveSpace_Err

    GetDriveSpace = GetDiskFreeSpaceEx(Drive, 0&, TotalSpace, FreeSpace) 'Retrieve drive size information

    If GetDriveSpace = 0 Then GetDriveSpace = GetLastError: Exit Function Else GetDriveSpace = 0 'Check if any errors occurred

    'Convert the results propertly to a currency value

    TotalSpace = TotalSpace * 10000
    FreeSpace = FreeSpace * 10000

    'Calculate used space and percent

    UsedSpace = TotalSpace - FreeSpace
    UsedSpacePercent = Round((UsedSpace / TotalSpace) * 100, 2)

    Exit Function

GetDriveSpace_Err:     GetDriveSpace = Err.Number: Set m_objLastError = Err

End Function

