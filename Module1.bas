Attribute VB_Name = "Module1"
'**************************************
'Windows API/Global Declarations for :AP
'     I AppendToLog
'**************************************
Const GENERIC_WRITE = &H40000000
Const FILE_SHARE_READ = &H1
Const CREATE_NEW = 1
Const OPEN_EXISTING = 3
Const FILE_ATTRIBUTE_NORMAL = &H80
Const FILE_BEGIN = 0
Const INVALID_HANDLE_VALUE = -1


Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long


Declare Function SetFilePointer Lib "kernel32" (ByVal hFile As Long, ByVal lDistanceToMove As Long, lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long


Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Any) As Long


Declare Function FlushFileBuffers Lib "kernel32" (ByVal hFile As Long) As Long


Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Function AppendToLog(ByVal lpFileName As String, ByVal sMessage As String) As Boolean
    'appends a string to a text file. it's u
    '     p to the coder to add a CR/LF at the end
    '
    'of the string if (s)he so desires.
    'assume failure
    AppendToLog = False
    
    'exit if the string cannot be written to
    '     disk
    If Len(sMessage) < 1 Then Exit Function
    
    'get the size of the file (if it exists)
    '
    Dim fLen As Long
    fLen = 0
    


    If (Len(Dir(lpFileName))) Then
        fLen = FileLen(lpFileName)
    End If
    
    'open the log file, create as necessary
    Dim hLogFile As Long
    hLogFile = CreateFile(lpFileName, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, _
    IIf(Len(Dir(lpFileName)), OPEN_EXISTING, CREATE_NEW), _
    FILE_ATTRIBUTE_NORMAL, 0&)
    
    'ensure the log file was opened properly
    '
    If (hLogFile = INVALID_HANDLE_VALUE) Then Exit Function
    
    'move file pointer to end of file if fil
    '     e was not created


    If (fLen <> 0) Then


        If (SetFilePointer(hLogFile, fLen, ByVal 0&, FILE_BEGIN) = &HFFFFFFFF) Then
            'exit sub if the pointer did not set cor
            '     rectly
            CloseHandle (hLogFile)
            Exit Function
        End If
    End If
    
    'convert the source string to a byte arr
    '     ay for use with WriteFile
    Dim lTemp As Long
    ReDim TempArray(0 To Len(sMessage) - 1) As Byte
    


    For lTemp = 1 To Len(sMessage)
        TempArray(lTemp - 1) = Asc(Mid$(sMessage, lTemp, 1))
    Next
    
    'write the string to the log file


    If (WriteFile(hLogFile, TempArray(0), Len(sMessage), lTemp, ByVal 0&) <> 0) Then
        'the data was written correctly
        AppendToLog = True
    End If
    
    'flush buffers and close the file
    FlushFileBuffers (hLogFile)
    CloseHandle (hLogFile)
    
End Function

