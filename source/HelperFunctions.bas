Attribute VB_Name = "HelperFunctions"
Private Declare PtrSafe Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Type SHFILEOPSTRUCT
    hWnd As LongPtr
    wFunc As LongPtr
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Boolean
    hNameMappings As LongPtr
    lpszProgressTitle As String
End Type

Private Const FO_DELETE = &H3
Private Const FOF_ALLOWUNDO = &H40
Private Const FOF_NOCONFIRMATION = &H10

Function DeleteFileToRecycleBin(strDeleteFile As String) As Boolean
    Dim shFileOp As SHFILEOPSTRUCT
    
    DeleteFileToRecycleBin = False

    If Not FileExists(strDeleteFile) Then
        Exit Function
    End If

    With shFileOp
        .wFunc = FO_DELETE
        .pFrom = strDeleteFile & vbNullChar & vbNullChar
        .fFlags = FOF_ALLOWUNDO Or FOF_NOCONFIRMATION
    End With
    
    SHFileOperation shFileOp
    DeleteFileToRecycleBin = True

End Function

Function FileExists(strFilePath As String) As Boolean
  Dim objFSO As Object

  ' Check if Microsoft Scripting Runtime is referenced
  On Error GoTo CheckRef
  Set objFSO = CreateObject("Scripting.FileSystemObject")

  ' Check if file exists
  If objFSO.FileExists(strFilePath) Then
    FileExists = True
  Else
    FileExists = False
  End If

  Exit Function

CheckRef:
  MsgBox ("To use this function, please follow these steps to add a reference to 'Microsoft Scripting Runtime' in your VBA project:" & vbCrLf & _
         "1. Go to the 'Tools' menu in the VBA editor." & vbCrLf & _
         "2. Select 'References.'" & vbCrLf & _
         "3. Check the box next to 'Microsoft Scripting Runtime.'" & vbCrLf & _
         "4. Click 'OK.'")
End Function



