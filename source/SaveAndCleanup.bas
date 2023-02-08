Attribute VB_Name = "SaveAndCleanup"
Function SaveAsLowerVersion(Optional cdrFile As String, Optional version As Long = 14) As Boolean
    
    Dim oCDRDoc As Document
    Dim oSaveOptions As StructSaveAsOptions ' https://community.coreldraw.com/sdk/api/draw/22/i/ivgstructsaveasoptions
    
    If Len(cdrFile) = 0 Then
        Set oCDRDoc = ActiveDocument
        cdrFile = oCDRDoc.FullFileName
    Else
        Set oCDRDoc = OpenDocument(cdrFile)
    End If
    
    ' Set the save options to save as CorelDRAW version 14 or other
    ' cdrVersion14
    Set oSaveOptions = CreateStructSaveAsOptions
    With oSaveOptions
        .version = version
        .EmbedVBAProject = False
        .EmbedICCProfile = False ' default
        .KeepAppearance = True ' default
    End With
    
    ' Save the document using the save options
    On Error GoTo saveFailed
        oCDRDoc.SaveAs cdrFile, oSaveOptions
        SaveAsLowerVersion = True
        
    Exit Function
    
saveFailed:
        SaveAsLowerVersion = False

End Function

Sub SaveAllAsLowerVersion()

    Dim oCDRDoc As Document
    Dim cdrFilePath As String
    Dim cdrFile As String
    Dim bSaveSuccsessFlag As Boolean

    Set oCDRDoc = ActiveDocument
    cdrFilePath = oCDRDoc.FilePath
    Set oCDRDoc = Nothing

    cdrFile = Dir(cdrFilePath & "*.cdr")

    Do While cdrFile <> ""
        If Not (cdrFile Like "Backup_of_*" Or cdrFile Like "Резервная_копия_*") Then
            bSaveSuccsessFlag = SaveAsLowerVersion(cdrFilePath & cdrFile, cdrVersion14)
            If bSaveSuccsessFlag Then
                ActiveDocument.Close
                bSaveSuccsessFlag = False
            End If
        End If
        cdrFile = Dir()
    Loop

End Sub

Sub SaveActiveDocAsLowerVersion()
    SaveAsLowerVersion
End Sub

Sub DeleteBackupFiles()
    Dim strPath As String
    Dim strFile As String
    Dim strDeleteFile As String

    strPath = ActiveDocument.FilePath
    strFile = Dir(strPath & "*.cdr")

    Do While strFile <> ""
        If strFile Like "Backup_of_*" Or strFile Like "Резервная_копия_*" Then
            strDeleteFile = strPath & strFile
            'Kill strDeleteFile
            DeleteFileToRecycleBin strDeleteFile
        End If
        strFile = Dir()
    Loop
End Sub


