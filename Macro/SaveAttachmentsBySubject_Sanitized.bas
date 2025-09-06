Public Sub SaveAttachmentsBySubject(MItem As Outlook.MailItem)
    Dim oAttachment As Outlook.Attachment
    Dim configData As Collection
    Dim configItem As Collection
    Dim logFilePath As String
    Dim logText As String
    Dim logFile As Integer
    Dim attachmentSaved As Boolean
    Dim recipientEmailAddress As String
    Dim emailSubject As String
    Dim matchFound As Boolean
    Dim destinationFolder As String
    Dim backupFolder As String
    Dim destinationFileName As String
    Dim primaryFilePath As String
    Dim backupFilePath As String
    Dim macroFolder As String
    Dim configFilePath As String
    
    ' Get the macro folder path using secure relative path resolution
    macroFolder = GetSecureMacroPath()
    If macroFolder = "" Then
        LogSecurityEvent "CRITICAL: Unable to determine secure macro path"
        Exit Sub
    End If
    
    configFilePath = macroFolder & "subject_config.csv"
    logFilePath = macroFolder & "Logs\AttachmentLog.txt"
    
    ' Create Logs folder if it doesn't exist (with validation)
    If Not CreateSecureFolder(macroFolder & "Logs") Then
        LogSecurityEvent "ERROR: Unable to create logs folder"
        Exit Sub
    End If
    
    ' Initialize variables
    logText = ""
    attachmentSaved = False
    emailSubject = SanitizeInput(MItem.Subject)
    
    ' Validate email subject
    If Not IsValidEmailSubject(emailSubject) Then
        LogSecurityEvent "WARNING: Invalid email subject detected"
        Exit Sub
    End If
    
    ' Load configuration from CSV with validation
    logText = logText & "Loading configuration..." & vbCrLf
    Set configData = LoadSubjectConfigSecure(configFilePath)
    If configData Is Nothing Then
        logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - ERROR: Could not load configuration file" & vbCrLf
        GoTo WriteLog
    Else
        logText = logText & "Configuration loaded successfully. Found " & configData.Count & " rules." & vbCrLf
    End If
    
    ' Log email details (sanitized)
    logText = logText & "=== Processing Email ===" & vbCrLf
    logText = logText & "Subject: " & SanitizeForLog(emailSubject) & vbCrLf
    logText = logText & "Received: " & Format(MItem.ReceivedTime, "yyyy-mm-dd hh:mm:ss") & vbCrLf
    
    ' Loop through all recipients and log their email addresses (sanitized)
    For Each Recipient In MItem.Recipients
        recipientEmailAddress = GetSanitizedRecipientEmail(Recipient)
        If recipientEmailAddress <> "" Then
            logText = logText & "Recipient: " & SanitizeForLog(recipientEmailAddress) & vbCrLf
        End If
    Next Recipient
    
    ' Check if email subject matches any configuration
    matchFound = False
    For Each configItem In configData
        If DoesSubjectMatchSecure(emailSubject, configItem("Subject"), configItem("MatchType")) Then
            matchFound = True
            destinationFolder = ValidateAndSanitizePath(configItem("DestinationFolder"))
            backupFolder = ValidateAndSanitizePath(configItem("BackupFolder"))
            destinationFileName = SanitizeFileName(configItem("DestinationFileName"))
            
            ' Validate paths before proceeding
            If destinationFolder = "" Or backupFolder = "" Then
                logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - ERROR: Invalid folder paths in configuration" & vbCrLf
                GoTo WriteLog
            End If
            
            logText = logText & "Match found - Type: " & configItem("MatchType") & vbCrLf
            Exit For
        End If
    Next configItem
    
    If Not matchFound Then
        logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - No matching subject configuration found." & vbCrLf
        GoTo WriteLog
    End If
    
    ' Process all attachments with security validation
    logText = logText & "Processing " & MItem.Attachments.Count & " attachments..." & vbCrLf
    
    For Each oAttachment In MItem.Attachments
        Dim sanitizedAttachmentName As String
        sanitizedAttachmentName = SanitizeFileName(oAttachment.DisplayName)
        
        ' Validate attachment
        If Not IsValidAttachment(oAttachment) Then
            logText = logText & "WARNING: Skipping invalid attachment: " & SanitizeForLog(sanitizedAttachmentName) & vbCrLf
            GoTo NextAttachment
        End If
        
        logText = logText & "Processing attachment: " & SanitizeForLog(sanitizedAttachmentName) & vbCrLf
        
        ' Replace placeholders in filename with validation
        Dim finalFileName As String
        finalFileName = ReplacePlaceholdersSecure(destinationFileName, sanitizedAttachmentName)
        
        ' Final validation of constructed filename
        If Not IsValidFileName(finalFileName) Then
            logText = logText & "ERROR: Invalid final filename generated" & vbCrLf
            GoTo NextAttachment
        End If
        
        logText = logText & "Final filename: " & SanitizeForLog(finalFileName) & vbCrLf
        
        ' Create folder paths if they don't exist (with validation)
        If Not CreateSecureFolderPath(destinationFolder) Or Not CreateSecureFolderPath(backupFolder) Then
            logText = logText & "ERROR: Unable to create destination folders" & vbCrLf
            GoTo NextAttachment
        End If
        
        ' Ensure destination folders end with backslash for proper path construction
        If Right(destinationFolder, 1) <> "\" Then
            destinationFolder = destinationFolder & "\"
        End If
        If Right(backupFolder, 1) <> "\" Then
            backupFolder = backupFolder & "\"
        End If
        
        ' Set full file paths with final validation
        primaryFilePath = destinationFolder & finalFileName
        backupFilePath = backupFolder & finalFileName
        
        ' Validate final paths
        If Not IsValidFilePath(primaryFilePath) Or Not IsValidFilePath(backupFilePath) Then
            logText = logText & "ERROR: Invalid file paths constructed" & vbCrLf
            GoTo NextAttachment
        End If
        
        ' Save to primary destination folder
        If SaveAttachmentSecurely(oAttachment, primaryFilePath) Then
            logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - Saved to primary location" & vbCrLf
            attachmentSaved = True
        Else
            logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - ERROR saving to primary location" & vbCrLf
        End If
        
        ' Save to backup folder
        If SaveAttachmentSecurely(oAttachment, backupFilePath) Then
            logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - Saved to backup location" & vbCrLf
            attachmentSaved = True
        Else
            logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - ERROR saving to backup location" & vbCrLf
        End If
        
NextAttachment:
    Next oAttachment
    
    If Not attachmentSaved Then
        logText = logText & Format(Now, "yyyy-mm-dd hh:mm:ss") & " - No attachments were saved." & vbCrLf
    End If
    
WriteLog:
    ' Write log to file securely
    WriteSecureLog logFilePath, logText
End Sub

' Security Functions

Private Function GetSecureMacroPath() As String
    Dim basePath As String
    Dim macroPath As String
    
    ' Try to get path from current workbook first
    On Error Resume Next
    basePath = ThisWorkbook.Path
    On Error GoTo 0
    
    If basePath <> "" Then
        macroPath = basePath & "\Macro\"
    Else
        ' Fallback to user profile with generic path
        basePath = Environ("USERPROFILE")
        If basePath <> "" Then
            macroPath = basePath & "\Documents\BIMailer\Macro\"
        End If
    End If
    
    ' Validate the path exists and is accessible
    If Dir(macroPath, vbDirectory) <> "" Then
        GetSecureMacroPath = macroPath
    Else
        GetSecureMacroPath = ""
    End If
End Function

Private Function SanitizeInput(inputStr As String) As String
    Dim result As String
    result = inputStr
    
    ' Remove potentially dangerous characters
    result = Replace(result, Chr(0), "")  ' Null bytes
    result = Replace(result, vbTab, " ")  ' Tabs
    result = Replace(result, vbCr, " ")   ' Carriage returns
    result = Replace(result, vbLf, " ")   ' Line feeds
    
    ' Limit length to prevent buffer overflow
    If Len(result) > 255 Then
        result = Left(result, 255)
    End If
    
    SanitizeInput = Trim(result)
End Function

Private Function SanitizeFileName(fileName As String) As String
    Dim result As String
    Dim invalidChars As String
    Dim i As Integer
    
    result = fileName
    invalidChars = "\/:*?""<>|"
    
    ' Remove invalid filename characters
    For i = 1 To Len(invalidChars)
        result = Replace(result, Mid(invalidChars, i, 1), "_")
    Next i
    
    ' Remove control characters
    For i = 0 To 31
        result = Replace(result, Chr(i), "")
    Next i
    
    ' Limit length
    If Len(result) > 200 Then
        result = Left(result, 200)
    End If
    
    ' Ensure not empty
    If Trim(result) = "" Then
        result = "attachment_" & Format(Now, "yyyymmdd_hhmmss")
    End If
    
    SanitizeFileName = Trim(result)
End Function

Private Function ValidateAndSanitizePath(folderPath As String) As String
    Dim result As String
    result = folderPath
    
    ' Basic validation
    If Len(result) = 0 Then
        ValidateAndSanitizePath = ""
        Exit Function
    End If
    
    ' Check for path traversal attempts
    If InStr(result, "..") > 0 Then
        LogSecurityEvent "WARNING: Path traversal attempt detected: " & result
        ValidateAndSanitizePath = ""
        Exit Function
    End If
    
    ' Remove dangerous characters
    result = Replace(result, Chr(0), "")
    result = Replace(result, "*", "")
    result = Replace(result, "?", "")
    result = Replace(result, """", "")
    result = Replace(result, "<", "")
    result = Replace(result, ">", "")
    result = Replace(result, "|", "")
    
    ' Validate length
    If Len(result) > 200 Then
        ValidateAndSanitizePath = ""
        Exit Function
    End If
    
    ValidateAndSanitizePath = result
End Function

Private Function IsValidEmailSubject(subject As String) As Boolean
    ' Basic validation for email subject
    If Len(subject) = 0 Or Len(subject) > 500 Then
        IsValidEmailSubject = False
    Else
        IsValidEmailSubject = True
    End If
End Function

Private Function IsValidAttachment(attachment As Outlook.Attachment) As Boolean
    ' Validate attachment properties
    If attachment Is Nothing Then
        IsValidAttachment = False
        Exit Function
    End If
    
    ' Check attachment size (limit to 50MB)
    If attachment.Size > 52428800 Then
        IsValidAttachment = False
        Exit Function
    End If
    
    ' Check for valid display name
    If Len(attachment.DisplayName) = 0 Or Len(attachment.DisplayName) > 255 Then
        IsValidAttachment = False
        Exit Function
    End If
    
    IsValidAttachment = True
End Function

Private Function IsValidFileName(fileName As String) As Boolean
    Dim invalidChars As String
    Dim i As Integer
    
    ' Check basic requirements
    If Len(fileName) = 0 Or Len(fileName) > 200 Then
        IsValidFileName = False
        Exit Function
    End If
    
    ' Check for invalid characters
    invalidChars = "\/:*?""<>|"
    For i = 1 To Len(invalidChars)
        If InStr(fileName, Mid(invalidChars, i, 1)) > 0 Then
            IsValidFileName = False
            Exit Function
        End If
    Next i
    
    ' Check for control characters
    For i = 0 To 31
        If InStr(fileName, Chr(i)) > 0 Then
            IsValidFileName = False
            Exit Function
        End If
    Next i
    
    IsValidFileName = True
End Function

Private Function IsValidFilePath(filePath As String) As Boolean
    ' Validate complete file path
    If Len(filePath) = 0 Or Len(filePath) > 260 Then
        IsValidFilePath = False
        Exit Function
    End If
    
    ' Check for path traversal
    If InStr(filePath, "..") > 0 Then
        IsValidFilePath = False
        Exit Function
    End If
    
    IsValidFilePath = True
End Function

Private Function SaveAttachmentSecurely(attachment As Outlook.Attachment, filePath As String) As Boolean
    On Error GoTo ErrorHandler
    
    ' Final validation before saving
    If Not IsValidFilePath(filePath) Then
        SaveAttachmentSecurely = False
        Exit Function
    End If
    
    ' Attempt to save
    attachment.SaveAsFile filePath
    SaveAttachmentSecurely = True
    Exit Function
    
ErrorHandler:
    LogSecurityEvent "ERROR: Failed to save attachment - " & Err.Description
    SaveAttachmentSecurely = False
End Function

Private Function CreateSecureFolder(folderPath As String) As Boolean
    On Error GoTo ErrorHandler
    
    If Dir(folderPath, vbDirectory) = "" And folderPath <> "" Then
        MkDir folderPath
    End If
    
    CreateSecureFolder = True
    Exit Function
    
ErrorHandler:
    CreateSecureFolder = False
End Function

Private Function CreateSecureFolderPath(folderPath As String) As Boolean
    Dim pathParts() As String
    Dim currentPath As String
    Dim i As Integer
    
    ' Validate input
    If Not IsValidFilePath(folderPath) Then
        CreateSecureFolderPath = False
        Exit Function
    End If
    
    ' Remove trailing backslash if present
    If Right(folderPath, 1) = "\" Then
        folderPath = Left(folderPath, Len(folderPath) - 1)
    End If
    
    ' Split path into parts
    pathParts = Split(folderPath, "\")
    
    ' Build path incrementally and create folders as needed
    For i = 0 To UBound(pathParts)
        If i = 0 Then
            currentPath = pathParts(i)
        Else
            currentPath = currentPath & "\" & pathParts(i)
        End If
        
        ' Validate each path component
        If Not IsValidFilePath(currentPath) Then
            CreateSecureFolderPath = False
            Exit Function
        End If
        
        ' Create folder if it doesn't exist
        If Not CreateSecureFolder(currentPath) Then
            CreateSecureFolderPath = False
            Exit Function
        End If
    Next i
    
    CreateSecureFolderPath = True
End Function

Private Function GetSanitizedRecipientEmail(Recipient As Outlook.Recipient) As String
    Dim emailAddress As String
    
    On Error GoTo ErrorHandler
    
    If Recipient.AddressEntry Is Nothing Then
        emailAddress = Recipient.Address
    Else
        If Recipient.AddressEntry.AddressEntryUserType = OlAddressEntryUserType.olExchangeUserAddressEntry Then
            Set exUser = Recipient.AddressEntry.GetExchangeUser()
            emailAddress = exUser.PrimarySmtpAddress
        Else
            emailAddress = Recipient.Address
        End If
    End If
    
    GetSanitizedRecipientEmail = SanitizeInput(emailAddress)
    Exit Function
    
ErrorHandler:
    GetSanitizedRecipientEmail = ""
End Function

Private Function SanitizeForLog(inputStr As String) As String
    Dim result As String
    result = inputStr
    
    ' Remove sensitive information patterns
    ' This is a basic implementation - expand as needed
    result = Replace(result, "@", "[at]")  ' Obfuscate email addresses
    
    ' Limit length for log files
    If Len(result) > 100 Then
        result = Left(result, 100) & "..."
    End If
    
    SanitizeForLog = result
End Function

Private Sub WriteSecureLog(logFilePath As String, logText As String)
    Dim logFile As Integer
    
    On Error GoTo ErrorHandler
    
    ' Validate log file path
    If Not IsValidFilePath(logFilePath) Then
        Exit Sub
    End If
    
    logFile = FreeFile
    Open logFilePath For Append As logFile
    Print #logFile, logText & vbCrLf
    Close logFile
    Exit Sub
    
ErrorHandler:
    ' Fail silently for logging errors to prevent infinite loops
    If logFile > 0 Then
        Close logFile
    End If
End Sub

Private Sub LogSecurityEvent(eventText As String)
    ' Log security events to a separate file
    Dim securityLogPath As String
    Dim macroPath As String
    
    macroPath = GetSecureMacroPath()
    If macroPath <> "" Then
        securityLogPath = macroPath & "Logs\SecurityLog.txt"
        WriteSecureLog securityLogPath, Format(Now, "yyyy-mm-dd hh:mm:ss") & " - " & eventText
    End If
End Sub

' Modified existing functions with security enhancements

Private Function LoadSubjectConfigSecure(configPath As String) As Collection
    Dim configData As Collection
    Dim configItem As Collection
    Dim fileNum As Integer
    Dim line As String
    Dim fields() As String
    Dim isFirstLine As Boolean
    Dim lineCount As Integer
    
    Set configData = New Collection
    isFirstLine = True
    lineCount = 0
    
    ' Validate config file path
    If Not IsValidFilePath(configPath) Then
        Set LoadSubjectConfigSecure = Nothing
        Exit Function
    End If
    
    ' Check if file exists
    If Dir(configPath) = "" Then
        Set LoadSubjectConfigSecure = Nothing
        Exit Function
    End If
    
    On Error GoTo ErrorHandler
    
    ' Read CSV file with limits
    fileNum = FreeFile
    Open configPath For Input As fileNum
    
    Do While Not EOF(fileNum) And lineCount < 1000  ' Limit number of lines
        Line Input #fileNum, line
        lineCount = lineCount + 1
        
        ' Sanitize input line
        line = SanitizeInput(line)
        
        ' Skip header line and comment lines
        If isFirstLine Or Left(Trim(line), 1) = "#" Or Trim(line) = "" Then
            isFirstLine = False
            GoTo NextLine
        End If
        
        ' Parse CSV line with validation
        fields = Split(line, ",")
        If UBound(fields) >= 4 Then
            Set configItem = New Collection
            configItem.Add SanitizeInput(fields(0)), "Subject"
            configItem.Add UCase(SanitizeInput(fields(1))), "MatchType"
            configItem.Add ValidateAndSanitizePath(fields(2)), "DestinationFolder"
            configItem.Add ValidateAndSanitizePath(fields(3)), "BackupFolder"
            configItem.Add SanitizeFileName(fields(4)), "DestinationFileName"
            
            ' Only add if all fields are valid
            If configItem("DestinationFolder") <> "" And configItem("BackupFolder") <> "" Then
                configData.Add configItem
            End If
        End If
        
NextLine:
    Loop
    
    Close fileNum
    Set LoadSubjectConfigSecure = configData
    Exit Function
    
ErrorHandler:
    If fileNum > 0 Then
        Close fileNum
    End If
    Set LoadSubjectConfigSecure = Nothing
End Function

Private Function DoesSubjectMatchSecure(emailSubject As String, configSubject As String, matchType As String) As Boolean
    Dim result As Boolean
    result = False
    
    ' Validate inputs
    If Len(emailSubject) = 0 Or Len(configSubject) = 0 Or Len(matchType) = 0 Then
        DoesSubjectMatchSecure = False
        Exit Function
    End If
    
    ' Convert to uppercase for case-insensitive comparison
    Dim upperEmailSubject As String
    Dim upperConfigSubject As String
    upperEmailSubject = UCase(emailSubject)
    upperConfigSubject = UCase(configSubject)
    
    Select Case matchType
        Case "EXACT"
            result = (upperEmailSubject = upperConfigSubject)
        Case "CONTAINS"
            result = (InStr(upperEmailSubject, upperConfigSubject) > 0)
        Case "STARTS_WITH"
            result = (Left(upperEmailSubject, Len(upperConfigSubject)) = upperConfigSubject)
        Case "ENDS_WITH"
            result = (Right(upperEmailSubject, Len(upperConfigSubject)) = upperConfigSubject)
        Case Else
            result = False
    End Select
    
    DoesSubjectMatchSecure = result
End Function

Private Function ReplacePlaceholdersSecure(fileName As String, originalName As String) As String
    Dim result As String
    result = fileName
    
    ' Validate inputs
    fileName = SanitizeFileName(fileName)
    originalName = SanitizeFileName(originalName)
    
    ' Replace timestamp placeholder
    result = Replace(result, "{timestamp}", Format(Now, "yyyymmdd_hhmmss"))
    
    ' Replace date placeholder
    result = Replace(result, "{date}", Format(Now, "yyyymmdd"))
    
    ' Replace time placeholder
    result = Replace(result, "{time}", Format(Now, "hhmmss"))
    
    ' Use configured filename, fallback to sanitized original if empty
    If Trim(fileName) = "" Then
        result = originalName
    End If
    
    ' Final sanitization
    result = SanitizeFileName(result)
    
    ReplacePlaceholdersSecure = result
End Function
