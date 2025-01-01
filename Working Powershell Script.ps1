# First, ensure MSAL.PS is installed
if (-not (Get-Module -ListAvailable -Name MSAL.PS)) {
    Install-Module -Name MSAL.PS -Force -Scope CurrentUser
}

Import-Module MSAL.PS

# Part 1: Configuration and Basic Functions
# Configuration details
$TenantId = "a5ae9ae1-3c47-4b70-b92c-ac3a0efffc6a"
$ClientId = "9ec41cd0-ae8c-4dd5-bc84-a3aeea4bda54"

# Function to securely save token
function Save-IntervalsToken {
    param (
        [string]$Token
    )
    try {
        $SecureToken = ConvertTo-SecureString $Token -AsPlainText -Force
        $SecureTokenText = $SecureToken | ConvertFrom-SecureString
        $TokenPath = Join-Path $env:USERPROFILE ".intervals_token"
        $SecureTokenText | Out-File $TokenPath
        Write-Host "Token saved successfully!" -ForegroundColor Green
    }
    catch {
        Write-Host "Error saving token: $($_.Exception.Message)" -ForegroundColor Red
    }
}

# Function to load saved token
function Get-SavedIntervalsToken {
    try {
        $TokenPath = Join-Path $env:USERPROFILE ".intervals_token"
        if (Test-Path $TokenPath) {
            $SecureTokenText = Get-Content $TokenPath
            $SecureToken = $SecureTokenText | ConvertTo-SecureString
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
            $Token = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
            return $Token
        }
    }
    catch {
        Write-Host "Error loading saved token: $($_.Exception.Message)" -ForegroundColor Yellow
    }
    return $null
}

# Get Intervals API Token
$SavedToken = Get-SavedIntervalsToken
if ($SavedToken) {
    $UseSaved = Read-Host "Found saved Intervals token. Use it? (Y/N)"
    if ($UseSaved.ToUpper() -eq 'Y') {
        $ApiToken = $SavedToken
        Write-Host "Using saved token" -ForegroundColor Green
    }
}

if (-not $ApiToken) {
    # Use PowerShell SecureString for masked input
    Write-Host "Please enter your Intervals API token: " -ForegroundColor Cyan -NoNewline
    $SecureToken = Read-Host -AsSecureString
    $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($SecureToken)
    $ApiToken = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)

    if ([string]::IsNullOrWhiteSpace($ApiToken)) {
        Write-Host "Error: Intervals API token is required." -ForegroundColor Red
        exit
    }

    # Ask to save token
    $SaveToken = Read-Host "Would you like to save this token for future use? (Y/N)"
    if ($SaveToken.ToUpper() -eq 'Y') {
        Save-IntervalsToken -Token $ApiToken
    }
}

# Add colon to API token if not present
if (-not $ApiToken.EndsWith(":")) {
    $ApiToken = "$($ApiToken):"
}

$EncodedToken = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes($ApiToken))
$HeadersIntervals = @{
    "Authorization" = "Basic $EncodedToken"
}

# Test Intervals API connection with retry logic
$MaxRetries = 3
$RetryCount = 0
$Connected = $false

while (-not $Connected -and $RetryCount -lt $MaxRetries) {
    try {
        $TestResponse = Invoke-RestMethod -Uri "https://api.myintervals.com/me" -Method Get -Headers $HeadersIntervals
        Write-Host "Successfully connected to Intervals as: $($TestResponse.intervals.me.item.firstname) $($TestResponse.intervals.me.item.lastname)" -ForegroundColor Green
        $Connected = $true
    }
    catch {
        $RetryCount++
        if ($RetryCount -ge $MaxRetries) {
            Write-Host "Error connecting to Intervals after $MaxRetries attempts. Please check your API token." -ForegroundColor Red
            Write-Host "Error details: $($_.Exception.Message)" -ForegroundColor Red
            exit
        }
        Write-Host "Connection attempt $RetryCount failed. Retrying..." -ForegroundColor Yellow
        Start-Sleep -Seconds 2
    }
}

# Azure OpenAI Configuration
$AzureOpenAIKey = "CWDspACTbjoETrgOOAi7i2cGXJiHRrFEg6ZciiqxXdy3u9aIWcuSJQQJ99ALACYeBjFXJ3w3AAABACOGQTIv"
$AzureOpenAIEndpoint = "https://rajesh-azure-open-ai.openai.azure.com/openai/deployments/gpt-4o/chat/completions?api-version=2024-08-01-preview"

# Add System.Web for URL encoding/decoding
Add-Type -AssemblyName System.Web

# MSAL Authentication function
function Get-ApplicationAccessToken {
    param (
        [string]$TenantId,
        [string]$ClientId
    )

    Write-Host "Getting Microsoft Graph access token using MSAL..." -ForegroundColor Cyan
    
    try {
        $Scopes = @(
            "https://graph.microsoft.com/Calendars.Read",
            "https://graph.microsoft.com/User.Read",
            "https://graph.microsoft.com/OnlineMeetings.Read"
        )

        $Token = Get-MsalToken -ClientId $ClientId `
            -TenantId $TenantId `
            -RedirectUri "https://login.microsoftonline.com/common/oauth2/nativeclient" `
            -Interactive `
            -Scopes $Scopes

        Write-Host "Successfully authenticated as: $($Token.Account.Username)" -ForegroundColor Green
        
        # Store the user ID globally for use in the script
        $script:TargetUserId = $Token.Account.Username
        
        return $Token.AccessToken
    }
    catch {
        Write-Host "Error getting access token: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Basic utility functions
function Format-Duration {
    param ([int]$TotalSeconds)
    
    $Hours = [math]::Floor($TotalSeconds / 3600)
    $Minutes = [math]::Floor(($TotalSeconds % 3600) / 60)
    $Seconds = $TotalSeconds % 60
    return "$Hours hours, $Minutes minutes, $Seconds seconds"
}

function Clean-Text {
    param (
        [string]$Text,
        [switch]$RemoveEmoji
    )
    
    $CleanText = $Text
    
    if ($RemoveEmoji) {
        # Remove emojis and special characters
        $CleanText = $CleanText -replace '[^\x00-\x7F]+', ''
    }
    
    # Clean up whitespace
    $CleanText = $CleanText -replace '\s+', ' '
    $CleanText = $CleanText -replace '[\r\n]+', ' '
    $CleanText = $CleanText.Trim()
    
    if ([string]::IsNullOrWhiteSpace($CleanText)) {
        return "Untitled"
    }
    
    return $CleanText
}

function GetMeetingDecimalTime {
    param ([int]$AttendanceSeconds)
    
    # Convert seconds directly to hours with precision
    $ExactHours = $AttendanceSeconds / 3600
    
    # Round to 1 decimal place
    $ExactHours = [math]::Round($ExactHours, 1)
    
    Write-Host "Exact attendance time: $ExactHours hours" -ForegroundColor Gray
    return $ExactHours
}

function Write-ProcessMessage {
    param (
        [string]$Message,
        [string]$Color = "White",
        [switch]$NoNewline
    )
    
    $Params = @{
        Object = $Message
        ForegroundColor = $Color
    }
    
    if ($NoNewline) {
        $Params.Add("NoNewline", $true)
    }
    
    Write-Host @Params
}

# Part 2: Meeting Fetching and Attendance Reports

function Get-UserMeetings {
    param (
        [string]$UserId,
        [string]$AccessToken
    )
    
    Write-Host "Retrieving meetings for user: $UserId..." -ForegroundColor Cyan
    
    $Headers = @{ 
        Authorization = "Bearer $AccessToken"
    }
    
    # Get user information first
    $TargetUser = Invoke-RestMethod -Method Get -Uri "https://graph.microsoft.com/v1.0/users/$UserId" -Headers $Headers
    if (-not $TargetUser) {
        Write-Host "Failed to retrieve target user information. Exiting..." -ForegroundColor Red
        return @()
    }
    Write-Host "Target User ID: $($TargetUser.id)" -ForegroundColor Cyan
    
    # Calculate date range
    $StartDate = (Get-Date).AddDays(-30).ToString("yyyy-MM-ddTHH:mm:ssZ")
    $EndDate = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ssZ")
    
    Write-Host "Retrieving meetings from $StartDate to $EndDate..." -ForegroundColor Cyan
    
    $Uri = "https://graph.microsoft.com/v1.0/users/$($TargetUser.id)/events?"
    $Filter = [System.Web.HttpUtility]::UrlEncode("start/dateTime ge '$StartDate' and end/dateTime le '$EndDate'")
    $MeetingsUri = "$Uri`$filter=$Filter&`$select=subject,start,end,onlineMeeting,onlineMeetingUrl,bodyPreview"
    
    try {
        $Meetings = Invoke-RestMethod -Method Get -Uri $MeetingsUri -Headers $Headers
        
        if ($Meetings.value) {
            $ProcessedMeetings = @()
            
            foreach ($Meeting in $Meetings.value) {
                Write-Host "`nProcessing meeting: $($Meeting.subject)" -ForegroundColor Green
                
                if ($Meeting.bodyPreview) {
                    Write-Host "Meeting Description:`n$($Meeting.bodyPreview)" -ForegroundColor Cyan
                }

                if ($Meeting.onlineMeeting) {
                    # Decode Join URL
                    $DecodedJoinUrl = [System.Web.HttpUtility]::UrlDecode($Meeting.onlineMeeting.joinUrl)
                    if ($DecodedJoinUrl -match "19:meeting_([^@]+)@thread.v2") {
                        $MeetingId = "19:meeting_" + $matches[1] + "@thread.v2"
                    }

                    if ($DecodedJoinUrl -match '"Oid":"([^"]+)"') {
                        $OrganizerOid = $matches[1]
                    }

                    $FormattedString = "1*${OrganizerOid}*0**${MeetingId}"
                    $Base64MeetingId = [System.Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($FormattedString))

                    Try {
                        Write-Host "Retrieving attendance report for Base64 Meeting ID: $Base64MeetingId..."
                        $AttendanceReport = Invoke-RestMethod -Method Get `
                            -Uri "https://graph.microsoft.com/v1.0/users/$($TargetUser.id)/onlineMeetings/$Base64MeetingId/attendanceReports" `
                            -Headers $Headers

                        if ($AttendanceReport.value) {
                            $ReportId = $AttendanceReport.value[0].id
                            Write-Host "Attendance Report ID: $ReportId" -ForegroundColor Green

                            # Retrieve Attendance Records
                            $AttendanceRecords = Invoke-RestMethod -Method Get `
                                -Uri "https://graph.microsoft.com/v1.0/users/$($TargetUser.id)/onlineMeetings/$Base64MeetingId/attendanceReports/$ReportId/attendanceRecords" `
                                -Headers $Headers

                            if ($AttendanceRecords.value) {
                                $AttendanceDescription = "Attendance Report`n"
                                foreach ($Record in $AttendanceRecords.value) {
                                    $UserName = $Record.identity.displayName ?? "Unknown User"
                                    $EmailAddress = $Record.emailAddress ?? "No Email Address"
                                    $AttendanceDescription += "User: $UserName, Email: $EmailAddress, Total Time: $($Record.totalAttendanceInSeconds) seconds`n"
                                }
                                $Meeting | Add-Member -NotePropertyName "description" -NotePropertyValue $AttendanceDescription -Force
                            }
                        }
                    }
                    Catch {
                        # Use scheduled duration as fallback
                        $StartTime = [DateTime]$Meeting.start.dateTime
                        $EndTime = [DateTime]$Meeting.end.dateTime
                        $DurationSeconds = ([int]($EndTime - $StartTime).TotalSeconds)
                        
                        $AttendanceDescription = "Attendance Report`n"
                        $AttendanceDescription += "Note: Using scheduled duration.`n"
                        $AttendanceDescription += "User: $($TargetUser.displayName), Email: $UserId, Total Time: $DurationSeconds seconds`n"
                        $Meeting | Add-Member -NotePropertyName "description" -NotePropertyValue $AttendanceDescription -Force
                    }
                } else {
                    # Handle non-Teams meetings
                    $StartTime = [DateTime]$Meeting.start.dateTime
                    $EndTime = [DateTime]$Meeting.end.dateTime
                    $DurationSeconds = ([int]($EndTime - $StartTime).TotalSeconds)
                    
                    $AttendanceDescription = "Attendance Report`n"
                    $AttendanceDescription += "Note: Not a Teams meeting - using scheduled duration.`n"
                    $AttendanceDescription += "User: $($TargetUser.displayName), Email: $UserId, Total Time: $DurationSeconds seconds`n"
                    $Meeting | Add-Member -NotePropertyName "description" -NotePropertyValue $AttendanceDescription -Force
                }
                
                $ProcessedMeetings += $Meeting
            }
            
            return $ProcessedMeetings
        } else {
            Write-Host "No meetings found in the specified date range." -ForegroundColor Yellow
            return @()
        }
    }
    catch {
        Write-Host "Error fetching meetings: $($_.Exception.Message)" -ForegroundColor Red
        return @()
    }
}

function Show-AttendanceReport {
    param (
        [Parameter(Mandatory = $true)]
        [object]$Meeting
    )
    
    Write-Host "`n=============================================" -ForegroundColor Cyan
    Write-Host "Meeting Attendance Report" -ForegroundColor Cyan
    Write-Host "=============================================" -ForegroundColor Cyan
    
    # Display meeting details
    Write-Host "`nMeeting Subject: " -NoNewline
    Write-Host $Meeting.subject -ForegroundColor Yellow
    
    Write-Host "Start Time: " -NoNewline
    Write-Host ([DateTime]$Meeting.start.dateTime).ToString("yyyy-MM-dd HH:mm:ss") -ForegroundColor Green
    
    Write-Host "End Time: " -NoNewline
    Write-Host ([DateTime]$Meeting.end.dateTime).ToString("yyyy-MM-dd HH:mm:ss") -ForegroundColor Green
    
    # Calculate scheduled duration
    $ScheduledDuration = ([DateTime]$Meeting.end.dateTime - [DateTime]$Meeting.start.dateTime).TotalMinutes
    Write-Host "Scheduled Duration: " -NoNewline
    Write-Host "$ScheduledDuration minutes" -ForegroundColor Green
    
    Write-Host "`n----- Attendance Details -----" -ForegroundColor Cyan
    
    $AttendanceData = @()
    $MaxAttendanceSeconds = 0
    
    if ($Meeting.description -match "Attendance Report") {
        $Lines = $Meeting.description -split "`n"
        $ProcessedUsers = @{}  # Add this to track processed users
        
        foreach ($Line in $Lines) {
            if ($Line -match "User: (.*), Email: (.*), Total Time: (\d+) seconds") {
                $UserName = $Matches[1].Trim()
                $Email = $Matches[2].Trim()
                $AttendanceSeconds = [int]$Matches[3]
                
                # Skip if we've already processed this user
                if ($ProcessedUsers.ContainsKey($Email)) {
                    continue
                }
                
                $ProcessedUsers[$Email] = $true
                $AttendanceMinutes = [math]::Round($AttendanceSeconds / 60, 1)
                
                # Update max attendance
                if ($AttendanceSeconds -gt $MaxAttendanceSeconds) {
                    $MaxAttendanceSeconds = $AttendanceSeconds
                }
                
                $AttendanceData += [PSCustomObject]@{
                    Name = $UserName
                    Email = $Email
                    Minutes = $AttendanceMinutes
                    Seconds = $AttendanceSeconds
                    Percentage = [math]::Round(($AttendanceMinutes / $ScheduledDuration) * 100, 1)
                }
            }
        }
        
        if ($AttendanceData.Count -gt 0) {
            foreach ($Attendee in ($AttendanceData | Sort-Object Minutes -Descending)) {
                Write-Host "`nAttendee: " -NoNewline
                Write-Host $Attendee.Name -ForegroundColor Yellow
                Write-Host "Email: " -NoNewline
                Write-Host $Attendee.Email -ForegroundColor Gray
                Write-Host "Duration: " -NoNewline
                Write-Host "$($Attendee.Minutes) minutes" -ForegroundColor Green
                Write-Host "Attendance: " -NoNewline
                Write-Host "$($Attendee.Percentage)%" -ForegroundColor $(
                    if ($Attendee.Percentage -ge 90) { "Green" }
                    elseif ($Attendee.Percentage -ge 75) { "Yellow" }
                    else { "Red" }
                )
            }
        }
        else {
            Write-Host "No attendance data found in the meeting description." -ForegroundColor Yellow
        }
    }
    else {
        Write-Host "No attendance report available for this meeting." -ForegroundColor Yellow
    }
    
    Write-Host "`n=============================================" -ForegroundColor Cyan
    
    # Return both attendance data and max duration
    return @{
        AttendanceData = $AttendanceData
        MaxAttendanceSeconds = $MaxAttendanceSeconds
    }
}

# Part 3: Task Matching and API Interactions

# Intervals API Functions
function Get-IntervalsCurrentUser {
    Write-Host "Fetching current user from Intervals..." -ForegroundColor Cyan
    
    try {
        $Response = Invoke-RestMethod -Uri "https://api.myintervals.com/me" -Method Get -Headers $HeadersIntervals
        
        if ($Response.intervals.me.item) {
            $User = $Response.intervals.me.item
            Write-Host "Successfully retrieved user: $($User.firstname) $($User.lastname)" -ForegroundColor Green
            
            if ($null -eq $User.personid -and $User.id) {
                $User | Add-Member -NotePropertyName "personid" -NotePropertyValue $User.id -Force
            }
            
            return $User
        }
        throw "No user information found"
    }
    catch {
        Write-Host "Error fetching user: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Get-IntervalsData {
    param (
        [string]$Endpoint
    )
    
    $Url = "https://api.myintervals.com$Endpoint"
    
    try {
        $Response = Invoke-RestMethod -Uri $Url -Method Get -Headers $HeadersIntervals
        return $Response
    }
    catch {
        Write-Host "Error fetching Intervals data: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Post-TimeEntry {
    param (
        [hashtable]$TimeEntry
    )
    
    Write-Host "Posting time entry..." -ForegroundColor Cyan
    
    try {
        # Clean description
        $TimeEntry.description = Clean-Text -Text $TimeEntry.description -RemoveEmoji
        
        $Response = Invoke-RestMethod `
            -Uri "https://api.myintervals.com/time" `
            -Method Post `
            -Headers $HeadersIntervals `
            -Body ($TimeEntry | ConvertTo-Json) `
            -ContentType 'application/json'
        
        Write-Host "Successfully posted time entry ($($TimeEntry.time) hours)" -ForegroundColor Green
        return $Response
    }
    catch {
        Write-Host "Error posting time entry: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
}

function Get-TaskContext {
    param (
        [array]$Tasks
    )
    
    Write-Host "Building task context..." -ForegroundColor Cyan
    $TaskContext = @{}
    
    foreach ($Task in $Tasks) {
        Write-Host "  Processing: $($Task.title)" -ForegroundColor Gray
        
        $TaskContext[$Task.id] = @{
            title_words = @()
            project_name = ""
            module_name = ""
            keywords = @()
            usage_count = 0
        }
        
        # Process title
        $CleanTitle = Clean-Text -Text $Task.title -RemoveEmoji
        $TitleWords = $CleanTitle.ToLower() -split '\W+' | 
            Where-Object { $_.Length -gt 2 } | 
            Where-Object { $_ -notmatch '^\d+$' }
        
        $TaskContext[$Task.id].title_words = $TitleWords
        $TaskContext[$Task.id].keywords += $TitleWords
        
        # Add project context if available
        try {
            if ($Task.projectid) {
                $ProjectInfo = Get-IntervalsData -Endpoint "/project/$($Task.projectid)"
                if ($ProjectInfo.project.name) {
                    $TaskContext[$Task.id].project_name = $ProjectInfo.project.name
                    
                    $ProjectWords = Clean-Text -Text $ProjectInfo.project.name -RemoveEmoji
                    $ProjectWords = $ProjectWords.ToLower() -split '\W+' | 
                        Where-Object { $_.Length -gt 2 } |
                        Where-Object { $_ -notmatch '^\d+$' }
                    $TaskContext[$Task.id].keywords += $ProjectWords
                }
            }
        }
        catch {
            Write-Host "    Unable to fetch project info for task $($Task.id)" -ForegroundColor Yellow
        }
        
        # Remove duplicates from keywords
        $TaskContext[$Task.id].keywords = $TaskContext[$Task.id].keywords | Select-Object -Unique
    }
    
    return $TaskContext
}

function Get-DirectMatch {
    param (
        [string]$MeetingTitle,
        [array]$Tasks,
        [hashtable]$TaskContext
    )
    
    $CleanTitle = Clean-Text -Text $MeetingTitle -RemoveEmoji
    $TitleWords = $CleanTitle.ToLower() -split '\W+' | 
        Where-Object { $_.Length -gt 2 } |
        Where-Object { $_ -notmatch '^\d+$' }
    
    $BestMatch = @{
        TaskId = $null
        Score = 0
    }
    
    foreach ($Task in $Tasks) {
        $MatchCount = 0
        $TaskKeywords = $TaskContext[$Task.id].keywords
        
        foreach ($Word in $TitleWords) {
            if ($TaskKeywords -contains $Word) {
                $MatchCount++
            }
        }
        
        if ($TitleWords.Count -gt 0) {
            $Score = $MatchCount / $TitleWords.Count
            
            if ($Score -gt $BestMatch.Score) {
                $BestMatch.TaskId = $Task.id
                $BestMatch.Score = $Score
            }
        }
    }
    
    if ($BestMatch.Score -gt 0.5) {
        Write-Host "Direct match found (Score: $([math]::Round($BestMatch.Score, 2)))" -ForegroundColor Green
        return $BestMatch.TaskId
    }
    
    return $null
}

function Get-AzureOpenAIMatch {
    param (
        [string]$MeetingSubject,
        [array]$Tasks,
        [hashtable]$TaskContext
    )
    
    $CleanSubject = Clean-Text -Text $MeetingSubject -RemoveEmoji
    
    # Build task analysis string
    $TaskAnalysis = $Tasks | ForEach-Object {
        $TaskId = $_.id
        $Context = $TaskContext[$TaskId]
        
        @"
Task ID: $TaskId
Title: $($_.title)
Project: $($Context.project_name)
Keywords: $($Context.keywords -join ', ')
"@
    } | Out-String

    $Prompt = @"
Match this meeting title with the most relevant task.

Meeting Title: $CleanSubject

Available Tasks:
$TaskAnalysis

Instructions:
1. Match based on meeting title and task titles/keywords
2. Look for direct keyword matches first
3. Consider project context
4. For infrastructure/network meetings, prefer infrastructure tasks
5. For client-specific meetings, match to respective tasks

Response format:
{
  "taskId": "numeric_id_or_NO_MATCH",
  "confidence": "high|medium|low"
}
"@

    try {
        Write-Host "Analyzing meeting: $CleanSubject" -ForegroundColor Cyan
        
        $Body = @{
            messages = @(
                @{
                    role = "system"
                    content = "You are a task matcher focusing on title keywords and project context."
                }
                @{
                    role = "user"
                    content = $Prompt
                }
            )
            temperature = 0.3
            max_tokens = 100
        } | ConvertTo-Json -Depth 10

        $Response = Invoke-RestMethod -Uri $AzureOpenAIEndpoint -Method Post -Headers @{
            "api-key" = $AzureOpenAIKey
            "Content-Type" = "application/json"
        } -Body $Body
        
        $MatchResult = $Response.choices[0].message.content | ConvertFrom-Json
        
        Write-Host "Match found:" -ForegroundColor Green
        Write-Host "  Task ID: $($MatchResult.taskId)" -ForegroundColor Yellow
        Write-Host "  Confidence: $($MatchResult.confidence)" -ForegroundColor $(
            switch ($MatchResult.confidence) {
                "high"   { "Green" }
                "medium" { "Yellow" }
                "low"    { "Red" }
                default  { "White" }
            }
        )
        
        if ($MatchResult.confidence -eq "low") {
            return "NO_MATCH"
        }
        
        return $MatchResult.taskId
    }
    catch {
        Write-Host "Error in AI matching: $($_.Exception.Message)" -ForegroundColor Red
        return "NO_MATCH"
    }
}

# Part 4: Main Processing Logic and Execution

function Process-Meeting {
    param (
        [object]$Meeting,
        [array]$Tasks,
        [hashtable]$TaskContext,
        [PSObject]$CurrentUser,
        [int]$CurrentCount,
        [int]$TotalCount
    )
    
    Write-Host "`nProcessing Meeting ($CurrentCount of $TotalCount): $($Meeting.subject)" -ForegroundColor Green
    
    # Show detailed attendance report first
    $AttendanceResult = Show-AttendanceReport -Meeting $Meeting
    
    # Use the max attendance seconds for billing
    $AttendanceSeconds = $AttendanceResult.MaxAttendanceSeconds
    
    # If no attendance data, use scheduled duration
    if ($AttendanceSeconds -eq 0) {
        $StartTime = [DateTime]$Meeting.start.dateTime
        $EndTime = [DateTime]$Meeting.end.dateTime
        $AttendanceSeconds = [int]($EndTime - $StartTime).TotalSeconds
    }
    
    Write-Host "Actual duration: $(Format-Duration -TotalSeconds $AttendanceSeconds)" -ForegroundColor Gray
    
    # Calculate billable time based on actual attendance
    $BillableHours = GetMeetingDecimalTime -AttendanceSeconds $AttendanceSeconds
    Write-Host "Billable hours calculated: $BillableHours" -ForegroundColor Gray
    
    # Match meeting to task
    $MatchedTaskId = Get-DirectMatch -MeetingTitle $Meeting.subject -Tasks $Tasks -TaskContext $TaskContext
    if (-not $MatchedTaskId) {
        $MatchedTaskId = Get-AzureOpenAIMatch -MeetingSubject $Meeting.subject -Tasks $Tasks -TaskContext $TaskContext
    }
    
    if ($MatchedTaskId -eq "NO_MATCH") {
        Write-Host "No task match found for meeting" -ForegroundColor Yellow
        return [PSCustomObject]@{
            Meeting = $Meeting.subject
            Time = ([DateTime]$Meeting.start.dateTime).ToString("yyyy-MM-dd HH:mm")
            TaskId = "N/A"
            TaskTitle = "No Match"
            MatchStatus = "Not Matched"
            Posted = "No"
            BillableDuration = $BillableHours
            Duration = $AttendanceSeconds
            ActualMinutes = [math]::Round($AttendanceSeconds / 60)
        }
    }
    
    $MatchedTask = $Tasks | Where-Object { $_.id -eq $MatchedTaskId }
    Write-Host "Matched to task: $($MatchedTask.title)" -ForegroundColor Green
    
    # Post time entry with actual attendance-based duration
    $TimeEntry = @{
        personid = $CurrentUser.personid
        projectid = $MatchedTask.projectid
        moduleid = $MatchedTask.moduleid
        taskid = $MatchedTask.id
        worktypeid = 799573
        date = ([DateTime]$Meeting.start.dateTime).ToString("yyyy-MM-dd")
        time = $BillableHours
        description = "$($Meeting.subject) (Actual attendance: $([math]::Round($AttendanceSeconds / 60)) minutes)"
        billable = $true
    }
    
    $PostResult = Post-TimeEntry -TimeEntry $TimeEntry
    $PostStatus = if ($PostResult) { "Yes" } else { "Failed" }
    
    return [PSCustomObject]@{
        Meeting = $Meeting.subject
        Time = ([DateTime]$Meeting.start.dateTime).ToString("yyyy-MM-dd HH:mm")
        TaskId = $MatchedTask.id
        TaskTitle = $MatchedTask.title
        MatchStatus = "Matched"
        Posted = $PostStatus
        BillableDuration = $BillableHours
        Duration = $AttendanceSeconds
        ActualMinutes = [math]::Round($AttendanceSeconds / 60)
    }
}

function Show-ProcessingStatistics {
    param (
        [Array]$Results
    )
    
    Write-Host "`n=== Meeting Processing Report ===" -ForegroundColor Cyan
    
    # Calculate statistics
    $TotalMeetings = $Results.Count
    $MatchedMeetings = ($Results | Where-Object { $_.MatchStatus -eq "Matched" }).Count
    $PostedEntries = ($Results | Where-Object { $_.Posted -eq "Yes" }).Count
    $TotalBillableHours = ($Results | Where-Object { $_.Posted -eq "Yes" } | 
        Measure-Object -Property BillableDuration -Sum).Sum
    
    # Display summary
    Write-Host "`nProcessing Summary:" -ForegroundColor White
    Write-Host "Total Meetings: $TotalMeetings" -ForegroundColor White
    Write-Host "Successfully Matched: $MatchedMeetings" -ForegroundColor Green
    Write-Host "Time Entries Posted: $PostedEntries" -ForegroundColor Cyan
    Write-Host "Unmatched Meetings: $($TotalMeetings - $MatchedMeetings)" -ForegroundColor Yellow
    
    # Calculate success rates
    if ($TotalMeetings -gt 0) {
        $MatchRate = [math]::Round(($MatchedMeetings / $TotalMeetings) * 100, 1)
        Write-Host "Match Success Rate: ${MatchRate}%" -ForegroundColor Cyan
    }
    
    if ($MatchedMeetings -gt 0) {
        $PostRate = [math]::Round(($PostedEntries / $MatchedMeetings) * 100, 1)
        Write-Host "Post Success Rate: ${PostRate}%" -ForegroundColor Cyan
    }
    
    # Time summary
    Write-Host "`nTime Summary:" -ForegroundColor White
    Write-Host "Total Billable Hours: $TotalBillableHours hours" -ForegroundColor Green
    
    # Task distribution
    Write-Host "`nTask Distribution:" -ForegroundColor White
    $Results | Where-Object { $_.Posted -eq "Yes" } | 
        Group-Object TaskTitle | 
        Sort-Object Count -Descending |
        ForEach-Object {
            Write-Host "$($_.Name): $($_.Count) entries" -ForegroundColor Cyan
        }
}

function Export-ProcessingResults {
    param (
        [Array]$Results,
        [string]$BasePath = "meeting_processing"
    )
    
    $Timestamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $DetailedReport = "${BasePath}_${Timestamp}.txt"
    $CsvReport = "${BasePath}_${Timestamp}.csv"
    
    try {
        # Create detailed report
        $Report = @"
Meeting Processing Report
Generated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')

Summary Statistics
-----------------
Total Meetings: $($Results.Count)
Matched: $(($Results | Where-Object { $_.MatchStatus -eq "Matched" }).Count)
Posted: $(($Results | Where-Object { $_.Posted -eq "Yes" }).Count)

Detailed Results
---------------

"@
        
        foreach ($Result in $Results) {
            $Report += @"
Meeting: $($Result.Meeting)
Time: $($Result.Time)
Status: $($Result.MatchStatus)
Task: $($Result.TaskTitle)
Posted: $($Result.Posted)
Duration: $($Result.BillableDuration) hours
Actual Minutes: $($Result.ActualMinutes)

"@
        }
        
        # Export reports
        $Report | Out-File -FilePath $DetailedReport -Encoding UTF8
        $Results | Export-Csv -Path $CsvReport -NoTypeInformation
        
        Write-Host "Detailed report exported to: $DetailedReport" -ForegroundColor Green
        Write-Host "CSV data exported to: $CsvReport" -ForegroundColor Green
    }
    catch {
        Write-Host "Error exporting results: $($_.Exception.Message)" -ForegroundColor Red
    }
}

function Start-MeetingProcessor {
    Write-Host "Starting Meeting Processor..." -ForegroundColor Cyan
    $StartTime = Get-Date
    $Results = New-Object System.Collections.ArrayList
    
    try {
        # Get Graph API access token
        Write-Host "Getting access token..." -ForegroundColor Cyan
        $AccessToken = Get-ApplicationAccessToken -TenantId $TenantId -ClientId $ClientId -ClientSecret $ClientSecret
        if (-not $AccessToken) { throw "Failed to get access token" }
        
        # Get current user
        $CurrentUser = Get-IntervalsCurrentUser
        if (-not $CurrentUser) { throw "Failed to get user information" }
        
        # Get tasks
        $TaskData = Get-IntervalsData -Endpoint "/task"
        if (-not $TaskData -or -not $TaskData.intervals.task.item) {
            throw "No tasks found in Intervals"
        }
        
        $Tasks = $TaskData.intervals.task.item
        
        # Build task context
        $TaskContext = Get-TaskContext -Tasks $Tasks
        
        # Get meetings
        $Meetings = Get-UserMeetings -UserId $TargetUserId -AccessToken $AccessToken
        
        if ($Meetings.Count -eq 0) {
            Write-Host "No meetings found for processing." -ForegroundColor Yellow
            return
        }

        $TotalMeetings = $Meetings.Count
        $CurrentMeeting = 0
        
        # Process each meeting
        foreach ($Meeting in $Meetings) {
            $CurrentMeeting++
            $Result = Process-Meeting -Meeting $Meeting -Tasks $Tasks -TaskContext $TaskContext -CurrentUser $CurrentUser -CurrentCount $CurrentMeeting -TotalCount $TotalMeetings
            [void]$Results.Add($Result)
        }
        
        # Show final results
        Show-ProcessingStatistics -Results $Results
        Export-ProcessingResults -Results $Results
        
        $EndTime = Get-Date
        $Duration = $EndTime - $StartTime
        Write-Host "`nTotal processing time: $($Duration.Minutes) minutes and $($Duration.Seconds) seconds" -ForegroundColor Cyan
    }
    catch {
        Write-Host "Error in main process: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Start the process
Start-MeetingProcessor
