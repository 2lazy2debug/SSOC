# Create Outlook COM object
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$curfolder = $namespace.GetDefaultFolder(6)
 
$global:showUnreadOnly = $true
$global:lastupdate = get-date
 
##### Folders Navigation and Management #####
function Get-MailFolders($folder, $indent = "") {
    
    if ($folder.DefaultItemType -eq 0) {
        Write-Output "$indent$($folder.Name)"
    }
 
    foreach ($subfolder in $folder.Folders) {
        Get-MailFolders -folder $subfolder -indent ("$indent  ")
    }
}
# TODO : Implement interactive folder selection 
function Change-MailboxFolder {
    Clear-Host 
    $root = $namespace.Folders.Item(1)
    Get-MailFolders -folder $root
 
    [void][System.Console]::ReadKey($true)
}
 
function Update-MessagesList { 
    # Filter and sort unread messages
    if ($global:showUnreadOnly) {
        $global:unreadMessages = $curfolder.Items | Where-Object { $_.UnRead -eq $true } | Sort-Object ReceivedTime -Descending
    }
    else { 
        $global:unreadMessages = $curfolder.Items | Sort-Object ReceivedTime -Descending
    }
    $global:selectedIndex = 0
    $global:lastupdate = get-date
}

function Show-EmailList {
    Clear-Host
    Write-Host "📬 SSOC - Last update 🕑: $global:lastupdate - Showing $(if($global:showUnreadOnly) {"unread only"} else {"all inbox"}) 
    - ⬆️⬇️ to navigate
    - 'u' to update the list
    - 't' to toggle view between unread/all
    - 'Enter' to view e-mail
    - 'd' to delete e-mail
    - 'm' to mark e-mail as read/unread
    - 'f' quick step : 'Fermé'
    - 'q' to quit`n"
    #- 'l' to change folder
    
    for ($i = 0; $i -lt $global:unreadMessages.Count; $i++) {
        $msg = $global:unreadMessages[$i]
        $from = $msg.SenderName
        $subject = if($msg.Subject.Length -gt 50) {$msg.subject.substring(0,50) + " ..."} else { $msg.subject}
        $received = $msg.ReceivedTime.ToString("dd.MM.yyyy - HH:mm")
        $unread = $msg.UnRead
        $important = 
        $hasAttachments = HasRealAttachments($msg)
 
        $cleanBody = $msg.Body -replace "(\r\n|\r|\n){2,}", "" # Collapse multiple line breaks
        $cleanBody = $cleanBody -replace "(\r\n|\r|\n)", "" # Normalize to single newlinestt
        $cleanBody = if ($cleanBody.Length -gt 50) { $cleanBody.Substring(0, 50) + " ..." } else { $cleanBody }
        
        if ($i -eq $global:selectedIndex) {
            Write-Host "➡️ |$(if($hasAttachments) {"📎"} else {"  "}) $(if($unread) {"✉️"} else {"✅"}) $received | $from | $subject | $cleanBody" -ForegroundColor Cyan
        }
        else {
            Write-Host "   |$(if($hasAttachments) {"📎"} else {"  "}) $(if($unread) {"✉️"} else {"✅"}) $received | $from | $subject | $cleanBody"
        }
    }
}
 
##### Email Actions #####
function Show-FullEmail {
    Clear-Host
    
    $msg = $global:unreadMessages[$global:selectedIndex]
    Write-Host "`nFrom: $($msg.SenderName)"
    Write-Host "Subject: $($msg.Subject)"
    Write-Host "Received: $($msg.ReceivedTime.ToString("dd.MM.yyyy - HH:mm"))"
    if ($msg.unread) {
        Write-Host "Read: No" -ForegroundColor Green
    }
    else { 
        Write-Host "Read: YES" -ForegroundColor Red
    }
    
    $cleanBody = $msg.Body -replace "(\r\n|\r|\n){2,}", "`n" # Collapse multiple line breaks
    $cleanBody = $cleanBody -replace "(\r\n|\r|\n)", "`n" # Normalize to single newlinestt
    Write-Host "`n$($cleanBody)"
    
    Write-Host "`n[Available actions] : 'd' to delete` | 'm' to mask as read/unread` | 'r' to type a quick reply` | 'o' to open in Outlook | any key to come back to inbox" -ForegroundColor Yellow
    $key = [System.Console]::ReadKey($true)
    
    
    switch ($key.Key) {
        'd' { Delete-Email }
        'm' { Toggle-Read-Mark }
        'r' { Reply }
        'o' { Open-InOutlook }
    }
    
}

function Reply {
    $msg = $global:unreadMessages[$global:selectedIndex]
    $reply = $msg.Reply()
    
    Write-Host "`nReplying to $($msg.SenderName)'s mail with subject : $($msg.Subject)"
    Write-Host "Compose your reply below. Type 'EOF' on a new line to finish.`n"
 
    $lines = @()
    while ($true) {
        $line = Read-Host
        if ($line -eq "EOF") { break }
        $lines += $line
    }
 
    $replyBody = ($lines -join "`n") -replace "`n", "`r`n"
 
    $reply.Body = "$replyBody`r`n`r`n" + $reply.Body
 
    Write-Host "`nReply composed. Press 's' to send or 'c' to cancel."
 
    $key = [System.Console]::ReadKey($true).KeyChar
    while ($true) {
        switch ($key) {
            's' {
                $reply.Send()
                $msg.Unread = $false
                Write-Host "`n✅ Reply sent. Press any key to continue..."
                [void][System.Console]::ReadKey($true)
                break
            }
            'c' {
                Write-Host "`n❌ Reply canceled. Press any key to continue..."
                [void][System.Console]::ReadKey($true)
                break
            }
            default {
                Write-Host "Press 's' to send or 'c' to cancel."
            }
        }
        break
    }
}
 
function Delete-Email { 
    $msg = $global:unreadMessages[$global:selectedIndex]
    $msg.Delete()
}
 
function Toggle-Read-Mark { 
    $msg = $global:unreadMessages[$global:selectedIndex]
    $msg.UnRead = !($msg.UnRead)
}
 
 
function HasRealAttachments($mailItem) {
    foreach ($attachment in $mailItem.Attachments) {
        # Try to detect inline attachments (usually have a ContentID and are referenced in the HTML body)
        $propertyAccessor = $attachment.PropertyAccessor
        $contentId = $null
        try {
            $contentId = $propertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001F")
        }
        catch {
            # Property might not exist, ignore
        }
 
        # If there's no ContentID or it's not referenced in the HTML body, it's likely a real attachment
        if (-not $contentId -or ($mailItem.HTMLBody -notmatch [regex]::Escape($contentId))) {
            return $true
        }
    }
    return $false
}
 

 
function Quick-Step-Ferme { 
    $msg = $global:unreadMessages[$global:selectedIndex] 
    $msg.Unread = $false
    [void]$msg.move($curfolder.Folders.Item("Fermé"))
}
 
function Open-InOutlook {
    $msg = $global:unreadMessages[$global:selectedIndex]
    $msg.Display()
}   
 
#Start here 
Update-MessagesList
 
# Main loop
do {
    Show-EmailList
    $key = [System.Console]::ReadKey($true)
    
    switch ($key.Key) {
        'UpArrow' { if ($global:selectedIndex -gt 0) { $global:selectedIndex-- } }
        'DownArrow' { if ($global:selectedIndex -lt $global:unreadMessages.Count - 1) { $global:selectedIndex++ } }
        'Enter' {
            Show-FullEmail 
            Update-MessagesList
        }
        'd' {
            Delete-Email
            Update-MessagesList
        }
        'm' {
            Toggle-Read-Mark
            Update-MessagesList
        }
        
        'u' { Update-MessagesList }
        't' { 
            $global:showUnreadOnly = !$global:showUnreadOnly
            Update-MessagesList
        }
        'f' { 
            Quick-Step-Ferme
            Update-MessagesList
        }
        <#'l' { 
            Change-MailboxFolder
        }#>
    }
} while (($key.Key -ne 'q'))
Clear-host

    
