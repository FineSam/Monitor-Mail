function Monitor-Mail{

    param(
        [parameter(position=0,mandatory=$true)]
        [string] $FolderName,
        [parameter(position=1,mandatory=$true)]
        [int] $MinAgo,
        [parameter(position=2,mandatory=$true)]
        [string] $MailSubject
    )

    #Get time selected
    $Time = Get-Date
    $SelectedMinAgo = $Time.AddMinutes(-$MinAgo)

    # 起動済みのOutlookがあるか確認
    $outlookProcess = Get-Process -Name "OUTLOOK" -ErrorAction SilentlyContinue
    $needQuit = $false
    if ($outlookProcess -eq $null) {
        $needQuit = $true
    }

    $outlook = New-Object -ComObject Outlook.Application
    try {
        $namespace = $outlook.GetNamespace("MAPI")
        $MailAddress = $namespace.Session.CurrentUser.Name
        $FolderPath = "\\$($MailAddress)\$($FolderName)"
        $MailBox = $namespace.Folders.Item($MailAddress).Folders | Where-Object -Property Name -EQ $FolderName
        $LastMail = $MailBox.Items |
            Sort-Object -Property ReceivedTime -Descending |
            Where-Object -Property Subject -Like $MailSubject | 
            Select-Object -Property Subject, ReceivedTime -First 1
        if ($LastMail.ReceivedTime -lt $SelectedMinAgo){
            # find random WAV file in your Windows folder
            $randomWAV = Get-ChildItem -Path C:\Windows\Media -Filter *.wav | 
              Get-Random |
              Select-Object -ExpandProperty Fullname

            # load Forms assembly to get a MsgBox dialog
            Add-Type -AssemblyName System.Windows.Forms

            # play random sound until MsgBox is closed
            $player = New-Object Media.SoundPlayer $randomWAV
            $player.Load();
            $player.PlayLooping()
            $result = [System.Windows.Forms.MessageBox]::Show("Mail `"$($MailSubject)`" is not coming`nfor more than $($MinAgo) minutes.", "Alert Mail", "Ok", "Exclamation")
            $player.Stop()

        }

    } finally {
        if ($needQuit) {
            [void]$outlook.Quit()
            [void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($outlook)
        }
    }   
}

#Monitor-Mail -FolderName "Inbox" -MinAgo 15 -MailSubject "Microsoft アカウントのセキュリティ情報*"