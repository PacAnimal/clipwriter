################################################################################
# First some settings...                                                       #
################################################################################
$ErrorActionPreference = 'Stop' # playing it safe?
$maxPasteList = 10 # how many lines to list in the operation preview
$defaultSleep = 10 # how many seconds to sleep by default
$fileBufferSize = 1000 # how many bytes to transfer in one command when pasting files


################################################################################
# We use System.Windows.Forms.SendKeys.SendWait to send stuff                  #
################################################################################
Add-Type -AssemblyName System.Windows.Forms
Function SendRaw([string]$keys) { [System.Windows.Forms.SendKeys]::SendWait($keys) }
Function SendLine([string]$line) { Send($line); SendRaw("{ENTER}") }
Function Send([string]$line)
{
    $sb = [System.Text.StringBuilder]::new()
    for ($i = 0; $i -lt $line.Length; $i++)
    {
        $c = $line[$i]
        switch($c)
        {
            "+" { $sb.Append("{+}") | Out-Null }
            "^" { $sb.Append("{^}") | Out-Null }
            "%" { $sb.Append("{%}") | Out-Null }
            "~" { $sb.Append("{~}") | Out-Null }
            "(" { $sb.Append("{(}") | Out-Null }
            ")" { $sb.Append("{)}") | Out-Null }
            "[" { $sb.Append("{[}") | Out-Null }
            "]" { $sb.Append("{]}") | Out-Null }
            "{" { $sb.Append("{{}") | Out-Null }
            "}" { $sb.Append("{}}") | Out-Null }
            default { $sb.Append($c) | Out-Null }
        }
    }
    SendRaw($sb.ToString())
}


################################################################################
# How to exit                                                                  #
################################################################################
Function Enter-Exit
{
    Write-Host -ForegroundColor Gray "Press <ENTER> to quit."
    Read-Host | Out-Null
    Exit 0
}


################################################################################
# How to write stuff to the screen                                             #
################################################################################
Function Info([string]$text) { Write-Host -ForegroundColor Cyan "$text" }
Function Spam([string]$text) { Write-Host -ForegroundColor DarkCyan "$text" }
Function Error([string]$text) { Write-Host -ForegroundColor Red "$text"; Enter-Exit }


################################################################################
# Grab the current clipboard and list it to the user                           #
################################################################################
Write-Host
if ($text = Get-Clipboard -Format Text -TextFormatType UnicodeText)
{
    $lines = New-Object System.Collections.Generic.List[System.String]
    $text | ForEach-Object `
    {
        $lines.Add($_)
    }
    $activity = "About to paste $($lines.Count) line$(if ($lines.Count -ne 1) {"s"}) of text"
    Info "${activity}:"
    $lines | Select -First $maxPasteList | ForEach-Object { Spam $_ }
    if ($lines.Count -gt $maxPasteList) { Spam "[...]" }
}
elseif ($dropList = Get-Clipboard -Format FileDropList)
{
    $files = @(
        $dropList | ForEach-Object `
        {
            $dir = $_.Directory.FullName
            $(
                Get-Item $_.FullName
                if (Test-Path -Path $_.FullName -PathType Container)
                {
                    Get-ChildItem -Path $_.FullName -Force -Recurse
                }
            ) |
            Where-Object { $_.FullName.StartsWith($dir) } |
            ForEach-Object { New-Object PSObject -Property (@{ "FullName" = $_.FullName; "TargetName" = $_.FullName.Substring($dir.Length + 1).Replace("``", "````") }) }
        }
    )
    $activity = "About to paste $($files.Count) FILE$(if ($files.Count -ne 1) {"S"})/DIRECTOR$(if ($files.Count -ne 1) {"IES"} else {"Y"})"
    Info "${activity}:"
    $files | Select -First $maxPasteList | ForEach-Object { Spam $_.FullName }
    if ($files.Count -gt $maxPasteList) { Spam "[...]" }
}
else
{
    Error "No idea how to deal with your current clipboard :("
}


################################################################################
# Let the user specify a sleep before the paste: this is your chance to cancel #
################################################################################
Write-Host
Write-Host "How long, in seconds, should I wait before pasting?"
Write-Host "Press <ENTER> for default: $defaultSleep seconds"
$seconds = Read-Host -Prompt "Seconds"
if ($seconds -eq "")
{
    $seconds = $defaultSleep
}
elseif ($seconds -notmatch "^[0-9]+$" -or $seconds -le 0)
{
    Error "Invalid sleep time: $seconds"
}


################################################################################
# Quiet before the storm...                                                    #
################################################################################
Write-Host
Info "Sleeping for $seconds seconds before pasting..."
for ($i = 0; $i -lt $seconds; $i++)
{
    Write-Progress -Activity "$activity" -Status "Pasting in: $($seconds - $i) seconds" -PercentComplete ( $i / $seconds * 100)
    Start-Sleep -Seconds 1
}
Write-Progress -Activity "$activity" -Status "Pasting in: $($seconds - $i) seconds" -PercentComplete 100
Write-Progress -Activity "$activity" -Completed


################################################################################
# Pasting text is rather easy, so we deal with that first                      #
################################################################################
if ($text)
{
    Write-Host
    Info "Paste all over the place!"
    $activity = "Pasting"
    for ($i = 0; $i -lt $lines.Count; $i++)
    {
        Write-Progress -Activity "$activity" -Status "Line ${i}/$($lines.Count)" -PercentComplete ( $i * 100 / $lines.Count)
        if ($i) { SendLine }
        Send($lines[$i])
    }
    Write-Progress -Activity "$activity" -Status "Line ${i}/$($lines.Count)" -PercentComplete 100
    Write-Progress -Activity "$activity" -Completed
    Enter-Exit
}


################################################################################
# Pasting files, on the other hand, is hard...                                 #
################################################################################
if ($dropList)
{
    Write-Host
    Info "This will take a while... $($files.Count)"
    $activity = "Writing files"
    for ($i = 0; $i -lt $files.Count; $i++)
    {
        $file = $files[$i]
        $status = "Item $($i+1)/$($files.Count): $($file.TargetName) .."
        Write-Progress -Activity "$activity" -Status $status -PercentComplete 0
        if (Test-Path -Path $file.FullName -PathType Container)
        {
            SendLine("New-Item -ItemType Directory -Force -Path `"$($file.TargetName)`" | Out-Null")
        }
        else
        {
            $size = (Get-Item -Path $file.FullName).Length
            $totalBytesRead = 0
            $fileStream = [System.IO.File]::OpenRead($file.FullName)
            $buffer = New-Object byte[] $fileBufferSize
            SendLine("New-Item -ItemType File -Force -Path `"$($file.TargetName)`" | Out-Null")
            while ( $bytesRead = $fileStream.Read($buffer, 0, $fileBufferSize) ){
                $base64 = [Convert]::ToBase64String($buffer, 0, $bytesRead);
                SendLine("Add-Content -Encoding Byte -Path `"$($file.TargetName)`" -Value ([Convert]::FromBase64String(`"$base64`"))")
                $totalBytesRead += $bytesRead
                Write-Progress -Activity "$activity" -Status $status -PercentComplete $($totalBytesRead * 100 / $size)
            }
            $fileStream.Close()
        }
        Write-Progress -Activity "$activity" -Status $status -PercentComplete 100
    }
    Write-Progress -Activity "$activity" -Completed
    Enter-Exit
}


################################################################################
# We shouldn't end up down here, unless we're half-way supporting something    #
################################################################################
Write-Host
Error "Ended up at the bottom of the script?! We should have exited before this..."