# send-email.ps1
# Send an email using the signed-in Outlook desktop profile.

param (
    [Parameter(Mandatory = $true)]
    [string]$To,

    [Parameter(Mandatory = $true)]
    [string]$Subject,

    [Parameter(Mandatory = $true)]
    [string]$Body,

    [Parameter(Mandatory = $false)]
    [string[]]$Attachment,

    [Parameter(Mandatory = $false)]
    [string]$Cc
)

$outlook = $null
$mail = $null

try {

    # ----------------------------------------------
    # Validate attachments before creating the mail
    # ----------------------------------------------

    if ($Attachment) {

        foreach ($file in $Attachment) {

            if (-not (Test-Path -LiteralPath $file -PathType Leaf)) {
                Write-Error "Attachment not found: $file"
                exit 2
            }
        }
    }


    # ----------------------------------------------
    # Connect to Outlook
    # ----------------------------------------------

    Write-Host ""
    Write-Host "Connecting to Outlook..."

    $outlook = New-Object -ComObject Outlook.Application


    # ----------------------------------------------
    # Create email
    # 0 = Outlook MailItem
    # ----------------------------------------------

    $mail = $outlook.CreateItem(0)

    $mail.To = $To
    $mail.Subject = $Subject
    $mail.Body = $Body

    if ($Cc) {
        $mail.CC = $Cc
    }


    # ----------------------------------------------
    # Add attachments
    # ----------------------------------------------

    if ($Attachment) {

        foreach ($file in $Attachment) {

            $resolvedPath = (Resolve-Path -LiteralPath $file).Path

            [void]$mail.Attachments.Add($resolvedPath)

            Write-Host "Attachment : $resolvedPath"
        }
    }


    # ----------------------------------------------
    # Send
    # ----------------------------------------------

    Write-Host ""
    Write-Host "To      : $To"

    if ($Cc) {
        Write-Host "CC      : $Cc"
    }

    Write-Host "Subject : $Subject"
    Write-Host ""
    Write-Host "Sending..."

    $mail.Send()

    Write-Host ""
    Write-Host "Email submitted to Outlook successfully."

    exit 0
}
catch {

    Write-Host ""
    Write-Error "Unable to send email through Outlook."
    Write-Error $_.Exception.Message

    exit 3
}
finally {

    if ($mail) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) |
            Out-Null
    }

    if ($outlook) {
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) |
            Out-Null
    }
}