# Outlook Email Automation

This folder contains a PowerShell script that creates and sends an email through
the currently signed-in Microsoft Outlook desktop profile. It can add one or
more local files as attachments. It does not use SMTP directly, store email
credentials, or send through a cloud API.

## What the automation does

At a high level, the script performs this sequence:

1. Accepts the recipient, subject, and body as required command-line parameters.
2. Optionally accepts a CC recipient and one or more attachment paths.
3. Checks that every requested attachment exists before connecting to Outlook.
4. Starts Outlook COM automation and creates a new mail item.
5. Populates the message fields and adds the resolved attachment paths.
6. Calls Outlook's `Send()` method and reports whether Outlook accepted the
   send request.
7. Releases the COM references used by the script.

The script is [email-send.ps1](./email-send.ps1). `test.txt` and `test1.txt`
are sample text files that can be used to try the attachment option.

## Requirements

- Windows with PowerShell.
- Microsoft Outlook desktop installed and configured with a signed-in profile.
- Permission for the current Outlook profile to send email.

The script uses the Outlook COM interface, so it is not a cross-platform
PowerShell mailer. Outlook security policy, an offline profile, authentication
issues, or organizational restrictions can prevent sending.

## Run it

Open PowerShell in this folder and call the script with the required values:

```powershell
.\email-send.ps1 `
  -To "recipient@example.com" `
  -Subject "Automation test" `
  -Body "Hello, this message was sent by the Outlook automation."
```

To include CC and multiple attachments:

```powershell
.\email-send.ps1 `
  -To "recipient@example.com" `
  -Cc "copy@example.com" `
  -Subject "Files for review" `
  -Body "Please see the attached files." `
  -Attachment ".\test.txt", ".\test1.txt"
```

You can also use an absolute path for the script and files if the current
directory is different:

```powershell
& "C:\path\to\EmailAutomation\email-send.ps1" `
  -To "recipient@example.com" `
  -Subject "Automation test" `
  -Body "Message body" `
  -Attachment "C:\path\to\report.pdf"
```

`-To`, `-Subject`, and `-Body` are mandatory. `-Cc` and `-Attachment` are
optional. The attachment parameter accepts a string array, so pass paths
separated by commas.

## How to interpret the result

On success the script prints the recipients, subject, attachment paths, and
`Email submitted to Outlook successfully.` It exits with code `0`. A missing
attachment is reported before Outlook is opened and exits with code `2`.
Other Outlook or script errors are reported and exit with code `3`.

Success means Outlook accepted the `Send()` request; it is not a delivery
receipt. The message may remain in the Outbox while Outlook is offline, and a
recipient server can still reject or filter it. Check Outlook's Outbox and Sent
Items, and verify the recipient address, when delivery matters.

## Troubleshooting

- **Outlook cannot be created:** Confirm the Outlook desktop application is
  installed and has a working profile. Check that the script is running under
  the Windows account associated with that profile.
- **Attachment not found:** Check the spelling and current working directory,
  or supply a fully qualified path. The script requires every attachment to
  exist as a file before creating the message.
- **The command is blocked by script policy:** Use your organization's approved
  PowerShell execution policy. Do not weaken machine-wide security policy just
  to run this script.
- **The command reports success but no message arrives:** Check Outlook's
  Outbox/Sent Items, network connectivity, junk-mail filtering, and any
  mail-flow rules or delivery failure notices.

## Safety

This script sends immediately; it does not open a draft for review. Test first
with an address you control and inspect the recipient, subject, body, and
attachments before using it for real messages. Avoid passing sensitive message
content directly on a shared command line or saving it in shell history.
