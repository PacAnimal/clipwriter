# ClipWriter
A PowerShell script that writes your current clipboard content (text or files) as keyboard input, used to transfer files, scripts and other data through VPN clients and other remote connections when the usual methods of file transfer, including regular clipboard pasting, are not available

## How to use it
1. Copy some text, or one or more files and directories
2. If you copied text, make sure you have a window ready in which ClipWriter should type what you copied.

   If you copied files, make sure you have an open PowerShell window, or an editor for ClipWriter to either execute its commands directly or to save the script for later. The files and directories will be created at the location the PowerShell window is currently at, so make sure you change your working directory to where you want them before continuing.
3. Enter a timeout for the script
4. Switch to the target window before the timeout runs out
5. Watch ClipWriter type :)

## What's supported
* Copying large amounts of text and long scripts that would otherwise take ages to type in
* Copying multiple files and directories

## How it works
ClipWriter uses the SendWait function from the System.Windows.Forms library to generate synthetic keystrokes. It uses SendWait() in favour of alternative calls such as SendKeys() in attempt to avoid mistakes when writing large amounts of text to slow applications. SendWait() is slower than SendKeys(), but also more accurate.

To paste files, ClipWriter generates PowerShell commands to create the required directory structure, and it transfers the files themselves by Base64-encoding them in chunks and generating commands to decode the chunks on the remote side.

## Issues
* Speed, obviously. Copying large files is time consuming, but if the alternative is to run through several hours of meetings about firewall rules and security issues, this might be quicker after all.
* The computer that's doing the "typing" is pretty much unusable while the script runs. You can watch Youtube on a separate screen though :D
