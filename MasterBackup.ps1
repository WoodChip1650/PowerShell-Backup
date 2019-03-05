$basedir = "Q:\KDHE"
$today   = (Get-Date).ToString('MM.dd.yy')
$exclude = @("*.bin","*.onesrvcache","*.ini",".849C9593-D756-4E56-8D6E-42412F2A707B","*Local Settings*","*AppData*","*3D Objects*","*MicrosoftEdgeBackups*","*Music*","*.dat")
$location = New-Item -Path $basedir -Type Directory -Name $today
$outlook = "C:\Users\jonathan.wood\AppData\Local\Microsoft\Outlook"
$finaldir19 = "Q:\KDHE\_Archives\2019"
$lrt = "Z:\Operations & Planning\Safety Program\Lay Responder Team"
$trg = "Z:\Training, Exercising & Evaluation\Training\ICS Training"
$outfile = "Q:\KDHE\KDHE_$today.7z"
$7z = "C:\7-Zip\7z.exe"

Copy-Item $lrt -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item $trg -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item 'C:\Users\jonathan.wood\OneDrive Home Folders' -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item 'C:\Users\jonathan.wood\Contacts' -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item 'C:\Users\jonathan.wood\Documents' -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item 'C:\Users\jonathan.wood\Downloads' -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item 'C:\Users\jonathan.wood\Pictures' -Destination $location -Recurse -Exclude $exclude -PassThru
Copy-Item 'C:\Users\jonathan.wood\Videos' -Destination $location -Recurse -Exclude $exclude -PassThru
Write-Host 'Close Outlook, Skype, and OneNote.  If errors close manually...'
Pause
Copy-Item $outlook -Destination $location\Outlook -Recurse -Exclude $exclude -PassThru
Write-Host "Backup to FlashDrive Complete"
Write-Host "Compressing Folder"
#Remove Empty Folders
$SB = {
    Get-ChildItem $location -Directory -Recurse |
    Where-Object {(Get-ChildItem $_.FullName -Force).Count -eq 0}
}
For ($Empty = &$SB ; $Empty -ne $null ; $Empty = &$SB) {Remove-Item (&$SB).FullName}
# 7-zip Arguments
[array]$7zArgs=@("a", "-t7z", "-r", "-mx=7", "-ssw", "-mtc=on", $outfile) + $location 
# Zip it up and save 7z output 
[string]$ZipFile = & $7z $7zArgs -PassThru
#Move Compressed Folder to Year Folder
Move-Item $outfile -Destination $finaldir19
#Clean Up
Get-ChildItem -Path $basedir\$today -Recurse | Remove-Item -force -recurse
Remove-Item $basedir\$today -Force 
