#Path to the folder of already processed data
$path = "Data/Done"

#Path to files triggering errors
$errorpath = "Data/Errors"

#Path to files to process
$csvFolderPath = "Data/To do"

#Path to the mail template
$bodyTemplate = "template.html"

#Emails to contact in case of error
$contacterror1 = "person1@example.com"
$contacterror2 = "person2@example.com"   

$date = Get-Date

$smtpServer = "smtp.example.com"
$smtpPort = 587

$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($smtpusername)
$result = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr) 
$username = $result

$Ptr = [System.Runtime.InteropServices.Marshal]::SecureStringToCoTaskMemUnicode($smtppassword)
$result = [System.Runtime.InteropServices.Marshal]::PtrToStringUni($Ptr)
[System.Runtime.InteropServices.Marshal]::ZeroFreeCoTaskMemUnicode($Ptr) 
$password = $result

$smtp = New-Object System.Net.Mail.SmtpClient($smtpServer, $smtpPort)
$smtp.Credentials = New-Object System.Net.NetworkCredential($username, $password)
$smtp.EnableSsl = $true

$filled = 0
$check = 0
[int]$filenumber = 1
$remove = 0


Write-Host "`nPowershell mail-gun app.`n"
if (!(test-path -PathType container $path))
{
      Write-Host ">Logs folder '$path' not found, creating it..."
      New-Item -ItemType Directory -Path $path | Out-Null
      Write-Host "✅ Folder '$path' created.`n"
}
else
{
    Write-Host "✅ Logs folder '$path' found.`n>Verifying container folder...`n"
}

if (!(test-path -PathType container $csvFolderPath))
{
      Write-Host ">Container folder '$csvFolderPath' not found, creating it..."
      New-Item -ItemType Directory -Path $csvFolderPath | Out-Null
      Write-Host "✅ Folder '$csvFolderPath' created, please insert csv files to process.`n"
      $filled = 1
}
else
{
    Write-Host "✅ Container folder '$csvFolderPath' found.`n>Verifying errors folder...`n"
}

if (!(test-path -PathType container $errorpath))
{
      Write-Host ">Folder containing errors '$errorpath' not found, creating it..."
      New-Item -ItemType Directory -Path $errorpath | Out-Null
      Write-Host "✅ Folder '$errorpath' created.`n"
      $filled = 1
}
else
{
    Write-Host "✅ Folder containing errors '$errorpath' found.`n>Sending emails...`n"
}


function Send-Email 
{
    param ([string]$MailTo,[string]$Cc,[string]$body, [string]$Subject)
    $enc = [System.Text.Encoding]::UTF8
    $message = New-Object System.Net.Mail.MailMessage
    $message.From = New-Object System.Net.Mail.MailAddress("mail@example.com", "Cybersecurity team")
    $message.To.Add($Mailto)
    $message.CC.Add($CC)
    $message.Subject = $Subject
    $message.IsBodyHtml = $true
    $message.Body = $body
    $smtp.Send($message)
}


$csvFiles = Get-ChildItem -Path $csvFolderPath -Filter *.csv -Recurse

foreach ($csvFile in $csvFiles) 
{
    $csvFilePath = $csvFile.FullName
    
    $folderName = Split-Path $csvFile.DirectoryName -Leaf

    $csvData = Import-Csv -Path $csvFilePath -Delimiter ";"

    Write-Host "File n.$filenumber ($csvFile) in $folderName" 
    $filenumber++

    foreach ($row in $csvData) 
    {
        $employeeEmail = $row."Email employe"
        $managerEmail = $row."Email Manager"
        
        $filled = 1

        $firstWord = $employeeEmail.Split(' ')[0].Trim()

        if ($firstWord -ne 'Date') 
        {    
            if (![string]::IsNullOrEmpty($employeeEmail) -and ![string]::IsNullOrEmpty($managerEmail)) 
            {
                $body = Get-Content -Path $bodyTemplate -Raw -Encoding UTF8
                Send-Email -MailTo $employeeEmail -Cc $managerEmail -body $body -Subject "This is the subject"
                Write-Host ">Mail sent to $employeeEmail in copy to manager $managerEmail"
                $check = 1
            }
            else
            {
                $errordate = $date = Get-Date -Format "yyyy-MM-dd HH-mm-ss"
                $errorbody = @"
                    Hello,<br><br>
                    There was an error while processing one of the files.<br><br>
                    <b>Subfolder</b> : $($folderName)<br>
                    <b>File</b> : $($csvFile)<br>
                    <b>Date of the error</b> : $($errordate)<br><br>
                    Please make sure you use the ";" separator inside the csv file.<br><br>
                    Sincerely yours,<br><br>
                    The cybersecurity team.<br><br>
"@
                Write-Host "❌ Line wrongly formatted (make sure the fields of the file are separated by ';')."
                Send-Email -MailTo $contacterror1 -Cc $contacterror2 -body $errorbody -Subject "Error, mail-gun" 
                $check = 0
            }
        }
        else
        {
            $remove = 1
        }

    }

    $logFolderPath = Join-Path -Path $path -ChildPath $folderName
    if (!(Test-Path $logFolderPath))
    {
        $null = New-Item -ItemType Directory -Force -Path $logFolderPath
    }

    if ($check -eq 1) 
    {
        $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $destinationFileName = $csvFile.BaseName + "_" + $date + $csvFile.Extension
        $destinationFilePath = Join-Path -Path $logFolderPath -ChildPath $destinationFileName
        Move-Item -Path $csvFile.FullName -Destination $destinationFilePath

        if (Test-Path $destinationFilePath) 
        {
            Write-Host "File $csvFile moved with success inside the logs folder.`n"
        } 
        else 
        {
            Write-Host "❌ Error : Impossible to move file $csvFile in logs folder.`n"
        }
    }
    else
    {
        $errorFolderPath = Join-Path -Path $errorpath -ChildPath $folderName
        if (!(Test-Path $errorFolderPath))
        {
            $null = New-Item -ItemType Directory -Force -Path $errorFolderPath
        }
        $date = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
        $destinationFileName = $csvFile.BaseName + "_" + $date + $csvFile.Extension
        $destinationFilePath = Join-Path -Path $errorFolderPath -ChildPath $destinationFileName
        Move-Item -Path $csvFile.FullName -Destination $destinationFilePath

        if (Test-Path $destinationFilePath) 
        {
            Write-Host "File $csvFile move with success in errors folder.`n"
        } 
        else 
        {
            Write-Host "❌ Error : Impossible to move file $csvFile in errors folder.`n"
        }
    }
}

if ($filled -eq 0)
{
    Write-Host "❔ Folder '$csvFolderPath' seems to be empty, please insert some csv files to process."
}

Write-Host "Program terminated.`n"