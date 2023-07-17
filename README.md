
# Powershell mail-gun

This project aims at helping cybersecurity teams to send emails in mass to collaborators with simple csv files.
## Run Locally

Clone the project

```bash
  git clone https://github.com/hanteed/Powershell-mail-gun
```

Go to the project directory

```bash
  cd Powershell-mail-gun
```

Run the code with permissions :

```bash
  ./main.ps1
```

Run the code without executing permissions :

```bash
  $AnyFile = Get-Content -Path 'main.ps1' -Raw -Encoding UTF8;$ScriptBlock = [System.Management.Automation.ScriptBlock]::Create($AnyFile);& $ScriptBlock
```

## Documentation

The project will create three folders inside a "Data" folder : 

- To do (containing csv files to process)
- Done (containing processed files)
- Errors (containing files triggering errors)

Once folders are created, you can add subfolders inside the "To do" folder to split files in different parts of the tree (it can be useful if you have to send emails to multiple and distinct entities). The program will then process all the files, add their processing date to their name and move them inside the "Done" folder (or the "Errors" folder if the file is not formatted correctly). Mails are sent to a first person (first column of the file) and are in copy of a second person (second column of the file).

The program is to be used with a secure SMTP server. You will need to create credentials with windows PowerShell in order to make it work :
```
$smtpusername = Read-Host "Please enter your SMTP username" -AsSecureString;$smtppassword = Read-Host "Please enter your SMTP password" -AsSecureString 
```


csv file example :

```
Email employe;Email manager
email_employe1@example.com;email_manager1@example.com
email_employe2@example.com;email_manager2@example.com
```
## Environment Variables

To run this project, you will need to edit the following environment variables inside the main.ps1 file : 

`contacterror1`

`contacterror2`

`smtpServer`

`smtpPort`

Optional variables :

`path`

`errorpath`

`bodyTemplate`