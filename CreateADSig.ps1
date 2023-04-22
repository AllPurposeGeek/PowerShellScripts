param (
    [Parameter(Mandatory=$true)][string]$templateFile,
    [Parameter(Mandatory=$true)][string]$signatureName,
    [Parameter()][switch]$SetCurrent
)


# Check if the template file exists
if (-not (Test-Path -Path $templateFile)) {
    Write-Host "Error: Template file not found at '$templateFile'"
    exit
}

function Generate-Signature {
    param (
        $templateFilePath,
        $user,
        $domain,
        $signatureName
    )

    $domainEntry = New-Object System.DirectoryServices.DirectoryEntry("LDAP://$domain")
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($domainEntry)
    $searcher.Filter = "(&(objectCategory=User)(sAMAccountName=$user))"
    $userEntry = $searcher.FindOne().GetDirectoryEntry()

    $attributes = @{
        "{FullName}"      = $userEntry.Properties["displayName"].Value
        "{EMail}"         = $userEntry.Properties["mail"].Value
        "{Title}"         = $userEntry.Properties["title"].Value
        "{PhoneNumber}"   = $userEntry.Properties["telephoneNumber"].Value
        "{FaxNumber}"     = $userEntry.Properties["facsimileTelephoneNumber"].Value
        "{OfficeLocation}"= $userEntry.Properties["physicalDeliveryOfficeName"].Value
        "{Department}"    = $userEntry.Properties["department"].Value
        "{MobileNumber}"  = $userEntry.Properties["mobile"].Value
        "{Website}"       = $userEntry.Properties["wWWHomePage"].Value
    }

    $templateContent = Get-Content $templateFile -Raw
    $signatureContent = $templateContent

    foreach ($key in $attributes.Keys) {
        $signatureContent = $signatureContent.Replace($key, $attributes[$key])
    }

    $signaturePath = "$env:APPDATA\Microsoft\Signatures"
    if (-not (Test-Path -Path $signaturePath)) {
        New-Item -ItemType Directory -Path $signaturePath | Out-Null
    }

    $signatureFile = Join-Path $signaturePath "$signatureName.htm"
    Set-Content -Path $signatureFile -Value $signatureContent

    if ($SetCurrent) {
        $officeVersions = @("16.0", "15.0")
        ForEach ($officeVersion in $officeVersions) {
            $regKeys = @(
                "HKCU:\Software\Microsoft\Office\$officeVersion\Common\MailSettings",
                "HKCU:\Software\Microsoft\Office\$officeVersion\Outlook\Options\Mail"
            )

            foreach ($regKey in $regKeys) {
                try {
                    if (!(Test-Path -Path $regKey)) {
                        New-Item -Path $regKey -Force
                    }

                    $newItemProps = @{
                        Path  = $regKey
                        Name  = "NewSignature"
                        Value = $signatureName
                        Type  = "String"
                        Force = $true
                    }
                    New-ItemProperty @newItemProps

                    $newItemProps.Name = "ReplySignature"
                    New-ItemProperty @newItemProps

                    if ($regKey -match "Options") {
                        $newItemProps.Name  = "EnableLogging"
                        $newItemProps.Value = 0
                        $newItemProps.Type  = "DWord"
                        New-ItemProperty @newItemProps
                    }
                } catch {
                    Write-Host "Error: Unable to set registry key: $regKey"
                }
            }

            $regKeyProfiles = "HKCU:\Software\Microsoft\Office\$officeVersion\Outlook\Profiles"
            if (Test-Path -Path $regKeyProfiles) {
                $profileNames = Get-Item -Path $regKeyProfiles | Get-ChildItem | Select-Object -ExpandProperty Name
                $profileNames = $profileNames | ForEach-Object { $_.Split('\')[-1] }

                # Set signature for each profile
                foreach ($profileName in $profileNames) {
                    $regKey = "HKCU:\Software\Microsoft\Office\$officeVersion\Outlook\Profiles\$profileName\9375CFF0413111d3B88A00104B2A6676\00000002"
                    if (!(Test-Path -Path $regKey)) {
                        New-Item -Path $regKey -Force
                    }

                    $newItemProps = @{
                        Path  = $regKey
                        Name  = "New Signature"
                        Value = $signatureName
                        Type  = "String"
                        Force = $true
                    }
                    New-ItemProperty @newItemProps

                    $newItemProps.Name = "Reply-Forward Signature"
                    New-ItemProperty @newItemProps
                }
            }
        }
    }
}

$user = $env:USERNAME
$domain = $env:USERDOMAIN
$domainCheck = (Get-WmiObject -Query "SELECT * FROM Win32_ComputerSystem").PartOfDomain

if ($domainCheck) {
    $context = New-Object System.DirectoryServices.ActiveDirectory.DirectoryContext('Domain', $domain)
    try {
        $domainController = [System.DirectoryServices.ActiveDirectory.DomainController]::FindOne($context).Name
        if (Test-NetConnection -ComputerName $domainController -Port 389 -InformationLevel Quiet) {
            Generate-Signature -templateFilePath $templateFile -user $user -signatureName $signatureName -domain $domain
        } else {
            Write-Host "No connectivity to the domain controller. Signature generation skipped."
        }
    } catch {
        Write-Host "Unable to find domain controller for domain: $domain"
    }
} else {
    Write-Host "This machine is not part of a domain. Signature generation skipped."
}