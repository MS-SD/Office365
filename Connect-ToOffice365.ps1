Function Connect-ToOffice365([switch] $UpdateMsOnlineModule) {
    $ModuleCheck = Get-Module -ListAvailable | Where-Object {$_.Name -eq "msonline"}

    If ($UpdateMsOnlineModule) {
        $InstalledVer = $ModuleCheck.Version.ToString()
        $CurrentVer = (Find-Module -Name MsOnline).Version.ToString()

        If ($CurrentVer -gt $InstalledVer) {
            Write-Host "[ACTION]: Installing newest version of MsOnline Module"
            Start-Process powershell -Verb runAs -ArgumentList "-command &{Update-Module MSonline -Force}" -Wait
        }
        Else {
            Write-Host "[ALERT]: MsOnline module version is up to date." -f DarkYellow
        }
    }
    Else {
        Write-Verbose "Bypassing Update"
    }
    If (-not($ModuleCheck)) {
        $admin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")
        If ($admin -eq "True") {
            Write-Host "[ACTION]: Missing MsOnline Module...Starting installation." -f DarkCyan
            Install-Module MSOnline
            Connect-MsolService
        }
        else {
            Write-Host "[ACTION]: Missing MsOnline Module...Starting installation." -f DarkCyan
            Write-Host "[INFO]: Elevating Powershell to install module"
            Start-Process powershell -Verb runAs -ArgumentList "-command &{Install-Module MSonline -Force}" -Wait
            Connect-MsolService
        }
    } #If $ModuleCheck
    Else {
        If (-Not($domainCheck = Get-MsolDomain -ErrorAction SilentlyContinue)) {

            #Asks for credentials from Operator
            $UserCredential = Get-Credential -Message "Enter Office 365 tenant credentials"

            #Connects to Office 365 Admin Portal
            Write-host "[ACTION]:Connecting to Office 365" -f DarkCyan
            $pass = $false

            If ($UserCredential) {
                Connect-MsolService -Credential $UserCredential -ErrorAction Stop

                #Stores domain name in variable for output on next line
                $domain = Get-MsolDomain | Select -ExpandProperty Name

                #Makes connection to Exchange Online using credentials provided above
                Write-Host "[ACTION]:Connecting to Exchange online. Partner: $domain" -f DarkCyan
                $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
                Import-Module -Global (Import-PSSession $Session -DisableNameChecking)
                $pass = $true
            }# If
            Else {
                Write-Host "[ALERT]:Credentials were incorrect. Please run again." -f DarkYellow
                Break
            }# Else

        }# If Get-MSoldomain
        Else {
            Write-Host "[ALERT]: You are already connected to $($domainCheck.Name[0]). Please disconnect by closing and re-opening Powershell." -f DarkYellow
        }

    }# Else
}# Function Connect-ToOffice365