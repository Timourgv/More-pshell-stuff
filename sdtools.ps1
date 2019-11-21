Import-Module activedirectory

$DomainName = (Get-ADDomain).DNSRoot
Write-Output "Domain = $DomainName"

$PreciseServer = (Get-ADDomain).PDCEmulator #Used for -precise switch
Write-Output "PDC emulator = $PreciseServer"

# Add servers to search through if 'all' is provided for -ComputerName. Used by Get-SDLockouts and Get-SDEvent
$DCs = (Get-ADDomainController -Filter *).HostName | sort

# Converts all date type properties to this format. Uses Unix formatting - https://ss64.com/bash/date.html#format
$TimeFormat = '%Y/%m/%d %H:%M:%S'


function GetSectionOutput($Objects, $OutputTemplate) {
    # Format output, used by Get-SDUser and Get-SDComputer
    $Padding = 21 # Controls left indentation
    $Color = 'DarkGray' # Color of property output. Valid choices are "Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White".

    foreach ($Object in $Objects) {
        foreach ($SectionKey in $OutputTemplate.keys) { #Cycle through HashTable keys
            Write-Host ($SectionKey.PadLeft($Padding,'/')) #Outputs section name
            foreach ($Property in $OutputTemplate.$SectionKey) { #Cycle through each $Property in section's Property list
                if (($Object.$Property) -is [System.DateTime]) { #Checks if $Property type is of date
                        $Object.$Property = Get-Date $Object.$Property -UFormat $TimeFormat} #Transforms Time/Date format
                
                $FirstElement = $true #Used to mark first element if property contains more than one value
                
                foreach ($PropertyItem in $Object.$Property) {
                    If ($FirstElement) {
                        $ResultString = ("{0,$Padding} = {1}" -f $Property, $PropertyItem)
                        $FirstElement = $false}
                    Else {
                        $ResultString = (' '*($Padding+3) + $PropertyItem)}
                    Write-Host ($ResultString) -ForegroundColor $Color
                }
            }
        }
    Write-Host (('\'*$Padding) + "`n") # Ends user output
    }
}

function Get-SDUser {
    param(
        [string]$Id,
        [switch]$Precise=$false,
        [int]$ListSG=0,
        [int]$ListDirectReports=0,
        [int]$ListManagedObjects=0,
        [string]$Server,
        [switch]$ShowAllProperties = $false,
        [array]$CustomProperty)
    # Contains Sections and Properties that will be shown on output
    $OutputTemplate = New-Object System.Collections.Specialized.OrderedDictionary
    
    # Sections (Details, AccountDetails, etc.) containing AD user object properties
    $OutputTemplate.Add('Details', ('DisplayName', 'Name', 'SamAccountName','UserPrincipalName','EmailAddress','MobilePhone','telephoneNumber','Manager', 'Department', 'Office', 'Title'))
    $OutputTemplate.Add('AccountDetails', ('Enabled', 'CanonicalName','DistinguishedName','HomeDirectory','WhenCreated','WhenChanged', 'AccountExpirationDate', 'logonCount'))
    $OutputTemplate.Add('PW Details', ('PasswordExpired','LockedOut','badPwdCount','badPasswordTime', 'pwdLastSet','pwAge', 'PasswordNeverExpires'))
    $OutputTemplate.Add('Extra', ('lastLogon','lastLogonTimestamp','proxyaddresses', 'targetAddress','msRTCSIP-UserEnabled','msRTCSIP-PrimaryUserAddress', 'msRTCSIP-FederationEnabled','memberof','directReports','managedObjects','info'))
    # To change order/add/remove new properties or sections just edit above lines
    if ($CustomProperty) {
        # Add Custom section if Property argument is provided. Actual properties will be added later
        $OutputTemplate.Add('Custom',1) }

    
    If (!$Server) {
        If ($Precise) {
            $Server = $PreciseServer}
        Else {$Server = $DomainName}}
    switch ($Id) { # Determines ID type - SAM, UPN or Full Name with Regex
        {$Id -match '^smtp:'} {$SearchBy = 'proxyaddresses'; Break}
        {$Id -match '^sip:'} {$SearchBy = 'proxyaddresses'; Break}
        {$Id -match '^disp:'} {$SearchBy = 'DisplayName'; $Id = $Id.Substring(5); Break}
        {$Id -match '^sam:'} {$SearchBy = 'SamAccountName'; $Id = $Id.Substring(4); Break}
        {$Id -match '^name:'} {$SearchBy = 'Name'; $Id = $Id.Substring(5); Break}
        {$Id -match '^emp:'} {$SearchBy = 'EmployeeNumber'; $Id = $Id.Substring(4); Break}
        {$Id -match '@'} {$SearchBy = 'UserPrincipalName'; Break}
        {$Id -match ' '} {$SearchBy = 'Name'; Break} #\p{L} matches any letter
        default {$SearchBy = 'SamAccountName'} }
    
    # Search if ID type determined
    if ($SearchBy) {
        Write-Host "Searching by $SearchBy = $ID"
        foreach ($OriginalUser in (Get-ADUser -Filter {$SearchBy -like $Id} -Properties * -Server $Server)) {
            $User = @{}
            # Add additional properties if -Property attribute is provided
            if ($CustomProperty) {
                $OutputTemplate.Item('Custom') = ($OriginalUser | select $CustomProperty | Get-Member -MemberType Properties).Name
            }
            
            $Properties = $OutputTemplate.Keys | %{$OutputTemplate.$_} # Extract keys from hash table to cycle through
            foreach ($Property in $Properties) {
                switch ($Property) { #format property
                    # Custom properties that need to be calculated
                    {$Property -eq 'pwAge'} { # Calculate and save as string
                       $User.$Property = (New-TimeSpan -Start ([datetime]::FromFileTime($OriginalUser.pwdLastSet)))
                       $User.$Property = "{0} days {1} hrs {2} minutes" -f ($User.$Property.days, $User.$Property.Hours, $User.$Property.Minutes); Break}

                    # Check if AD user has attribute
                    {!$OriginalUser.Contains($Property) -and $ShowAllProperties} {
                        $User.$Property = "<Attribute does not exist>"; Break}
                    
                    # Transform int64 values to date 
                    {$OriginalUser.$Property -is [System.Int64] -and $Property -notlike "msExch*"} {
                        $User.$Property = Get-Date ([datetime]::FromFileTime($OriginalUser.$Property)); Break}
                    
                    # Select only those values in ProxyAddresses that start with SIP or SMTP
                    {$Property -eq 'ProxyAddresses'} {
                        $User.$Property = $OriginalUser.$Property | sort | ?{$_ -match '^[smtp|sip]'}; Break}
                                            
                    # Handle -ListSG -ListDirectReports -ListManagedObjects parameters
                    {(($Property -eq 'memberof') -and ($ListSG -eq 0)) -or
                        (($Property -eq 'directReports') -and ($ListDirectReports -eq 0)) -or
                        (($Property -eq 'managedObjects') -and ($ListManagedObjects -eq 0))} {Break} #hide these properties if switch is set to 0 (default beahavior)
                    

                    {(($Property -eq 'memberof') -and ($ListSG -eq 1)) -or
                        ($Property -eq 'Manager') -or
                        (($Property -eq 'directReports') -and ($ListDirectReports -eq 1)) -or
                        (($Property -eq 'managedObjects') -and ($ListManagedObjects -eq 1))} {
                            $User.$Property = $OriginalUser.$Property | sort | %{($_ -split "CN=|,OU=|,CN=")[1]}; Break}
                    
                    {(($Property -eq 'memberof') -and ($ListSG -eq 2)) -or
                        (($Property -eq 'directReports') -and ($ListDirectReports -eq 2)) -or
                        (($Property -eq 'managedObjects') -and ($ListManagedObjects -eq 2))} {
                            $User.$Property = $OriginalUser.$Property | sort; Break}

                    {(($Property -eq 'memberof') -and ($ListSG -eq 3))} {
                        $User.$Property = $OriginalUser.$Property | sort | 
                            %{$Group = Get-ADGroup $_ -Properties Members; $Group.Name + ' (' + $Group.GroupCategory + ')' + ', has ' + $Group.Members.count + ' members'} }

                    default {$User.$Property = $OriginalUser.$Property}
                }
            }
            
            GetSectionOutput $User -OutputTemplate $OutputTemplate
        }
    } Else {Write-Host "Please provide with ID"}
}

function Get-SDComputer {
    param (
        [string]$Identity,
        [switch]$Precise=$false,
        [switch]$ObjectOutput=$false)

    if ($Precise) {$ServerName = $PreciseServer} else {$ServerName = $DomainName}

    if ($Identity) {
        switch ($Identity) {
            {$Identity -match '^desc:'} {$SearchBy = 'Description'; $Identity = $Identity.Substring(5); Break}
            default {$SearchBy = 'Name'}
        }
    Write-Host "Searching by $SearchBy = $Identity"
    $PCList = Get-ADComputer -Filter {$SearchBy -like $Identity} -Properties * -Server $ServerName
    } else {return "Error: Nothing was provided"}

    If ($PCList) {
        $PCList = $PCList | select Name, Description, OperatingSystem, OperatingSystemVersion, whenCreated, whenChanged, `
        @{n='lastLogon';e={[datetime]::FromFileTime($_.lastLogon)}}, lastLogonDate, logonCount, `
        @{n='pwdLastSet';e={[datetime]::FromFileTime($_.pwdLastSet)}}, ms-Mcs-AdmPwd, `
        @{n='ms-Mcs-AdmPwdExpirationTime';e={[datetime]::FromFileTime($_."ms-Mcs-AdmPwdExpirationTime")}}, PasswordExpired
        
        if ($ObjectOutput) {
            $PCList
        } else {
            $OutputTemplate = New-Object System.Collections.Specialized.OrderedDictionary
            $OutputTemplate.Add('Description', ('Name', 'Description', 'OperatingSystem', 'OperatingSystemVersion'))
            $OutputTemplate.Add('When', ('whenCreated', 'whenChanged', 'lastLogon', 'lastLogonTimestamp'))
            $OutputTemplate.Add('Extra', ('logonCount', 'pwdLastSet','ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime','PasswordExpired'))
            GetSectionOutput $PCList -OutputTemplate $OutputTemplate
        }
    } else {return "Error: Nothing was found"}
}

function Get-SDGroup {
    param(
        [string]$Id,
        [switch]$ObjectOutput=$false,
        [switch]$MemberInfo
    )
    switch ($Id) { # Determines ID type - SAM, UPN or Full Name with Regex
        {$Id -match '^smtp:'} {$SearchBy = 'proxyaddresses'; Break}
        {$Id -match '^desc:'} {$SearchBy = 'Description'; $Id = $Id.Substring(5); Break}
        {$Id -match '^disp:'} {$SearchBy = 'DisplayName'; $Id = $Id.Substring(5); Break}
        {$Id -match '^sam:'} {$SearchBy = 'SamAccountName'; $Id = $Id.Substring(4); Break}
        {$Id -match '^name:'} {$SearchBy = 'Name'; $Id = $Id.Substring(5); Break}
        default {$SearchBy = 'Name'} }

    if ($SearchBy) {
        Write-Host "Searching by $SearchBy = $ID"
        $Groups = @(Get-ADGroup -Filter {$SearchBy -like $Id} -Properties *)
    }
    $Groups = $Groups | select Name, DisplayName, SamAccountName, Description, GroupCategory, GroupScope, CanonicalName, `
        DistinguishedName, ManagedBy, MemberOf, proxyAddresses, whenCreated, whenChanged
        
    if ($ObjectOutput) {
        $Groups
    }
    else {
        $OutputTemplate = New-Object System.Collections.Specialized.OrderedDictionary
        $OutputTemplate.Add('General', ('Name', 'DisplayName', 'SamAccountName', 'Description', 'GroupCategory', 'GroupScope', 'CanonicalName', 'DistinguishedName'))
        $OutputTemplate.Add('Extra', ('proxyAddresses', 'ManagedBy', 'MemberOf', 'whenCreated', 'whenChanged'))
        GetSectionOutput $Groups -OutputTemplate $OutputTemplate
    }
}

function Get-SDGroupMember {
    param(
        [string]$Identity,
        [switch]$Recursive = $false,
        [switch]$DetailedInfo = $false)
        if ($Recursive) {
            $Members = Get-ADGroupMember $Identity -Recursive}
            else {$Members = Get-ADGroupMember $Identity}

        $SelectStatement = $SelectStatement = @('ObjectClass','Name','DisplayName','sAMAccountName','userPrincipalName')
        if ($DetailedInfo) {
            $SelectStatement += 'mail','mailNickname',@{n='whenCreated';e={Get-Date $_.whenCreated -UFormat $TimeFormat}},`
            @{n='whenChanged';e={Get-Date $_.whenChanged -UFormat $TimeFormat}}}
        <#
        foreach ($Member in $Members) {
            switch ($Member.objectClass) {
                'user' {$MemberDetailed = Get-ADUser $Member.objectGUID | select Name, DisplayName, SamAccountName; Break}
                default {Write-Host 'unknown'}
            }
        } #>

        $Members | Get-ADObject -Properties * | select $SelectStatement 

        #@{n='MemberCountRec';e={@(Get-ADGroupMember $_.DistinguishedName -Recursive).Count}}
        # If member count = 0 or 1, it would return nothing. Have to wrap expression in @()
        # to make it an array and report count properly
}

function Get-SDEvent ($Sam = '', $ComputerName = $PreciseServer, $ID = 6273, [int]$HoursAgo) {
   $HashTable = @{LogName="Security"; ID=$ID}
   if ($HoursAgo) {$HashTable.Add('StartTime',(Get-Date).AddHours(-$HoursAgo))}
   if ($ComputerName -eq 'All') {
        $ComputerName = $DCs}
    if ($ID -eq "*") {
        $HashTable.Remove('ID')
    }

    $ComputerName | %{Write-Host "Searching in $_ :"; Get-WinEvent -ComputerName $_ -FilterHashtable $HashTable |
        where {$_.message -like "*$Sam*"}} | fl
}

function Get-SDLockoutsLegacy ($Sam = '', $ComputerName = $PreciseServer, [int]$HoursAgo) {
    $HashTable = @{LogName="Security"; ID=4740}
    if ($HoursAgo) {$HashTable.Add('StartTime',(Get-Date).AddHours(-$HoursAgo))}
    if ($ComputerName -eq 'All') {
        $ComputerName = $DCs}

    $ComputerName | %{Write-Host "Searching in $_ :"; Get-WinEvent -ComputerName $_ -FilterHashtable $HashTable -ErrorAction Silent |
        where {$_.message -like "*$Sam*"} |
        %{Write-Host ("4740 - {0} - {1} - {2}" -f (
            Get-Date ([xml]$_.ToXml()).event.system.timecreated.systemtime -UFormat $TimeFormat),
            ([xml]$_.ToXml()).event.eventdata.data[0].'#text',
            ([xml]$_.ToXml()).event.eventdata.data[1].'#text'
            )}}
}

function Get-SDLockouts ($Sam = '', $DomainController = $PreciseServer, [int]$HoursAgo, [switch]$ResolveUPN = $False) {
    $HashTable = @{LogName="Security"; ID = 4740}
    if ($HoursAgo) {$HashTable.Add('StartTime',(Get-Date).AddHours(-$HoursAgo))}
    if ($DomainController -eq 'All') {
        $DomainController = $DCs}
    $SelectStatement = @(
        @{n='DC';e={$DC}},
        @{n='Time';e={Get-Date ([xml]$_.ToXml()).event.system.timecreated.systemtime -UFormat $TimeFormat}},
        @{n='SAM';e={([xml]$_.ToXml()).event.eventdata.data[0].'#text'}},
        @{n='Caller';e={([xml]$_.ToXml()).event.eventdata.data[1].'#text'}}
    )

    if ($ResolveUPN) { # Get UPN if switch = True
        $SelectStatement += @{n='UPN';e={(Get-ADUser ([xml]$_.ToXml()).event.eventdata.data[0].'#text').UserPrincipalName}}
    }

    foreach ($DC in $DomainController) {
        Write-Host "Searching in $DC :"
        Get-WinEvent -ComputerName $DC -FilterHashtable $HashTable -ErrorAction Silent | ?{$_.message -like "*$Sam*"} |
            select $SelectStatement
    }
}

function Get-SDUserByPhone ($Number) {
    Get-ADUser -Filter {mobilephone -ne '$null'} -Properties mobilephone |
        where-object {($_.mobilephone).replace(' ','') -like $Number}
    Get-ADUser -Filter {telephoneNumber -ne '$null'} -Properties telephoneNumber | where-object {($_.telephoneNumber).replace(' ','') -like $Number}
}

function Get-SDUserSGDiff ([array]$Users, [switch]$DNFormat=$false) {
    If ($Users.Length -ge 2) {
        $UserInfo = @()
        foreach ($User in $Users) {
            $UserInfo += (Get-ADUser $User -Properties memberof | select SamAccountName, UserPrincipalName, memberof)
        }
        
        $AllSGs = @()
        foreach ($User in $UserInfo) {
            foreach ($UserSG in $User.memberof) {
                $AllSGs += $UserSG
            }
        }
        Write-Host "`n=== Same membership ==="
        foreach ($UserSG in $UserInfo[0].memberof | sort) {
            [regex]$regex = $userSG
            If ($regex.Matches($AllSGs).count -eq $Users.Length) {
                if ($DNFormat -eq $false) {
                Write-Host ($UserSG -split "CN=|,OU=|,CN=")[1]
                } Else {Write-Host "$UserSG"}
            }
        }
        Write-Host "`n=== Unique membership ==="
        foreach ($User in $UserInfo) {
            Write-Host "/////" $User.SamAccountName "/" $user.UserPrincipalName
            foreach ($UserSG in $User.memberof | sort) {
                [regex]$regex = $UserSG
                If ($regex.Matches($AllSGs).count -eq 1) {
                    if ($DNFormat -eq $false) {
                        Write-Host "   " ($UserSG -split "CN=|,OU=|,CN=")[1]
                    } Else {Write-Host "   $UserSG"}
                }
            }
        }
        If ($Users.Length -ge 3) {
            $AllSGsGrouped = $AllSGs | group | ?{$_.Count -eq $Users.Length -1}
            Write-Host "`n=== Other ", ($Users.Length - 1), " users have these groups ==="
            foreach ($User in $UserInfo) {
                Write-Host "/////" $User.SamAccountName "/" $user.UserPrincipalName
                foreach ($GroupedSG in $AllSGsGrouped.Name | sort) {
                    [regex]$regex = $GroupedSG
                    If (!$regex.Matches($User.memberof).count -eq 1) {
                        if ($DNFormat -eq $false) {
                            Write-Host "   " ($GroupedSG -split "CN=|,OU=|,CN=")[1]
                        } Else {Write-Host "   $GroupedSG"}
                    }
                }
            }
            
        }

    } Else {Write-Host "<2 users were provided"}
}

function Get-SDLockoutStatus ([string]$Id, [switch]$AllDCs=$False, [switch]$Detailed=$False, [switch]$NormalOutput=$False) {
    $DCs = Get-ADDomainController -Filter *
    $Output = @()
    foreach ($DC in $DCs | sort name) {
        $UserInfo = Get-ADUser $Id -Server $DC.Hostname -Properties LastBadPasswordAttempt, PasswordLastSet, badPwdCount, 
                                                            ` logonCount, lastLogon, LockedOut, PasswordExpired
        $pwAge = (New-TimeSpan -Start $UserInfo.PasswordLastSet)
        $pwAge = "{0} days {1} hrs {2} min" -f ($pwAge.days, $pwAge.Hours, $pwAge.Minutes)
        
        If ($UserInfo.PasswordLastSet) {
            $PasswordLastSet = (Get-Date $UserInfo.PasswordLastSet -UFormat $TimeFormat) }
        else {$PasswordLastSet = 'None'}

        If ($UserInfo.LastBadPasswordAttempt) {
            $LastBadPasswordAttempt = (Get-Date $UserInfo.LastBadPasswordAttempt -UFormat $TimeFormat)}
        else {if (!$AllDCs) {Continue} # If LastBadPasswordAttempt is empty + AllDCs switch !True skip below actions
            else {$LastBadPasswordAttempt = 'None'}}
        
        $lastLogon = Get-Date ([datetime]::FromFileTime($UserInfo.lastLogon)) -UFormat $TimeFormat

        $Result = [ordered] @{
            DC = $DC.Name
            DCSite = $DC.Site
            Locked = $UserInfo.LockedOut
            pwBadCnt = $UserInfo.badPwdCount
            pwLastBad = $LastBadPasswordAttempt
            pwSet = $PasswordLastSet
            pwAge = $pwAge
            pwExpired = $UserInfo.PasswordExpired }
        if ($Detailed) {
            $Result.Add("lastLogon", $lastLogon)
            $Result.Add("logonCount", $UserInfo.logonCount) }

        $Output += New-Object -TypeName psobject -Property $Result
    }
    if ($NormalOutput -or $Detailed) {
    $Output }
    else {$Output | ft -AutoSize}
}

function Get-SDComputerLive ([string]$ComputerName='localhost', [switch]$ObjectOutput=$false) {
    if (Test-Connection $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {$IsReachable = $True } else {$IsReachable = $false}
    if ($IsReachable) {
        $OutputTemplate = New-Object System.Collections.Specialized.OrderedDictionary
        $OutputTemplate.Add('OS', ('Hostname', 'LastBootUp', 'LastBootUpRelative', 'LocalDateTime','OSName', 'Version', 'User','InstallDate', 'Languages', 'PhysicalMemoryUsage', 'PageFileSize', 'LogicalDisks', 'Shares'))
        $OutputTemplate.Add('System',('Manufacturer', 'Model', 'SystemFamily', 'SerialNumber', 'ChassisTypes'))
        $OutputTemplate.Add('HW', ('CPU', 'PhysicalDisks', 'DisplayAdapters', 'RAM', 'NetAdapter'))

        $WMIOS = Get-WmiObject Win32_OperatingSystem -ComputerName $ComputerName -ErrorAction Stop
        $WMIComputerSystem = Get-WmiObject Win32_ComputerSystem -ComputerName $ComputerName
        $WMISystemEnclosure = Get-WmiObject Win32_SystemEnclosure -ComputerName $ComputerName
        $WMILogicalDisk = Get-WmiObject win32_LogicalDisk -ComputerName $ComputerName
        $WMIPhysicalDisk =  Get-WmiObject Win32_DiskDrive -ComputerName $ComputerName
        $WMIDisplayAdapters = Get-WmiObject Win32_VideoController -ComputerName $ComputerName
        $WMIPhysicalMemory = Get-WmiObject Win32_PhysicalMemory -ComputerName $ComputerName
        $WMICPU = Get-WmiObject Win32_Processor -ComputerName $ComputerName
        $WMINetworkAdapter = Get-WmiObject Win32_NetworkAdapter -ComputerName $ComputerName | ?{$_.NetConnectionID}
        $WMIShare = Get-WmiObject Win32_Share -ComputerName $ComputerName

        $LastBootUpRelative = ($WMIOS.ConvertToDateTime($WMIOS.LocalDateTime) - $WMIOS.ConvertToDateTime($WMIOS.LastBootUpTime))
        
        $FreePhysicalMemory = [math]::Round($WMIOS.FreePhysicalMemory / 1MB,2)
        $TotalPhysicalMemory = [math]::Round($WMIOS.TotalVisibleMemorySize / 1MB,2)
        $UsedPhysicalMemory = $TotalPhysicalMemory - $FreePhysicalMemory

        
        $LogicalDisks = @()
        foreach ($Disk in $WMILogicalDisk) {
            if ($Disk.Size) {
                $LogicalDiskSize = [math]::Round($Disk.Size / 1GB,2)
                $LogicalDiskFree = [math]::Round($Disk.FreeSpace / 1GB,2)
                $LogicalDiskUsed = $LogicalDiskSize - $LogicalDiskFree
                $LogicalDiskPercentageFree = [math]::Round(($LogicalDiskFree / $LogicalDiskSize*100),2)
                $Disk = "{0} {1} | {2} GB / {3} GB used | {4} GB, {5} % free" -f $Disk.Caption, $Disk.Description, $LogicalDiskUsed, $LogicalDiskSize, $LogicalDiskFree, $LogicalDiskPercentageFree
                $LogicalDisks += $Disk }
             else {
                $Disk = "{0} {1} | no size" -f $Disk.Caption, $Disk.Description
                $LogicalDisks += $Disk } }

        $Result = [ordered]@{
            # OS
            Hostname = $WMIOS.PSComputerName
            LastBootUp = Get-Date $WMIOS.ConvertToDateTime($WMIOS.LastBootUpTime) -UFormat $TimeFormat
            LastBootUpRelative = "{0} days {1} hr {2} min {3} sec" -f $LastBootUpRelative.Days, $LastBootUpRelative.Hours, $LastBootUpRelative.Minutes, $LastBootUpRelative.Seconds
            LocalDateTime = Get-Date $WMIOS.ConvertToDateTime($WMIOS.LocalDateTime) -UFormat $TimeFormat
            OSName = "{0} - {1}" -f $WMIOS.Caption, $WMIOS.OSArchitecture
            Version = $WMIOS.Version
            User = $WMIComputerSystem.UserName
            InstallDate = Get-Date $WMIOS.ConvertToDateTime($WMIOS.InstallDate) -UFormat $TimeFormat           
            Languages = $WMIOS.MUILanguages # Available language packs
            Architecture = $WMIOS.OSArchitecture
            PhysicalMemoryUsage = "{0} GB / {1} GB | {2} GB free" -f $UsedPhysicalMemory, $TotalPhysicalMemory, $FreePhysicalMemory
            UsedPhysicalMemory = $UsedPhysicalMemory
            TotalPhysicalMemory = $TotalPhysicalMemory
            FreePhysicalMemory = $FreePhysicalMemory
            PageFileSize = "{0} MB" -f [math]::Round($WMIOS.SizeStoredInPagingFiles / 1KB,2)
            LogicalDisks = $LogicalDisks
            Shares = $WMIShare | %{"{0} > {1} | {2}" -f $_.Name, $_.Path, $_.Description}
            
            # System
            Manufacturer = $WMIComputerSystem.Manufacturer
            Model = $WMIComputerSystem.Model
            SystemFamily = $WMIComputerSystem.SystemFamily
            SerialNumber = $WMISystemEnclosure.SerialNumber
            ChassisTypes = $WMISystemEnclosure.ChassisTypes

            # HW
            CPU = $WMICPU | %{"{0} - {1} ({2} MHz) | {3} Cores | {4} Logical" -f $_.DeviceID, $_.Name, $_.CurrentClockSpeed,
                ` $_.NumberOfCores, $_.NumberOfLogicalProcessors}

            PhysicalDisks = $WMIPhysicalDisk | %{"{1} (s/n: {2}) | {0} | {3} GB" -f $_.MediaType, $_.Model,
                ` $_.SerialNumber, [math]::Round($_.Size / 1GB,2)}
            
            DisplayAdapters = $WMIDisplayAdapters | %{"{0} - {1} | {2}" -f $_.Name, $_.AdapterDACType, $_.VideoModeDescription}
            
            RAM = $WMIPhysicalMemory | %{"{0} MB - {1} / {2} MHz | {3} / {4}" -f ($_.Capacity /1MB), $_.DeviceLocator, $_.Speed,
                ` $_.Manufacturer, $_.PartNumber}
            
            NetAdapter = $WMINetworkAdapter | %{"{0} - {1} | Enabled:{2} | {3}" -f $_.NetConnectionID, $_.Description,
                ` $_.NetEnabled, $_.MACAddress}
        }
        if ($ObjectOutput) {
            New-Object -TypeName psobject -Property $Result
        } else { GetSectionOutput $Result -OutputTemplate $OutputTemplate }
    } else {return "$ComputerName can't be reached"}
}

function Get-SDInstalledSoftware ([string]$ComputerName='localhost',[switch]$NormalOutput=$false ) {
    if (Test-Connection $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {$IsReachable = $True } else {$IsReachable = $false}
    if ($IsReachable) {
        $Output = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
            "", "\Wow6432Node" | %{Get-ItemProperty "HKLM:\Software$_\Microsoft\Windows\CurrentVersion\Uninstall\*"} }
        $Output = $Output | ?{$_.DisplayName} | select DisplayName, DisplayVersion, Publisher, InstallDate, InstallLocation | sort displayname 
        if (!$NormalOutput) {
            $Output | ft
        } else {$Output}
    } else {Write-Host "$ComputerName can't be reached"}
}

function Get-SDDrivers ([string]$ComputerName='localhost', [switch]$IncludeSystem = $false, [switch]$IncludeAllProperties = $false) {
    if (Test-Connection $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue) {$IsReachable = $True } else {$IsReachable = $false}
    if ($IsReachable) {
        $Output = Get-WmiObject Win32_PnPSignedDriver -ComputerName $ComputerName | ?{$_.DeviceClass -ne $None}
        if (!$IncludeSystem) {
            $Output = $Output | ?{$_.DeviceClass -ne "SYSTEM"}
        }
        if (!$IncludeAllProperties) {
            $Output = $Output | select DeviceClass, DeviceName, FriendlyName, Manufacturer,
            ` @{n="DriverDate"; e={Get-date $_.ConvertToDateTime($_.DriverDate) -UFormat "%Y/%m/%d"}}, DriverVersion | sort DeviceClass
        }
        $Output
    } else {Write-Host "$ComputerName can't be reached"}
}

#Export-ModuleMember *-* #Export only Noun-Verb CMDlets