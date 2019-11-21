function Compare-ObjectProperties {
    Param (
        [PSObject]$ReferenceObject,
        [PSObject]$DifferenceObject,
        [Switch]$IncludeEqual )
    $ReferenceProperties = ($ReferenceObject | Get-Member -MemberType Properties).Name
    $DifferenceProperties = ($DifferenceObject | Get-Member -MemberType Properties).Name
    $Properties = $ReferenceProperties + $DifferenceProperties | select -Unique
    #$Properties
    foreach ($Property in $Properties) {
        $CompareResult = $ReferenceObject.$Property -eq $DifferenceObject.$Property
        $Result = [ordered] @{
            Property = $Property
            Equal = $CompareResult
            ReferenceValue = $ReferenceObject.$Property
            DifferenceValue = $DifferenceObject.$Property
            ReferenceHas = $Property -in $ReferenceProperties
            DifferenceHas = $Property -in $DifferenceProperties }
        if ($IncludeEqual) {
            New-Object -TypeName psobject -Property $Result }
            else {
                New-Object -TypeName psobject -Property $Result | where Equal -eq $False}
    }
}

function Get-RandomCommand {
    Get-Random -InputObject (Get-Command -Module Microsoft*,Cim*,PS*,ISE) | man
}

function Get-RandomTopic {
    Get-Random -InputObject (Get-Help -Category HelpFile) | man
}

function Test-Host ([array]$HostList, [int]$Count=1) {
    $HostList | select @{n='Host';e={$_}}, @{n='IsUp';e={Test-Connection $_ -Count $Count -Quiet}}
}

function Get-RadiusLog {
    param (
        [string]$LogContent,
        [string]$LogPath)
    $Header = "ComputerName","ServiceName","Record-Date","Record-Time","Packet-Type","User-Name",
        "Fully-Qualified-Distinguished-Name","Called-Station-ID","Calling-Station-ID","Callback-Number",
        "Framed-IP-Address","NAS-Identifier","NAS-IP-Address","NAS-Port","Client-Vendor","Client-IP-Address",
        "Client-Friendly-Name","Event-Timestamp","Port-Limit","NAS-Port-Type","Connect-Info","Framed-Protocol",
        "Service-Type","Authentication-Type","Policy-Name","Reason-Code","Class","Session-Timeout","Idle-Timeout",
        "Termination-Action","EAP-Friendly-Name","Acct-Status-Type","Acct-Delay-Time","Acct-Input-Octets","Acct-Output-Octets",
        "Acct-Session-Id","Acct-Authentic","Acct-Session-Time","Acct-Input-Packets","Acct-Output-Packets","Acct-Terminate-Cause",
        "Acct-Multi-Ssn-ID","Acct-Link-Count","Acct-Interim-Interval","Tunnel-Type","Tunnel-Medium-Type","Tunnel-Client-Endpt",
        "Tunnel-Server-Endpt","Acct-Tunnel-Conn","Tunnel-Pvt-Group-ID","Tunnel-Assignment-ID","Tunnel-Preference","MS-Acct-Auth-Type",
        "MS-Acct-EAP-Type","MS-RAS-Version","MS-RAS-Vendor","MS-CHAP-Error","MS-CHAP-Domain","MS-MPPE-Encryption-Types",
        "MS-MPPE-Encryption-Policy","Proxy-Policy-Name","Provider-Type","Provider-Name","Remote-Server-Address",
        "MS-RAS-Client-Name","MS-RAS-Client-Version"
    if ($LogPath) {Import-Csv $LogPath -Header $Header}
        elseif ($LogContent) {ConvertFrom-Csv $LogContent -Header $Header}
}

function Connect-O365 {
    # Connect to o365 Exchange
    param(
        [switch]$UseProxymethodRPS = $false,
        [switch]$UsePrefix = $false)
    $UserName = 'your.username@domain.com'
    $O365Credential = Get-Credential $UserName -Message "o365 account"
    if ($UseProxymethodRPS) {
        $ConnectionUri = 'https://outlook.office365.com/powershell-liveid/?proxymethod=rps'
    } else {$ConnectionUri = 'https://outlook.office365.com/powershell-liveid/'}
    $Session = New-PSSession -ConfigurationName Microsoft.Exchange `
        -ConnectionUri $ConnectionUri `
        -Credential $O365Credential -Authentication Basic -AllowRedirection
    if ($UsePrefix) {
        Import-PSSession $Session -prefix "o365"
    } else {Import-PSSession $Session}
}