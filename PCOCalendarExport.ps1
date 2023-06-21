# Planning Center Online Calendar Data Export
# To assist in migration to another service (FMX) which insists on a CSV input format.
# I suspect other PCO services can be exported with a modified script which includes a different list of $PCOTypes.
# Initial API URLs were discovered at https://api.planningcenteronline.com/explorer

# Wesley Deal 2023-06-21 for Fellowship Christian School
# No rights reserved: I hereby release this software into the public domain.
# This software is supplied "as is", without warranty express or implied.

# invoke cmdlets by dot-sourcing this script first
# for example:   C:\> . .\PCOCalendarExport.ps1
#                C:\> Export-AllPCOData

# *** CONFIGURATION ***
# Set your authorization info below
# Use the Personal Access Tokens from https://api.planningcenteronline.com/oauth/applications
$PCOAppID  = "YOUR_APP_ID_HERE"
$PCOSecret = "YOUR_SECRET_HERE"
# *** END CONFIGURATION ***

$PCOBase64AuthInfo = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($PCOAppID + ":" + $PCOSecret))

# via Ronald Bode ("iRon") @ https://web.archive.org/web/20190616050616/https://powersnippets.com/flatten-object/
# originally found on https://stackoverflow.com/a/46081131
# license status for this Cmdlet is unknown
Function Flatten-Object {                                       # Version 00.02.12, by iRon
    [CmdletBinding()]Param (
        [Parameter(ValueFromPipeLine = $True)][Object[]]$Objects,
        [String]$Separator = ".", [ValidateSet("", 0, 1)]$Base = 1, [Int]$Depth = 5, [Int]$Uncut = 1,
        [String[]]$ToString = ([String], [DateTime], [TimeSpan]), [String[]]$Path = @()
    )
    $PipeLine = $Input | ForEach {$_}; If ($PipeLine) {$Objects = $PipeLine}
    If (@(Get-PSCallStack)[1].Command -eq $MyInvocation.MyCommand.Name -or @(Get-PSCallStack)[1].Command -eq "<position>") {
        $Object = @($Objects)[0]; $Iterate = New-Object System.Collections.Specialized.OrderedDictionary
        If ($ToString | Where {$Object -is $_}) {$Object = $Object.ToString()}
        ElseIf ($Depth) {$Depth--
            If ($Object.GetEnumerator.OverloadDefinitions -match "[\W]IDictionaryEnumerator[\W]") {
                $Iterate = $Object
            } ElseIf ($Object.GetEnumerator.OverloadDefinitions -match "[\W]IEnumerator[\W]") {
                $Object.GetEnumerator() | ForEach -Begin {$i = $Base} {$Iterate.($i) = $_; $i += 1}
            } Else {
                $Names = If ($Uncut) {$Uncut--} Else {$Object.PSStandardMembers.DefaultDisplayPropertySet.ReferencedPropertyNames}
                If (!$Names) {$Names = $Object.PSObject.Properties | Where {$_.IsGettable} | Select -Expand Name}
                If ($Names) {$Names | ForEach {$Iterate.$_ = $Object.$_}}
            }
        }
        If (@($Iterate.Keys).Count) {
            $Iterate.Keys | ForEach {
                Flatten-Object @(,$Iterate.$_) $Separator $Base $Depth $Uncut $ToString ($Path + $_)
            }
        }  Else {$Property.(($Path | Where {$_}) -Join $Separator) = $Object}
    } ElseIf ($Objects -ne $Null) {
        @($Objects) | ForEach -Begin {$Output = @(); $Names = @()} {
            New-Variable -Force -Option AllScope -Name Property -Value (New-Object System.Collections.Specialized.OrderedDictionary)
            Flatten-Object @(,$_) $Separator $Base $Depth $Uncut $ToString $Path
            $Output += New-Object PSObject -Property $Property
            $Names += $Output[-1].PSObject.Properties | Select -Expand Name
        }
        $Output | Select ([String[]]($Names | Select -Unique))
    }
}


Function Export-PCOData {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
          [string]$InitialAPIURL,
        [Parameter(Mandatory=$true)]
          [string]$OutputJSON
    )
    
    $next = $InitialAPIURL
    while($next){
        write-host "GET $next"
        $result = irm -Headers @{Authorization=("Basic {0}" -f $PCOBase64AuthInfo)} -uri $next
        $data += $result.data
        if ($result.links.next) {
            $next = $result.links.next
        } else {
            $next = ""
        }
        start-sleep -milliseconds 250
    }
    Write-Host Save JSON to $OutputJSON ...
    $data | convertto-json -depth 100 | out-file $OutputJSON
}

Function Export-PCOAttachments {
    [CmdletBinding()]
    Param (
        [string]$AttachmentInitialURI = "https://api.planningcenteronline.com/calendar/v2/attachments?per_page=100",
        [Parameter(Mandatory=$true)]
          [string]$OutputDir
    )
    
    $next = $AttachmentInitialURI
    while($next){
        write-host "GET $next"
        $result = irm -Headers @{Authorization=("Basic {0}" -f $PCOBase64AuthInfo)} -uri $next
        $data += $result.data
        if ($result.links.next) {
            $next = $result.links.next
        } else {
            $next = ""
        }
        start-sleep -milliseconds 250
    }
    $Data | % {
        $URL = $_.attributes.url
        $Filepath = $OutputDir + $_.attributes.name.replace("/","_").replace("\","_").replace(":","_")
        try{
            $Filepath += "." + $_.attributes.content_type.split("/")[1].replace("vnd.openxmlformats-officedocument.wordprocessingml.document","docx").replace("jpeg","jpg")
        } catch {
            write-host [WARN] Could not determine file extension
        }
        
        Write-Host Getting $Filepath ...
        irm -uri $URL -outfile $Filepath
    }
    
}

Function Convert-PCODataToCSV {
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
          [string]$InputJSONPath,
        [Parameter(Mandatory=$true)]
          [string]$OutputPath
    )
    Write-Host Flatten structured data to $OutputPath ...
    Get-Content $InputJSONPath | ConvertFrom-Json | Flatten-Object | Convertto-CSV -NoTypeInformation | out-file $OutputPath
}

Function Export-AllPCOData {
    [CmdletBinding()]
    
    $PCOTypes = [ordered]@{
        attachments              = "https://api.planningcenteronline.com/calendar/v2/attachments?per_page=100";
        conflicts                = "https://api.planningcenteronline.com/calendar/v2/conflicts?per_page=100";
        event_instances          = "https://api.planningcenteronline.com/calendar/v2/event_instances?per_page=100";
        event_resource_requests  = "https://api.planningcenteronline.com/calendar/v2/event_resource_requests?per_page=100";
        events                   = "https://api.planningcenteronline.com/calendar/v2/events?per_page=100";
        feeds                    = "https://api.planningcenteronline.com/calendar/v2/feeds?per_page=100";
        people                   = "https://api.planningcenteronline.com/calendar/v2/people?per_page=100";
        report_templates         = "https://api.planningcenteronline.com/calendar/v2/report_templates?per_page=100";
        resource_approval_groups = "https://api.planningcenteronline.com/calendar/v2/resource_approval_groups?per_page=100";
        resource_bookings        = "https://api.planningcenteronline.com/calendar/v2/resource_bookings?per_page=100";
        resource_folders         = "https://api.planningcenteronline.com/calendar/v2/resource_folders?per_page=100";
        resource_questions       = "https://api.planningcenteronline.com/calendar/v2/resource_questions?per_page=100";
        resources                = "https://api.planningcenteronline.com/calendar/v2/resources?per_page=100";
        room_setups              = "https://api.planningcenteronline.com/calendar/v2/room_setups?per_page=100";
        tag_groups               = "https://api.planningcenteronline.com/calendar/v2/tag_groups?per_page=100";
    }
    
    $Date = Get-Date -UFormat "%Y-%m-%d %Hh%Mm%Ss"
    $OutputDirJSON         = "./PCO Export " + $Date + "/JSON/"
    $OutputDirCSV          = "./PCO Export " + $Date + "/CSV/"
    $OutputDirAttachments  = "./PCO Export " + $Date + "/Attachments/"
    mkdir $OutputDirJSON
    mkdir $OutputDirCSV
    mkdir $OutputDirAttachments
    
    Write-Host *** Getting attachments...
    Export-PCOAttachments -OutputDir $OutputDirAttachments
    Write-Host *** Getting data in structured JSON format...
    $i=1
    $PCOTypes.GetEnumerator() | % {
        Write-Host Downloading $_.key [$i/$($PCOTypes.count)]
        Export-PCOData -InitialAPIURL $($_.value) -OutputJSON $($OutputDirJSON + $_.key + ".json")
        $i++
    }
    Write-Host *** Beginning JSON to CSV flattening. This process will take a while.
    $i=1
    $PCOTypes.Keys | % {
        Write-Host Converting $_ [$i/$($PCOTypes.count)]
        Convert-PCODataToCSV -InputJSONPath  $($OutputDirJSON + $_ + ".json") -OutputPath $($OutputDirCSV + $_ + ".csv")
        $i++
    }
}
