<#
.SYNOPSIS
Analyze all Active Directory users by frequency of assigned properties. 
Export Analysis results to HTML report.
Find out how many AD users are assigned a specific property like "AdminCount" or "PasswordNotRequired".
Find out how many unique property assignments exist for a specific property.
.DESCRIPTION
Analyze all Active Directory users by frequency of assigned properties. 
Export Analysis results to HTML report.
.PARAMETER Server
Specifies the Active Directory server domain to query.
.PARAMETER Enabled
Specifies to only return results from enabled users.
.EXAMPLE
PS C:\> Export-ADUserPropertyFrequencies_To-Html.ps1 -Server CyberCondor.local -Enabled $true

## Input
Template.html
- Template for HTML layout 
styles.css
- Styles to display HTML layout properly
Users
- Contains list of all users and their properties
## Output
~\_FrequenyAnalysisReports\ADUserPropertiesFrequenyAnalysisReport_$Date
- Users.html
- styles.css
#>
param(
    [Parameter(mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [system.String]$Server,

    [Parameter(mandatory=$True, Position=0, ValueFromPipeline=$false)]
    [System.Boolean]$Enabled
)
Write-Host "`n`t`tAttempting to query Active Directory.'n" -BackgroundColor Black -ForegroundColor Yellow
try{Get-ADUser -server $Server -filter 'Title -like "*Admin*"' > $null -ErrorAction stop
}
catch{$errMsg = $_.Exception.message
    if($errMsg.Contains("is not recognized as the name of a cmdlet")){
        Write-Warning "`t $_.Exception"
        Write-Output "Ensure 'RSAT Active Directory DS-LDS Tools' are installed through 'Windows Features' & ActiveDirectory PS Module is installed"
    }
    elseif($errMsg.Contains("Unable to contact the server")){
        Write-Warning "`t $_.Exception"
        Write-Output "Check server name and that server is reachable, then try again."
    }
    else{Write-Warning "`t $_.Exception"}
    break
}

function Get-ExistingUsers_AD{
    try{$ExistingUsers = Get-ADUser -Server $Server -Filter * -Properties * -ErrorAction Stop
        return $ExistingUsers
    }
    catch{$errMsg = $_.Exception.message
        Write-Warning "`t $_.Exception"
        return $null
    }
}
function Get-UserRunningThisProgram($ExistingUsers_AD){
    foreach($ExistingUser in $ExistingUsers_AD){
        if($ExistingUser.SamAccountName -eq $env:UserName){return $ExistingUser}
    }
    Write-Warning "User Running this program not found."
    return $null
}
function SanitizeManagerPropertyFormat($ExistingUsers_AD){
    foreach($ExistingUser in $ExistingUsers_AD){
        [string]$UnsanitizedName = $ExistingUser.Manager
        $NameSanitized = $false
        if(($UnsanitizedName -ne $null) -and ($UnsanitizedName -ne "") -and ($UnsanitizedName -ne "`n") -and ($UnsanitizedName -match '[a-zA-Z]') -and ($UnsanitizedName.Length -ne 1)){
            $index = 0
            while($NameSanitized -eq $false){
                $SanitizedName = $ExistingUser.Manager.Substring(3,$index++)
                if($ExistingUser.Manager[$index] -eq ','){
                    $ExistingUser.Manager = $SanitizedName.Substring(0,$SanitizedName.Length - 2)
                    $NameSanitized = $true
                }
            }
        }
        else{$ExistingUser.Manager = "NULL"}
    }
}

function Create-NewHtmlReportFromTemplate($TemplateFilePath, $NameOfPage, $ReportFolder, $divInfo, $NavLinkNames, $InfoTitles){
    if(!(Test-Path $TemplateFilePath)){return $null}
    if(!(Test-Path $ReportFolder.Path)){return $null}
    $NewHtmlReport = New-Object -TypeName PSObject -Property @{Name="$($NameOfPage).html";Path="$($ReportFolder.Path)\$($NameOfPage).html"}
    $InfoDivTitles = $(foreach($i in $InfoTitles){("<a href=#$($i)> $($i) </a>")})
    $H2AndDivInfo  = "            <h2 class='glow'>
                $($NameofPage)
            </h2>
            <p class='div-info'>
                $(foreach($i in $divInfo){($i + '<br>')})
            </p>"

    Get-Content $TemplateFilePath | Select -First ((Select-String "<title>" $TemplateFilePath | select -First 1).linenumber - 1) | Out-File $NewHtmlReport.Path -Encoding utf8
    Write-Output "        <title>$($NameofPage)</title>" | Out-File $NewHtmlReport.Path -Append -Encoding utf8

    Get-Content $TemplateFilePath | Select -First ((Select-String "<h1>" $TemplateFilePath | select -First 1).linenumber - 1) | 
        Select -Last (((Select-String "<h1>" $TemplateFilePath | select -First 1).linenumber - 1) - ((Select-String "<title>" $TemplateFilePath | select -First 1).linenumber)) | 
        Out-File $NewHtmlReport.Path -Append -Encoding utf8
    Write-Output "                        <h1>$($ReportFolder.Name)</h1>"                   | Out-File $NewHtmlReport.Path -Append -Encoding utf8
    Write-Output "                        <p class='info-div-titles'>$($InfoDivTitles)</p>" | Out-File $NewHtmlReport.Path -Append -Encoding utf8

    Get-Content $TemplateFilePath | Select -First ((Select-String "</ul>" $TemplateFilePath | select -First 1).linenumber - 1) | 
        Select -Last (((Select-String "</ul>" $TemplateFilePath | select -First 1).linenumber) - ((Select-String "</p>" $TemplateFilePath | select -First 1).linenumber) - 1) | 
        Out-File $NewHtmlReport.Path -Append -Encoding utf8
    foreach($NavLink in $NavLinkNames){
        if(!(Get-Content $TemplateFilePath | Select-String "href='$($NavLink).html'")){
            Write-Output "                            <li class='nav-link'><a href='$($NavLink).html' alt='$($NavLink)'>$($NavLink)</a></li>" | 
                Out-File $NewHtmlReport.Path -Append -Encoding utf8
        }
    }
    
    Get-Content $TemplateFilePath | Select -First ((Select-String "<h2 " $TemplateFilePath | select -First 1).linenumber - 1) | 
        Select -Last (((Select-String "<h2 " $TemplateFilePath | select -First 1).linenumber) - ((Select-String "</ul>" $TemplateFilePath | select -First 1).linenumber)) | 
        Out-File $NewHtmlReport.Path -Append -Encoding utf8
    Write-Output "$($H2AndDivInfo)" | Out-File $NewHtmlReport.Path -Append -Encoding utf8
    
    return $NewHtmlReport
}
function New-InfoDiv($HtmlReportPath, $Title, $Info){
    if(!(Test-Path $HtmlReportPath)){return $null}
    $Output = "            <div class='info-div'>
                <h3 id='$($Title)'>$($Title)</h3>
                <p class='div-info'>
                    $(foreach($i in $Info){("<a href=#$($i)> $($i) </a><br>")})
                </p>"
    write-Output $Output | Out-File $HtmlReportPath -Append -Encoding utf8
}
function ExportTo-HTML_NewTable-Div($Title, $Info, $Object){
    $htmlParams = @{
      PreContent = "                <div class='table-div'>
                    <h4 id='$($Title)'>$($Title)</h4>
                    <p class='table-info'>
                        Total: $($Info)
                    </p>"
      PostContent = "                </div>"
    }
   return $Object | ConvertTo-Html -As Table -Fragment @htmlParams
}
function End-InfoDiv($HtmlReportPath){
    if(!(Test-Path $HtmlReportPath)){return $null}
    Write-Output "            </div>" | Out-File $HtmlReportPath -Append -Encoding utf8
}
function Append-HtmlReport_Footer($HtmlReportPath){
    if(!(Test-Path $HtmlReportPath)){return $null}
    $Footer = "            <footer class='footnav'>
                Report Produced by $($env:USERNAME) on hostname: $(hostname)
                $(Get-Date -Format yyy-MM-dd)
                CONFIDENTIAL - NOT FOR DISTRIBUTION
            </footer>
        </div>
    </body>
</html>"
    $Footer  | Out-File $HtmlReportPath -Append -Encoding utf8
}

function New-PropertyFrequenciesHTML_Report-FromTemplate($TemplateFilePath, $ReportFolder, $MainObject, $MainObjectNames, $MainObjectName){
    $PropertyObjects = $MainObject | Select -First 1 | Get-Member | where{($_.MemberType -eq "Property") -and ($_.Definition -notlike "*ValueCollection*") -and ($_.Definition -notlike "*list*")} 
    $Properties = New-Object -TypeName PSObject -Property @{Int=@();Boolean=@();String=@();Datetime=@()}
    $Properties | Add-Member -NotePropertyMembers @{Name=$($Properties | Get-Member | where{($_.MemberType -eq "NoteProperty")} | select -ExpandProperty Name)}

    foreach($Property in $PropertyObjects){
        foreach($PropertyName in $Properties.Name){
            if($Property.Definition -like "*$($PropertyName)*"){
                $Properties.$($PropertyName) += New-Object -TypeName PSObject -Property @{Name=$($Property | select -ExpandProperty Name);
                                                                                          Frequency=@($(Get-PropertyFrequencies $($Property | select -ExpandProperty Name) $MainObject))}
            }
        }
    }

    $HTML_Report = Create-NewHtmlReportFromTemplate $TemplateFilePath $MainObjectName $ReportFolder $(foreach($PropertyName in $Properties.Name){write-output "$(($Properties.$PropertyName).count) $($PropertyName) Frequencies`n---"}) $MainObjectNames $($Properties.Name) 
    foreach($PropertyName in $Properties.Name){
        New-InfoDiv $HTML_Report.Path "$($PropertyName)" $Properties.$PropertyName.Name
        foreach($Property in $Properties.$PropertyName | sort Frequency){
            ExportTo-HTML_NewTable-Div "$($Property.Name)" $(($Property.Frequency).count) $($Property.Frequency) | Out-File $HTML_Report.Path -Append -Encoding utf8
        }
        End-InfoDiv              $HTML_Report.Path
    }
    Append-HtmlReport_Footer $HTML_Report.Path

    return $HTML_Report
}

function Get-PropertyFrequencies($Property, $Object){
    $Total = ($Object).count
    $ProgressCount = 0
    $AllUniquePropertyValues = $Object | select $Property | sort $Property | unique -AsString # Get All Uniques
    $PropertyFrequencies = @()                                                                # Init empty Object
    $isDate = $false                                                                                                                                                          
    foreach($UniqueValue in $AllUniquePropertyValues){
        if(!($isDate -eq $true)){
            if([string]$UniqueValue.$Property -as [DateTime]){$isDate = $true}
        }
        $PropertyFrequencies += New-Object -TypeName PSobject -Property @{$Property=$($UniqueValue.$Property);Count=0;Frequency="100%"} # Copy Uniques to Object Array and Init Count as 0
    }
    if(($isDate -eq $true) -and (($Object | Select $Property | Get-Member).Definition -like "*datetime*")){
        foreach($PropertyFrequency in $PropertyFrequencies){
            if(($PropertyFrequency.$Property) -and ([string]$PropertyFrequency.$Property -as [DateTime])){
                try{$PropertyFrequency.$Property = $PropertyFrequency.$Property.ToString("yyyy-MM")}
                catch{# Nothing
                }
            }
        }
        foreach($PropertyName in $Object.$Property){                                                            # For each value in Object
            if($Total -gt 0){Write-Progress -id 1 -Activity "Finding $Property Frequencies -> ( $([int]$ProgressCount) / $Total )" -Status "$(($ProgressCount++/$Total).ToString("P")) Complete"}
            foreach($PropertyFrequency in $PropertyFrequencies){                                                # Search through all existing Property values
                if(($PropertyName -eq $null) -and ($PropertyFrequency -eq $null)){$PropertyFrequency.Count++}   # If Property value is NULL, then add to count - still want to track this
                elseif($PropertyName -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}             # Else If Property value is current value, then add to count
                else{
                    try{if($PropertyName.ToString("yyyy-MM") -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}}
                    catch{# Nothing
                    }
                }
            }
        }
    }
    else{
        foreach($PropertyName in $Object.$Property){                                                            # For each value in Object
            if($Total -gt 0){Write-Progress -id 1 -Activity "Finding $Property Frequencies -> ( $([int]$ProgressCount) / $Total )" -Status "$(($ProgressCount++/$Total).ToString("P")) Complete"}
            foreach($PropertyFrequency in $PropertyFrequencies){                                                # Search through all existing Property values
                if(($PropertyName -eq $null) -and ($PropertyFrequency -eq $null)){$PropertyFrequency.Count++}   # If Property value is NULL, then add to count - still want to track this
                elseif($PropertyName -ceq $PropertyFrequency.$Property){$PropertyFrequency.Count++}             # Else If Property value is current value, then add to count
            }
        }
    }
    Write-Progress -id 1 -Completed -Activity "Complete"
    if($Total -gt 0){
        foreach($PropertyFrequency in $PropertyFrequencies){$PropertyFrequency.Frequency = ($PropertyFrequency.Count/$Total).ToString("P")}
    }
    return $PropertyFrequencies | select Count,$Property,Frequency | sort Count,$Property | Unique -AsString
}

function main{
    Write-Output "Starting Query on all properties for AD Users.`nThis could take some time.`nStart Query: $(date)"
    if($Enabled -eq $true){$ExistingUsers_AD = Get-ExistingUsers_AD | where{$_.Enabled -eq $true}}
    else{                  $ExistingUsers_AD = Get-ExistingUsers_AD}
    if($ExistingUsers_AD -eq $null){break}
    Write-Output "End Query: $(date)"
    $UserRunningThisProgram = Get-UserRunningThisProgram $ExistingUsers_AD
    if($UserRunningThisProgram -ne $null){
        Write-Host "`n`t`tHello '$($UserRunningThisProgram.Name) ($($UserRunningThisProgram.Title))'!" -ForegroundColor Green
    }
    Write-Host "This program will export all Active Directory USER property frequencies from 'AD Users' to an HTML report.`n"
   
    SanitizeManagerPropertyFormat $ExistingUsers_AD

    $MainFolderPath   = "~\_FrequenyAnalysisReports"
    if($Enabled -eq $true){$ReportFolderName = ("ADUserPropertiesFrequenyAnalysisReport-$($Server)_EnabledUsersOnly_" + $(get-date -format yyy-MM-dd))}
    else{                  $ReportFolderName = ("ADUserPropertiesFrequenyAnalysisReport-$($Server)_" + $(get-date -format yyy-MM-dd))}
    $ReportFolder     = New-Object -TypeName PSObject -Property @{Name=$ReportFolderName;Path="$MainFolderPath\$ReportFolderName"}
    $CurrDir          = (pwd).Path
    if((!(Test-Path Template.html)) -or (!((Get-Content Template.html | select -First 1).Contains("<!DOCTYPE HTML>")))){Write-Output "Cannot Find HTML Template File." ; break}
    if(!(Test-Path $MainFolderPath)){mkdir $MainFolderPath}
    if(!(Test-Path $ReportFolder.Path)){
        mkdir $ReportFolder.Path
        try{Copy "$CurrDir/styles.css" $ReportFolder.Path}
        catch{$errMsg = $_.Exception.message
            Write-Warning "`t $_.Exception"
            break
        }
    }
    else{Write-Warning "Folder '$($ReportFolder.Path)' already exists.`nProgram already ran today." ; break}

    Write-Output "`n$($env:UserName) - Started this program @ $(date)"
    Write-Host "$($env:UserName) - Started this program @ $(date)"
    ################################################################

    $MainObjectNames = @("Users")

    New-PropertyFrequenciesHTML_Report-FromTemplate "Template.html" $ReportFolder $ExistingUsers_AD $MainObjectNames $($MainObjectNames | select -First 1 | select -Last 1) 

    ################################################################
    Write-Output "Your Report is ready and can be found @: "
    $ReportFolder.Path

    Write-Output "`nProgram completed @ $(date)"
    Write-Host "Program completed @ $(date)"
}

main
