#Requires -Version 7.1
#Requires -Modules CognosModule

<#

    .SYNOPSIS
    HallPass
    Craig Millsap - 2/2022
    Data Upload for students, photos, guardians, and faculty to HallPass.
    Looking for more Automation? https://www.camtehcs.com

#>

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
Set-Location $currentPath

$hostkey = '8e:20:fc:57:99:c1:0b:9c:42:b9:bb:3d:fd:a4:89:27'

if (-Not(Test-Path $currentPath\hallpass)) { New-Item -ItemType Directory hallpass }
if (-Not(Test-Path $currentPath\files)) { New-Item -ItemType Directory files }

if (-Not(Test-Path $PSScriptRoot\settings.ps1)) {
    Write-Host "Error: Failed to find settings.ps1 file."    
    exit 1
} else {
    . $PSScriptRoot\settings.ps1
}

$filesToZip = New-Object System.Collections.ArrayList

if (-Not($authorizationFile = Get-ChildItem -Path $currentPath -Filter *.ud -Recurse -File)) {
    Write-Host "Error: Authorization File not found in any directory here. Please save the .ud file provided by HallPass to this directory."
    exit 1
} else {
    $fileName = $authorizationFile[0].baseName
    $authorizationFile = $authorizationFile[0].FullName
    $filesToZip.Add($authorizationFile) | Out-Null

    #this is -ne $false for backwards compatibility with existing configs. If not defined then use the config.ini.
    if ($eSchoolIsAuthoratative -ne $false) {
        $filesToZip.Add("files\config.ini") | Out-Null
    }
}

Save-CognosReport -report students -TeamContent -cognosfolder "_Shared Data File Reports\HallPass" -savepath "$PSScriptRoot\files"

$students = Import-CSV $PSScriptRoot\files\students.csv | Where-Object { $validbuildings -contains $PSItem.'school id' }

if ($removeHomeroomTeachers) {
    $students = $students | Select-Object -ExcludeProperty teacher | Select-Object *,teacher
}

if ($studentRFID) {
    $students | ForEach-Object {
        $PSitem.rfid = "$($PSItem.'student id')$($RFIDSuffix)"
    }
}

#There has to be valid data before we continue.
if ($students.Count -ge 1) {

    $studentIds = $students | Select-Object -ExpandProperty 'student id'

    #lets include photos if we can find them and they are the right size.
    if ($includeStudentPhotos) {
        if (Test-Path $studentPhotosFolder) {

            $studentsWithPhotoIds = @()

            $photos = Get-ChildItem -Path $studentPhotosFolder -Recurse -File -Filter "*.jpg" | Where-Object { $PSItem.Length -lt 150kb }
            $photosGrouped = $photos | Group-Object -Property BaseName
            $photosHashed = $photosGrouped | Group-Object -Property Name -AsHashTable

            $studentIds | ForEach-Object {
                if ($photosHashed."$PSitem") {
                    $studentsWithPhotoIds += $PSitem
                    $photo = $photosHashed."$PSitem"
                    
                    #Hall Pass only accepts the student photo on Fridays. So we won't add them to the zip on other days.
                    if ((Get-Date -Format dddd) -eq 'Friday') {
                        $filesToZip.Add("$($photo.Group[0].FullName)") | Out-Null
                    }
                }
            }

            #now we need to mark the students as having a photo in the output CSV.
            $students | Where-Object { $studentsWithPhotoIds -contains $PSItem.'student id' } | ForEach-Object {
                $PSItem.picture = 1
            }

        } else {
            Write-Host "Error: Student Photos path has been specified but can not be found. Please check the path." -ForegroundColor Red
        }
    }

    #No header file on the students.sd file.
    $students | ConvertTo-CSV -Delimiter '|' -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1 | Out-File -Path $PSScriptRoot\hallpass\students.sd -Force
    if ($islinux) { (Get-Content -Raw -Path $PSScriptRoot/hallpass/students.sd) -replace '\n',"`r`n" | Set-Content -Path $PSScriptRoot/hallpass/students.sd }
    $filesToZip.Add("hallpass\students.sd") | Out-Null

    
} else {
    Write-Host "Error: There are no students to process. Please check your settings.ps1 file and that you have the proper buildings specified."
    exit 1
}

if ($includeGuardians) {

    #we save the file instead of Get-CognosReport in case the Cognos Report fails it would use the cached file.
    if ($IncludeAllContacts) {
        Save-CognosReport -report allcontacts -TeamContent -cognosfolder "_Shared Data File Reports\HallPass" -savepath "$PSScriptRoot\files" -Filename "guardians.csv"
    } else {
        Save-CognosReport -report guardians -TeamContent -cognosfolder "_Shared Data File Reports\HallPass" -savepath "$PSScriptRoot\files"
    }

    #pull in guardians for only valid students filtered out above
    $guardians = Import-CSV $PSScriptRoot\files\guardians.csv | Where-Object { $studentIds -contains $PSItem.'student id' }

    if (-Not($includeGuardianPhoneNumbers)) {
        $guardians = $guardians | Select-Object -ExcludeProperty phone | Select-Object *,phone
    }

    if ($IgnoreGuardianOrder) {
        $guardians | ForEach-Object {
            $PSitem.sortorder = 0
        }
    }
    
    #There has to be valid data before we continue.
    if ($guardians.Count -ge 1) {
        #No header file on the students.sd file.
        $guardians | ConvertTo-CSV -Delimiter '|' -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1 | Out-File -Path $PSScriptRoot\hallpass\guardians.gd -Force
        if ($islinux) { (Get-Content -Raw -Path $PSScriptRoot/hallpass/guardians.gd) -replace '\n',"`r`n" | Set-Content -Path $PSScriptRoot/hallpass/guardians.gd }
        $filesToZip.Add("hallpass\guardians.gd") | Out-Null

    } else {
        Write-Host "Error: There are no guardians to process. Please check your settings.ps1 file and that you have the proper buildings specified."
        exit 1
    }
}

try {
    Compress-Archive -Path $filesToZip -CompressionLevel Optimal -DestinationPath ".\hallpass\$($filename)s.zip" -Force
} catch {
    Write-Host "Error: Failed to create ZIP file."
    exit 1
}

if ($IncludeFaculty) {
    # $schoolids = @{}
    
    # Get-CogSchool | ForEach-Object {
    #     $schoolids.($PSitem.School_name) = [int]$PSItem.School_id
    # }

    Save-CognosReport -report faculty -TeamContent -cognosfolder "_Shared Data File Reports\HallPass" -savepath "$PSScriptRoot\files"
    Save-CognosReport -report faculty_locations -TeamContent -cognosfolder "_Shared Data File Reports\HallPass" -savepath "$PSScriptRoot\files"

    $facultyLocations = Import-CSV $PSScriptRoot\files\faculty_locations.csv | Where-Object { $validbuildings -contains $PSItem.'school id' }

    if ($facultyRFID) {
        $facultyLocations | ForEach-Object {
            $PSitem.'school id' = 0 #Per Nate at HallPass. #$schoolids.($PSItem.'school id')
            $PSitem.rfid = "$($PSItem.'employee id')$($RFIDSuffix)"
        }
    }

    $validFacultyIds = $facultyLocations | Select-Object -ExpandProperty 'employee id'
    $faculty = Import-CSV $PSScriptRoot\files\faculty.csv | Where-Object { $validFacultyIds -contains $PSItem.'employee id' }

    $faculty | ConvertTo-CSV -Delimiter '|' -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1 | Out-File -Path $PSScriptRoot\hallpass\faculty.fd -Force
    if ($islinux) { (Get-Content -Raw -Path $PSScriptRoot/hallpass/faculty.fd) -replace '\n',"`r`n" | Set-Content -Path $PSScriptRoot/hallpass/faculty.fd }
    $facultyLocations | ConvertTo-CSV -Delimiter '|' -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1 | Out-File -Path $PSScriptRoot\hallpass\faculty.ld -Force
    if ($islinux) { (Get-Content -Raw -Path $PSScriptRoot/hallpass/faculty.ld) -replace '\n',"`r`n" | Set-Content -Path $PSScriptRoot/hallpass/faculty.ld }

    try {
        Compress-Archive -Path ($authorizationFile,"hallpass\faculty.fd","hallpass\faculty.ld") -CompressionLevel Optimal -DestinationPath ".\hallpass\$($filename)f.zip" -Force
    } catch {
        Write-Host "Error: Failed to create ZIP file."
        exit 1
    }

}

#Upload File
if ($username -and $password) {
    & "$currentPath\bin\pscp.exe" @('-pw',"$password",'-batch','-hostkey',"$hostkey",".\hallpass\$($filename)*.zip","$($username)@push.starthallpass.com:")
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to upload file." -ForegroundColor Red
        exit 1
    } else {
        Write-Host "Info: Successfully uploaded zip file to HallPass."
        exit 0
    }
}
