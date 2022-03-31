#Requires -Version 7.1

<#

HallPass
Craig Millsap - 2/2022
Data Upload for students, photos, and guardians.
Looking for more Automation? https://www.camtehcs.com

#>

$currentPath=(Split-Path ((Get-Variable MyInvocation -Scope 0).Value).MyCommand.Path)
$hostkey = '8e:20:fc:57:99:c1:0b:9c:42:b9:bb:3d:fd:a4:89:27'

if (-Not(Test-Path $currentPath\hallpass)) { New-Item -ItemType Directory hallpass }
if (-Not(Test-Path $currentPath\files)) { New-Item -ItemType Directory files }

if (-Not(Test-Path .\settings.ps1)) {
    Write-Host "Error: Failed to find settings.ps1 file."    
    exit 1
} else {
    . .\settings.ps1
}

$filesToZip = New-Object System.Collections.ArrayList

if (-Not($authorizationFile = Get-ChildItem -Filter *.ud -Recurse -File)) {
    Write-Host "Error: Authorization File not found in any directory here. Please save the .ud file provided by HallPass to this directory."
    exit 1
} else {
    $fileName = $authorizationFile[0].baseName
    $authorizationFile = $authorizationFile[0].FullName
    $filesToZip.Add($authorizationFile) | Out-Null
    $filesToZip.Add("files\config.ini") | Out-Null
}

..\CognosDownload.ps1 -report students -teamcontent -cognosfolder "_Shared Data File Reports\HallPass" -savepath .\files

$students = Import-CSV .\files\students.csv | Where-Object { $validbuildings -contains $PSItem.'school id' }

if ($removeHomeroomTeachers) {
    $students = $students | Select-Object -ExcludeProperty teacher | Select-Object *,teacher
}

#There has to be valid data before we continue.
if ($students.Count -ge 1) {

    $studentIds = $students | Select-Object -ExpandProperty 'student id'

    #lets include photos if we can find them and they are the right size.
    if ($includeStudentPhotos) {
        if (Test-Path $studentPhotosFolder) {

            if ((Get-Date -Format dddd) -eq 'Friday') {

                $studentsWithPhotoIds = @()

                $photos = Get-ChildItem -Path $studentPhotosFolder -Recurse -File -Filter "*.jpg" | Where-Object { $PSItem.Length -lt 150kb }
                $photosGrouped = $photos | Group-Object -Property BaseName
                $photosHashed = $photosGrouped | Group-Object -Property Name -AsHashTable

                $studentIds | ForEach-Object {
                    if ($photosHashed."$PSitem") {
                        $studentsWithPhotoIds += $PSitem
                        $photo = $photosHashed."$PSitem"
                        $filesToZip.Add("$($photo.Group[0].FullName)") | Out-Null
                    }
                }

                #now we need to mark the students as having a photo in the output CSV.
                $students | Where-Object { $studentsWithPhotoIds -contains $PSItem.'student id' } | ForEach-Object {
                    $PSItem.picture = 1
                }

            } else {
                Write-Host "Info: HallPass only processes Student Photos on Fridays. We will not process all student photos."
            }

        } else {
            Write-Host "Error: Student Photos path has been specified but can not be found. Please check the path." -ForegroundColor Red
        }
    }

    #No header file on the students.sd file.
    $students | ConvertTo-CSV -Delimiter '|' -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1 | Out-File -Path .\hallpass\students.sd -Force
    $filesToZip.Add("hallpass\students.sd") | Out-Null

    
} else {
    Write-Host "Error: There are no students to process. Please check your settings.ps1 file and that you have the proper buildings specified."
    exit 1
}

if ($includeGuardians) {
    ..\CognosDownload.ps1 -report guardians -teamcontent -cognosfolder "_Shared Data File Reports\HallPass" -savepath .\files

    #pull in guardians for only valid students filtered out above
    $guardians = Import-CSV .\files\guardians.csv | Where-Object { $studentIds -contains $PSItem.'student id' }

    if (-Not($includeGuardianPhoneNumbers)) {
        $guardians = $guardians | Select-Object -ExcludeProperty phone | Select-Object *,phone
    }
    
    #There has to be valid data before we continue.
    if ($guardians.Count -ge 1) {
        #No header file on the students.sd file.
        $guardians | ConvertTo-CSV -Delimiter '|' -NoTypeInformation -UseQuotes AsNeeded | Select-Object -Skip 1 | Out-File -Path .\hallpass\guardians.gd -Force
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

#Upload File
if ($username -and $password) {
    & "$currentPath\bin\pscp.exe" @('-pw',"$password",'-batch','-hostkey',"$hostkey",".\hallpass\$($filename)s.zip","$($username)@push.starthallpass.com:")
    if ($LASTEXITCODE -ne 0) {
        Write-Host "Error: Failed to upload file." -ForegroundColor Red
        exit 1
    } else {
        Write-Host "Info: Successfully uploaded zip file to HallPass."
        exit 0
    }
}
