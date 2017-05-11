Function Render-MultiPartFormFields
{
    Param(
        $FieldsHash,
        $Boundary
    )
    $ReturnString = "";
    ForEach ($Field in $FieldsHash.GetEnumerator())
    {
        if ($Field.Key -eq "fFile")
        {
            $ReturnString += $Boundary + "`r`n"
            $ReturnString += "Content-Disposition: form-data; name=""fFile""; filename=""upload.txt""`r`n"
            $ReturnString += "Content-Type: text/plain`r`n`r`n"
            $ReturnString += Get-Content $($Field.Value) -Raw
            $ReturnString += "`r`n"
        }
        else {
            $ReturnString += $Boundary + "`r`n"
            $ReturnString += "Content-Disposition: form-data; name=""$($Field.Key)""`r`n`r`n"
            $ReturnString += "$($Field.Value)`r`n"
        }
    }
    
    $ReturnString +="$Boundary--`r`n"

    $ReturnString
}

Function Upload-Contacts
{
    <#
    .PARAMETER UserName
        UserName for the BlackBoard Connect Upload Service

    .PARAMETER Password
        Password for the BlackBoard Connect Upload Service
    
    .PARAMETER ContactType
        The ContactType to assign to contacts in this upload.
                    
    .PARAMETER RefreshType
        All contacts of the selected tpye wil be removed and replaced by Contact Type data in your import file.  
        Selecting anything other than "None" presents a risk of data loss.  
        Selecting "None" will only affect the contacts contained in your file.
                    
    .PARAMETER PreserveData
        Will only update the fields provided in the upload.  
        Fields with blank data for a contact will be cleared.  
        ReferenceCode is required when this is selected.
    #>
    Param(
        [String]
        $UserName,
        [String]
        $Password,
        [ValidateSet("All","Student","Admin","Faculty","Staff","Other")]
        [String]
        $ContactType,
        [ValidateSet("All","Student","Admin","Faculty","Staff","Other","None")]
        [String]
        $RefreshType="None",
        [bool]
        $PreserveData = $true ,
        [String]
        $UploadFilePath
    )

    $Boundary = "-----------------------------AaB03x"
    
    $UploadFields = @{}
    $UploadFields['fNTIUser'] = $UserName
    $UploadFields['fNTIPassEnc'] = $Password
    $UploadFields['fContactType'] = $ContactType
    if ( $RefreshType -ne "None" )
    {
        $UploadFields['fRefreshType'] = $RefreshType
    }
    $UploadFields['fPreserveData'] = [int]$PreserveData
    $UploadFields['fFile'] = $UploadFilePath
    $UploadFields['fSubmit'] = 1

    $PostBody = $(Render-MultiPartFormFields -FieldsHash $UploadFields -Boundary $Boundary)

    $Response = Invoke-WebRequest -Uri "https://www.blackboardconnected.com/contacts/importer_portal.asp?qDest=imp" -Method POST -Body $PostBody -ContentType "multipart/form-data; boundary=$Boundary" -UserAgent "Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)"    
    
    $Response

}

Function New-AttendanceCall
{
 <#
    .PARAMETER UserName
        UserName for the BlackBoard Connect Upload Service

    .PARAMETER Password
        Password for the BlackBoard Connect Upload Service

    .PARAMETER ContactType
        The ContactType to assign to contacts in this upload.

    .PARAMETER TemplateName
        Name of the BlackBoard Connect Attendance Template to use

    .PARAMETER CallTime
        This is the time the call will be scheduled on the website.
        Defaults to 7:30:00 PM
        It should be formatted by XX:XX:XX AM. 
        Set blank if you would like to schedule immediately 

    .PARAMETER TimeZone
        Your TimeZone
        Defaults to ET

    .PARAMETER AttCode
        List of Attendance Codes

    .PARAMETER AttendanceFile
        File containing Attendance Call Data
#>
    Param(
        [String]
        $UserName,
        [String]
        $Password,
        [ValidateSet("All","Student","Admin","Faculty","Staff","Other")]
        [String]
        $ContactType,
        [String]
        $TemplateName,
        [String]
        $CallTime = "07:30:00 PM",
        [String]
        [ValidateSet("HA","AK","PT","MT","CT","ET","AT")]
        $TimeZone = "ET",
        [String]
        $AttCode,
        [String]
        $AttendanceFile
    )

    if (-Not $(Test-Path $AttendanceFile))
    {
        Throw "Attendance File required."
    }
    
    $Boundary = "-------------------------AaB03x"
    
    $UploadFields = @{}
    $UploadFields['fNTIUser'] = $UserName
    $UploadFields['fNTIPass'] = $Password
    $UploadFields['fContactType'] = $ContactType
    $UploadFields['fMessageTitle'] = $TemplateName
    $UploadFields['fNTITime'] = $CallTime
    $UploadFields['fNTITimeZone'] = $TimeZone
    $UploadFields['fFile'] = $AttendanceFile

    $PostBody = $(Render-MultiPartFormFields -FieldsHash $UploadFields -Boundary $Boundary)

    $Response = Invoke-WebRequest -Uri "https://upload.blackboardconnect.com/AutoNotification/SendNotification" -Method POST -Body $PostBody -ContentType "multipart/form-data; boundary=$Boundary" -UserAgent "Mozilla/4.0 (compatible; Win32; WinHttp.WinHttpRequest.5)"    
    
    $Response
}