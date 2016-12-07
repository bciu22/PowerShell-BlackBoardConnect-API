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
            $ReturnString += $Boundary + "`n"
            $ReturnString += "form-data; name=""fFile""; filename=""$($Field.Value)"""
            $ReturnString += "Content-Type: text/plain`n`n"
            $ReturnString += Get-Content $($Field.Value)
        }
        else {
            $ReturnString += $Boundary + "`n"
            $ReturnString += "Content-Disposition: form-data; name=""$($Field.Key)""`n`n"
            $ReturnString += "$($Field.Value)`n"
        }
        

    }

    
    $ReturnString +="$Boundary--"

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
        The ContactType to assign to contacts in this uplaod.
                    
    .PARAMETER RefreshType
        Any of the Contact Types you select below will be removed and replaced by Contact Type data in your import file.
                    
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
        [ValidateSet("All","Student","Admin","Faculty","Staff","Other")]
        [String]
        $RefreshType,
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
    $UploadFields['fRefreshType'] = $RefreshType
    $UploadFields['fPreserveData'] = [int]$PreserveData
    $UploadFields['fFile'] = "Users.csv"
    $UploadFields['fSubmit'] = 1

    $Body = $(Render-MultiPartFormFields -FieldsHash $UploadFields -Boundary $Boundary)

    Write-Host $Body

    $Response = Invoke-WebRequest -Uri "https://www.blackboardconnected.com/contacts/importer_portal.asp?qDest=imp" -Method POST -Body $Body -ContentType "multipart/form-data; boundary=$Boundary" 
    $Response

}
