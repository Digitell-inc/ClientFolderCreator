#Form ----------------------------------------------------------

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$form                       = New-Object system.Windows.Forms.Form
$form.ClientSize            = New-Object System.Drawing.Point(400,307)
$form.text                  = "Client Folder Creation Tool"
$form.TopMost               = $false
$form.BackColor             = [System.Drawing.ColorTranslator]::FromHtml("#e5e5e5")

$userInput                       = New-Object system.Windows.Forms.TextBox
$userInput.multiline             = $false
$userInput.width                 = 163
$userInput.height                = 20
$userInput.location              = New-Object System.Drawing.Point(50,45)
$userInput.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',12)

$pathInput                       = New-Object system.Windows.Forms.ComboBox
$pathInput.width                 = 163
$pathInput.height                = 20
$pathInput.location              = New-Object System.Drawing.Point(50,109)
$pathInput.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',12)


$submit                          = New-Object system.Windows.Forms.Button
$submit.text                     = "Create"
$submit.width                    = 70
$submit.height                   = 35
$submit.location                 = New-Object System.Drawing.Point(288,35)
$submit.Font                     = New-Object System.Drawing.Font('Myanmar Text',10)

$outputBox                        = New-Object system.Windows.Forms.TextBox
$outputBox.width                  = 371
$outputBox.height                 = 139
$outputBox.multiline              = $True 
$outputBox.ReadOnly               = $True
$outputBox.location               = New-Object System.Drawing.Point(16, 153)
$outputBox.ScrollBars             = "Vertical"
$outputBox.Font                   = New-Object System.Drawing.Font('Myanmar Text',12) 

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "Meeting Name"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(56,21)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Client Folder"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(56,83)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$form.controls.AddRange(@($userInput,$pathInput,$submit,$outputBox,$Label1,$Label2))


#----------------------------------------------------------------

#Static Variables
$sourcePath = "P:\2022\ClientTemplate"
$targetPath = "P:\2022"

#User Variables
$userMeetingName = ""
$userClientName = ""

#Moved to line 140
#Initialize App
#createListOfFolders -rootPath "P:\2021"

#User Warning
$outputBox.Text = ( "2022 Client Creation Tool `r`n" )

# Event Handlers---------------------------------------------------------------------------------------------

# Submit Button Click event handler
$submit.Add_Click( {
    #Read Meeting textbox, If blank Throw Error
    $userMeetingName = $userInput.Text
    if($userMeetingName -eq ""){
        $outputBox.Text = ( "Error: Meeting Name Cannot be Blank" )
        Exit
    }

    #Read dropdown box
    $selectedClient = $pathInput.SelectedItem
    #if no dropdowns have been selected get text content, if nothing has been selected Throw Error
    if($pathInput.SelectedItem -eq $null){
        $selectedClient = $pathInput.Text
        if($selectedClient -eq ""){
            $outputBox.Text = ( "Error: Client Folder Cannot be Blank" )
            Exit
        }
    }
    
    $userClientName = $selectedClient

    #Debugging: Log User input passed to copy
    write-host "Target-${userClientName} MeetingFolder-${userMeetingName}"

    #Pass Input Variables into rCopy
    rCopy -Source $sourcePath -DestinationRoot $targetPath -ClientFolder $userClientName -Meeting $userMeetingName
   
} )

#---------------------------------------
#               Functions
#---------------------------------------

# enable or disable a button
function toggleButton{
    param(
        $element
    )
    if($element.enabled -eq $true){
        $element.enabled = $false
    }
    else {
        $element.enabled = $true
    }
}

#Get List of Folders
function createListOfFolders{
    param(
        [string]$rootPath
    )

    $pathInput.Items.add( $rootPath )

    ForEach( $file in Get-ChildItem -Path $rootPath -Directory){
        $pathInput.Items.Add( $file )
    }
     
}

#Check for P:\
function checkDrives{
    # Template folder is unavailible
    if(!(Test-Path -Path $sourcePath)){
       $outputBox.Text = ( "The Path of the Source Template is unreachable. (The P:\ Drive does not exist) `r`n" )
       toggleButton -element $submit 
    }
    # Network drive is unavailible
    elseif (!(Test-Path -Path $targetPath)) {
        $outputBox.Text = 'Warning: The P:\ Drive is unreachable. Contact IT (Tony or Chris) for setup.'
        toggleButton -element $submit 
    }
    # Drives and folders are availible 
    else{
       $outputBox.AppendText("P:\ Connected. READY")
        #Initialize  list of FoldersApp
        createListOfFolders -rootPath $targetPath
    }

}

# Check Drives
checkDrives


#Create Permissions 
# Microsoft cant make up their fucking minds as to how ACL works so all permissions errors were prorobably cause by updates
# They have changed ACL twice now (as of 5/21/21) since this was written 
function createPermissions{
    param(
         [string]$meetingFolder
    )
            # Get ACL From Template Folder
            $AllowSource = "${sourcePath}\MeetingName\Content"
            $DenySource = "${sourcePath}\MeetingName\Artwork"

            # Paths for target folders to edit
            $AllowFolder = "${meetingFolder}\Content"
            $DenyFolders = Get-ChildItem "${meetingFolder}" -Exclude "Content"  

            # Filter by only the properties that we want to change
            $AllowACL = (Get-Item $AllowSource).GetAccessControl('Access')
            $DenyACL = (Get-Item $DenySource).GetAccessControl('Access')

            #Set ACL on root Content folder and recursivley everything inside of it
               Get-ChildItem $AllowFolder -Recurse | Set-Acl -ACLObject $AllowACL
               Set-Acl -ACLObject $AllowACL -Path $AllowFolder

            # Loop Through Children in $deny Folder and Set ACL on them
                ForEach($folder in $DenyFolders){
                Set-Acl -ACLObject $DenyACL -path $folder   
                Get-ChildItem "$folder" -Recurse | Set-Acl -ACLObject $DenyACL
            }
}

#Copy Configuration
function rCopy{
    param(
        [string]$Source,
        [string]$DestinationRoot,
        [string]$ClientFolder,
        [string]$Meeting
        )
         
    #debug logging
    write-host "Source-$Source MeetingName-$Meeting ClientFolder-$ClientFolder Root-$DestinationRoot}"

    #create folder
    #Copy-Item -Path $Source -destination $DestinationRoot\$ClientFolder -Recurse
    
    
    # Check to see if client folder exists and changes adjusts target
    if(Test-Path -Path "$DestinationRoot\${ClientFolder}"){
        
        Copy-Item -Path "$Source\MeetingName" -destination $DestinationRoot\$ClientFolder -Recurse
        
        Start-Sleep 1
        
        Rename-Item "${DestinationRoot}\${ClientFolder}\MeetingName" "${DestinationRoot}\${ClientFolder}\${Meeting}"

        $outputBox.AppendText( "Success. Your folder: $DestinationRoot\${ClientFolder}\${Meeting} `r`n" )
    }
    else{
        
        Copy-Item -Path "${Source}" -destination $DestinationRoot\$ClientFolder -Recurse

        Start-Sleep 1
        
        Rename-Item "${DestinationRoot}\${ClientFolder}\MeetingName" "${DestinationRoot}\${ClientFolder}\${Meeting}"
  
        $outputBox.AppendText( "Success. Your folder: $DestinationRoot\${ClientFolder}\${Meeting} `r`n" )
    }
    
    #add Permissions
    createPermissions -meetingFolder "${DestinationRoot}\${ClientFolder}\${Meeting}"


}

# End Form
[void]$form.ShowDialog()