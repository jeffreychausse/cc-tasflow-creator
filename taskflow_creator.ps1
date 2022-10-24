# Description: Powershell script to automate the creation of taskflows for projectors
# By: Jeffrey Chausse (2022)

###################################
########## Module Import ##########
###################################

Import-Module Mdbc # https://github.com/nightroman/Mdbc

##########################################
########## Variable Declaration ##########
##########################################

$ccIp = "172.27.93.37" # Enter the Control Center IP Address
$ProjectId = "jeff" # Enter the project identifier
$TaskflowCreatedCount = 0

#########################################
########## Function Definition ##########
#########################################

# Request projector list from CC variables
function GetProjectorList {
    $Body = Get-Content .\Variable.Get.json -raw # Read JSON file
    $Response = Invoke-WebRequest -Uri http://$($ccIp):3030/sc-datastore/projectData/variable -Method POST -Body $Body -ContentType "application/json" | ConvertFrom-Json

    # Add ALL field to projector list to generate the tasks to control ALL projectors
    if ($Response.result.data.value.Count -gt 1) {
        $Response.result.data.value += [PSCustomObject]@{key = "ALL" ; value = ""}
    }

    return $Response.result.data.value
}

#############################
########## MongoDB ##########
#############################

# Connect the client
Connect-Mdbc -ConnectionString "mongodb://$($ccIp):27055"

# Then get the database
$Database = Get-MdbcDatabase sc-datastore
$Collection = Get-MdbcCollection -Name tasks -Database $Database

# Initialize variables with each task oid
$lampOn_oid = Get-MdbcData -Filter '{"internalName": "lampOn"}' -Collection $Collection | ForEach-Object {$_._id.ToString()}
$lampOff_oid = Get-MdbcData -Filter '{"internalName": "lampOff"}' -Collection $Collection | ForEach-Object {$_._id.ToString()}
$openShutter_oid = Get-MdbcData -Filter '{"internalName": "openShutter"}' -Collection $Collection | ForEach-Object {$_._id.ToString()}
$closeShutter_oid = Get-MdbcData -Filter '{"internalName": "closeShutter"}' -Collection $Collection | ForEach-Object {$_._id.ToString()}

#Write-Host "$lampOn_oid `n$lampOff_oid `n$openShutter_oid `n$closeShutter_oid"

#######################################
########## ControlCenter API ##########
#######################################

# Read JSON file and convert to Object
$TaskflowObj_Master = Get-Content '.\Taskflow.UpdateOrCreate(Projector).json' -raw | ConvertFrom-Json

$ProjectorList = GetProjectorList

foreach ($Projector in $ProjectorList) {
    for ($i = 0; $i -lt 4; $i++) {
        switch ($i) { # Set variables for current value of $i
            0 {  
                $TaskSource = $lampOn_oid
                $TaskInternalName = "lampOn"
                $TaskflowNameSuffix = "Lamp ON"
            }
            1 {  
                $TaskSource = $lampOff_oid
                $TaskInternalName = "lampOff"
                $TaskflowNameSuffix = "Lamp OFF"
            }
            2 {  
                $TaskSource = $openShutter_oid
                $TaskInternalName = "openShutter"
                $TaskflowNameSuffix = "Shutter OPEN"
            }
            3 {  
                $TaskSource = $closeShutter_oid
                $TaskInternalName = "closeShutter"
                $TaskflowNameSuffix = "Shutter CLOSE"
            }
            Default {Write-Debug "ERROR: switch input not in range"}
        }

        $TaskflowObj = $TaskflowObj_Master #Restore the original obj

        if ($Projector.key -eq "ALL") {
            
            $TaskflowObj.params.items[0].taskParams[0].value = @() # Assign empty array to erase original content
            
            foreach ($Prj in $ProjectorList | Where-Object {$_.key -ne "ALL"}) { # For each prj in ProjectorList, except ALL
                    $TaskflowObj.params.items[0].taskParams[0].value += @{key = $Prj.key ; value = $Prj.value} # Insert a new key/value pair for every projector
            }

            $TaskflowObj.params.items[0].taskSource = $TaskSource # Set task source oid
            $TaskflowObj.params.items[0].taskInternalName = $TaskInternalName # Set task internal name
            $TaskflowObj.params.displayName = "ALL PRJ - " + $TaskflowNameSuffix # Set taskflow display name
            $TaskflowObj.params.internalName = $TaskflowObj.params.displayName.Replace(" ", "_") | ForEach-Object {$_.ToLower()} # Set taskflow internal name
            $TaskflowObj.params.projectIdentifier = $ProjectId # Set project identifier
        }
        else {
            $TaskflowObj.params.items[0].taskParams[0].value[0].key = $Projector.key # Set projector hostname
            $TaskflowObj.params.items[0].taskParams[0].value[0].value = $Projector.value # Set projector oid
            $TaskflowObj.params.items[0].taskSource = $TaskSource # Set task source oid
            $TaskflowObj.params.items[0].taskInternalName = $TaskInternalName # Set task internal name
            $TaskflowObj.params.displayName = $Projector.key.Substring($Projector.key.IndexOf("-") + 1) + " - " + $TaskflowNameSuffix # Set taskflow display name
            $TaskflowObj.params.internalName = $TaskflowObj.params.displayName.Replace(" ", "_") | ForEach-Object {$_.ToLower()} # Set taskflow internal name
            $TaskflowObj.params.projectIdentifier = $ProjectId # Set project identifier  
        }

        # Convert to JSON before sending
        $TaskflowJson = ConvertTo-Json -InputObject $TaskflowObj -Depth 100
        #$TaskflowJson

        try {
            $Response = Invoke-WebRequest -Uri http://$($ccIp):3030/sc-datastore/projectData/taskFlow -Method POST -Body $TaskflowJson -ContentType "application/json"
        }
        catch {
            Write-Error $Response
            # $StatusCode = $_.Exception.Response.StatusCode
            # $ErrorMessage = $_.ErrorDetails.Message
        
            # Write-Error "$([int]$StatusCode) $($StatusCode) - $($ErrorMessage)"
        }

        $TaskflowCreatedCount ++ # Increment taskflow counter
    }
}

Write-Host "$($TaskflowCreatedCount) taskflows were created"

#################################
########## End Of File ##########
#################################