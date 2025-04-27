
if($Null-eq $(get-module -ListAvailable PSSQLite)){
    Get-PSRepository | Set-PSRepository -InstallationPolicy Trusted
    Install-Module PSSQLite
}
if($Null-eq $(get-module PSSQLite)){
    Import-Module PSSQLite
}

#region Variables and enums
$RulesDBPath = 'C:\programdata\T4W\.db\Tagging.adb'  #Database for AutoTagging
$TagDBPath = 'C:\programdata\T4W\.db\Tagging.odb'    #Primary Database for Tags and Paths


$T4WRegexGUID           = [regex] "[{]?(?<guid>[a-zA-Z0-9-]{32,36})[{]?"
$T4WRegexVolume         = [regex] "(?i)^(?<volume>\\\\[?]\\Volume[\\]?{[a-z0-9-]{32,36}})(?<path>.*)"
$T4WRegexUNC            = [regex] "(?i)^\\\\(?<server>[^?$\\]+)\\(?<share>[^\\]+)(?<path>.*)$"
$T4WRegexDriveLetter    = [regex] "(?i)^(?<driveletter>[a-z])[:](?<path>\\.*)"
$T4WRegexDriveFormat    = [regex] "(?i)^(?<driveletter>[a-z])[:]?[\\]?"




enum T4W_IconColor {
    Orange      =   00
    Red         =   01
    Red_Dark    =   02
    Pink        =   03
    Purple      =   04
    Purple_Dark =   05
    Blue_Dark   =   06
    Blue        =   07
    Blue_Light  =   08
    Teal        =   09
    Green       =   10
    Green_Dark  =   11
    Brown_Dark  =   12
    Brown       =   13
    Silver      =   14
    Grey        =   15
    Grey_Dark   =   16
    Black       =   17
    White       =   18
    Yellow      =   19

}
enum T4W_IconShape {
    Tag                     =   00
    Circle                  =   20
    Circle_Outline          =   40
    Rounded_Square          =   60
    Rounded_Square_Outline  =   80
    Triangle                =   100
    Triangle_Outline        =   120 
    Square                  =   140
    Square_Outline          =   160 
    Pentagon                =   180
    Pentagon_Outline        =   200
    Hexagon                 =   220
    Hexagon_Outline         =   240 
    Octagon                 =   260
    Octagon_Outline         =   280
    Diamond                 =   300
    Diamond_Outline         =   320
    Star                    =   340
    Heart                   =   360
    CheckMark               =   380
    X                       =   400
    Plus                    =   420
    ArrowRight              =   440
    ArrowLeft               =   460
    Balloon                 =   480
}

enum T4W_TagGroupState  {
    Enable = 1
    Closed = -2
}

enum T4W_TagType {
    Tag      = 1
    TagGroup = -1

}

enum T4W_GUIDType {
    File
    File_Comment
    File_Tag
    Folder
    Folder_Comment
    Folder_Tag
    Tag
    Tag_Comment
}

enum AutoTagging_State{
    Rescan   = 8 #Find New Changes
    Rematch  = 16   #Rescan + Match Previous Files
    Relink   = 32    #Rematch +  Link the matching tags again
    None     = 0

}

enum AutoTagging_SearchIn{
    Content   = 1
    Filename  = 2
    Both      = 3
}

#endregion

#region MISC

function Remove-T4WOrphanedObjects {
    [cmdletbinding()]
    param(
        $Database = $TagDBPath
    )

    $Query =  " select * FROM xfis "
    $Query += " LEFT OUTER JOIN fis "
    $Query += " ON xfi_fi = fi_id "
    $Query += " WHERE fi_id IS NULL"

    $Results = Invoke-SqliteQuery -DataSource $Database -query $Query
    $OrphanedIDs = $Results | Select-Object -Unique -ExpandProperty xfi_guid
    foreach($OrphanedID in $OrphanedIDs){
        if("$OrphanedID".Length -gt 36){
            Remove-T4WFileTag -GUID $OrphanedID
        }
    }

    Remove-Variable -Name 'Results','OrphanedIDs' -ErrorAction SilentlyContinue

    $Query =  " select * FROM xfos "
    $Query += " LEFT OUTER JOIN fos "
    $Query += " ON xfo_fo = fo_id "
    $Query += " WHERE fo_id IS NULL"

    $Results = Invoke-SqliteQuery -DataSource $Database -query $Query
    $OrphanedIDs = $Results | Select-Object -Unique -ExpandProperty xfo_guid
    foreach($OrphanedID in $OrphanedIDs){
        if("$OrphanedID".Length -gt 36){
            Remove-T4WFolderTag -GUID $OrphanedID
        }
    }
    
    Remove-Variable -Name 'Results','OrphanedIDs' -ErrorAction SilentlyContinue

    $Query =  " select * FROM cfis "
    $Query += " LEFT OUTER JOIN fis "
    $Query += " ON cfi_fi = fi_id "
    $Query += " WHERE fi_id IS NULL"

    $Results = Invoke-SqliteQuery -DataSource $Database -query $Query
    $OrphanedIDs = $Results | Select-Object -Unique -ExpandProperty cfo_id
    foreach($OrphanedID in $OrphanedIDs){
        if([int]$OrphanedID -gt 0){
            Remove-T4WFileComment -CommentID $OrphanedID
        }
    }

    Remove-Variable -Name 'Results','OrphanedIDs' -ErrorAction SilentlyContinue

    $Query =  " select * FROM cfos "
    $Query += " LEFT OUTER JOIN fos "
    $Query += " ON cfo_fo = fo_id "
    $Query += " WHERE fo_id IS NULL"

    $Results = Invoke-SqliteQuery -DataSource $Database -query $Query
    $OrphanedIDs = $Results | Select-Object -Unique -ExpandProperty cfi_id
    foreach($OrphanedID in $OrphanedIDs){
        if([int]$OrphanedID -gt 0){
            Remove-T4WFolderTag -CommentID $OrphanedID
        }
    }
    
    Remove-Variable -Name 'Results','OrphanedIDs' -ErrorAction SilentlyContinue



}

Function Convert-T4WIconToEnum {
    <#
    .SYNOPSIS
    Converts Integer Based Icon Back to Enums
    
    .DESCRIPTION
    Converts Integer Based Icon Back to Enums
    
    .PARAMETER IconID
    [Integer] Icon ID
    
    .EXAMPLE
    Convert-T4WIconToEnum -IconID 325

      IconShape   IconColor
      ---------   ---------
Diamond_Outline Purple_Dark
    
    #>
    param(
        [Parameter(Mandatory)]
        [int]$IconID
    )
    
    process{

        if($IconID -lt 0 -or $Icon -gt 499){ Write-Error -ErrorAction Stop -Message "Value outside the range of an Icon"}

        $IconShape =$Null
        $IConColor = $Null

        :iconshape foreach($key in $([Enum]::GetNames([T4W_IconShape]) )){
            $NewVal=$IconID-([int][T4W_IconShape]::$key.value__)
            
            if($NewVal -ge 0 -and $NewVal -le 19){
                $IconShape=[t4W_IconShape] $Key
                break iconshape
            }
        }
        :iconcolor foreach($key in $([Enum]::GetNames([T4W_IconColor]) )){
            $NewVal=$IconID-([int][T4W_IconColor]::$key.value__) - [int]$IconShape
            
            if($NewVal -eq 0){
                $IconColor=[T4W_IconColor] $Key

                break iconcolor
            }
        }

        return [PSCustomObject]@{
            IconShape = [T4W_IconShape] $IconShape
            IconColor = [T4W_IconColor] $IconColor
       }
    }



    
}


Function New-T4WGUID {
<#
.SYNOPSIS
Used to Generate a new Unique GUID

.DESCRIPTION
Will check the archived tables and active tables to validate that the GUID generated is Unique

.PARAMETER Type
Tables to search through

.PARAMETER Database
Database file

.EXAMPLE
 New-T4WGuid -Type File_Tag    
0f58962d-0d2a-48cc-99a9-a6d300e58c7c

#>    
    [cmdletbinding()]
    [alias('Create-GUID','Create-T4WGUID')]
    param(
        [Parameter(Mandatory=$True,Position=0)]
        [T4W_GUIDType]$Type,
        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath
    )
    begin{


        $GUID = New-Guid
        Write-Verbose "New Guid Created: {$($Guid.Guid)}"
        #CHECK the Guid for a collision
        if($Type -eq [T4W_GUIDType]::File ){
            

            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT fi_guid FROM fis WHERE fi_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dfi_guid FROM dfis WHERE dfi_guid = '{$($GUID.guid)}'"))
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }
            
        }elseif( $Type -eq [T4W_GUIDType]::File_Comment ) {
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT cfi_guid FROM cfis WHERE cfi_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dcfi_guid FROM dcfis WHERE dcfi_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }

        }elseif( $Type -eq [T4W_GUIDType]::File_Tag ){
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT xfi_guid FROM xfis WHERE xfi_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dxfi_guid FROM dxfis WHERE dxfi_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }
            
        }elseif( $Type -eq [T4W_GUIDType]::Folder ){
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT fo_guid FROM fos WHERE fo_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dfo_guid FROM dfos WHERE dxfo_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }

            
        }elseif( $Type -eq [T4W_GUIDType]::Folder_Comment ){
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT cfo_guid FROM cfos WHERE cfo_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dcfo_guid FROM dcfos WHERE dcfo_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }

        }elseif( $Type -eq [T4W_GUIDType]::Folder_Tag ){
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT xfo_guid FROM xfos WHERE xfo_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dxfo_guid FROM dxfos WHERE dxfo_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }


        }elseif( $Type -eq [T4W_GUIDType]::Tag ){
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dto_guid FROM dtos WHERE dto_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT to_guid FROM tos WHERE to_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }

        }elseif( $Type -eq [T4W_GUIDType]::Tag_Comment ){
            while(
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT cto_guid FROM ctos WHERE cto_guid = '{$($GUID.guid)}'")) -and
                ($Null -ne (Invoke-SqliteQuery -DataSource $Database -query "SELECT dcto_guid FROM dctos WHERE dcto_guid = '{$($GUID.guid)}'"))
            
            ){
                $GUID = New-Guid
                Write-Verbose "GUID Collision Occurred New Guid: {$($Guid.Guid)}"
            }

        }else{
            Write-Error -ErrorAction Stop -Message "Found a Missing Condition in Create-GUID"
        }
            
        return "$($Guid.guid)"

    }
}


Function Get-T4WPath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path
    )
    Process{
        $Return = [PSCustomObject]@{
            DriveLetter  = ''
            Volume       = ''
            Server       = ''
            Share        = ''
            RelativePath = ''
        }
        switch -Regex ($Path){
            $T4WRegexDriveLetter.ToString() {
                $REGEX = $T4WRegexDriveLetter.Match($Path)
                $Return.DriveLetter = "$($REGEX.Groups['driveletter'].Value)"
                $Return.Volume = Get-Volume | Where-Object DriveLetter -eq $Return.DriveLetter | Select-Object -ExpandProperty UniqueId
                $Return.RelativePath = "$($REGEX.Groups['path'].Value)"
                break
            }
            $T4WRegexVolume.ToString(){
                $REGEX = $T4WRegexVolume.Match($Path)
                $Return.Volume = "$($REGEX.Groups['volume'].Value)\"
                $Return.DriveLetter = Get-Volume | Where-Object UniqueId -eq $Return.Volume | Select-Object -ExpandProperty DriveLetter
                if($REturn.DriveLetter -ne ''){$Return.DriveLetter = "$($Return.DriveLetter):\"}
                $Return.RelativePath = "$($REGEX.Groups['path'].Value)"

                break
            }
            $T4WRegexUNC.ToString(){
                $REGEX = $T4WRegexUNC.Match($Path)

                $Return.Server = $REGEX.Groups['server'].value
                $Return.share   = $REGEX.Groups['share'].Value
                $Return.RelativePath = $REGEX.Groups['path'].Value

                #Is the Drive Mapped?
                $PSDrive = Get-PSDrive | Where-Object {$_.DisplayRoot -eq "\\$($Return.Server)\$($Return.share)"}

                if($Null -ne $PSDrive){
                    $Return.DriveLetter = "$($T4WRegexDriveFormat.match($PSDrive.Root).Groups['driveletter'].value):\"
                }
                if($REturn.Server -match "(?i)(localhost|127.0.0.\d+|${env:computername})"){
                    $Return.DriveLetter = $T4WRegexDriveFormat.Match($(Get-SmbShare | Where-Object Name -ieq $Return.share|Select-Object -ExpandProperty Path)).Groups['driveletter'].value
                    if($Return.DriveLetter -ne ''){
                        
                        $Return.Volume = Get-Volume | Where-Object DriveLetter -eq $Return.DriveLetter| Select-Object -ExpandProperty UniqueId
                        $Return.DriveLetter = "$($Return.DriveLetter):\"
                    }
                }
                break
            }
        }
        Return $Return
    }
}

function Get-T4WIcon {
    <#
    .SYNOPSIS
    Creates the IconID value
    
    .DESCRIPTION
    Converts the T4W_IconColor & T4W_IconShape enums to an integer
    
    .PARAMETER Shape
    Icon Shape
    
    .PARAMETER Color
    Icon Color 
    
    .EXAMPLE
    Get-T4WIcon -Shape x -Color Yellow
    419
    
    #>
        [cmdletbinding()]
        param(
            [Parameter(Mandatory=$True)]
            [T4W_IconShape]$Shape,
            [Parameter(Mandatory=$True)]
            [T4W_IconColor]$Color
        )
        Process{
            
            return [int]$Shape+[int]$Color
        }
    }
    
#endregion


#region Volumes
function Get-T4WVolumes {
<#
.SYNOPSIS
Gets Volumes found in Database

.DESCRIPTION
Gets Volumes found in Database

.PARAMETER ID
Volume ID

.PARAMETER Label
Volume Label

.PARAMETER Share
UNC Share

.PARAMETER Machine
UNC Server

.PARAMETER Letter
Drive Letter

.PARAMETER GUID
VOLUME ID or GUID

.PARAMETER SerialNumber
Volume Serial

.PARAMETER Database
Database Path

.EXAMPLE
get-t4wvolumes -id 1

id      : 1
serial  : 1e24fcsc
guid    : \\?\Volume{0a83034e-e083-43b5-8a28-9ba7cc2c7819}\
machine : Your-PC
share   : C$
letter  : C:\
label   : NVME
state   : 2
#>    
    [cmdletbinding(DefaultParameterSetName = 'Default')]
    param( 
        [Parameter(Mandatory=$True, ParameterSetName="UseID")]
        [int]$ID,
        [Parameter(Mandatory=$True, ParameterSetName="UseLabel")]
        [string]$Label,
        [ValidatePattern("^[^\\!@#%^&*()+=|\?,.<>{}\[\]]+$")]
        [Parameter(Mandatory=$True, ParameterSetName="UseShare")]
        [string]$Share,        
        [ValidatePattern("^[a-z0-9A-Z-_.]")]
        [Parameter(Mandatory=$True, ParameterSetName="UseShare")]
        [string]$Machine,
        [Parameter(Mandatory=$True, ParameterSetName="UseLetter")]
        [string]$Letter,
        [Parameter(Mandatory=$True, ParameterSetName="UseGUID")]
        [string]$GUID,
        [Parameter(Mandatory=$True, ParameterSetName="UseSerial")]
        [string]$SerialNumber,
        [Parameter(Mandatory=$False, ParameterSetName="Default")]
        [Parameter(Mandatory=$False, ParameterSetName="UseShare")]
        [Parameter(Mandatory=$False, ParameterSetName="UseLetter")]
        [Parameter(Mandatory=$False, ParameterSetName="UseName")]
        [Parameter(Mandatory=$False, ParameterSetName="UseID")]
        [string]$Database=$TagDBPath)
    begin{
        Write-Verbose "=============  Beginning Get-T4WVolumes"
        
    }
    process {
        

        switch($PsCmdlet.ParameterSetName){
            'UseID'{ 
                        $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                        $Query += " vb_label,vb_state "
                        $Query += " FROM vbs"
                        $Query += " WHERE vb_id = '$ID'"

                        break;
            }

            'UseLabel'{ 
                        $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                        $Query += " vb_label,vb_state "
                        $Query += " FROM vbs "
                        $Query += " WHERE vb_label LIKE '%$Label%'"
                        break;
            }
            'UseSerial'{
                $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                $Query += " vb_label,vb_state "
                $Query += " FROM vbs "
                $Query += " WHERE vb_serial LIKE '$SerialNumber'"
                break
            }
            'UseGUID'{ 
                switch -regex ($GUID){
                    $T4WRegexVolume.ToString() {
                        $VolumeGUID =  $T4WRegexVolume.match($GUID).Groups['volume'].value
                        Write-Verbose "GUID Volume: $VolumeGUID"
                        $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                        $Query += " vb_label,vb_state "
                        $Query += " FROM vbs "
                        $Query += " WHERE vb_guid LIKE '$VolumeGUID\'"
                        #Using LIKE here with no wildcards to ignore case sensitivity
                        break
                    }
                    $T4WRegexGUID.ToString() {
                        #Look for Standard Guid Format without the volume path
                        $VolumeGUID = [guid]$T4WRegexGUID.match($GUID).Groups['guid'].value
                        Write-Verbose "GUID Volume: $VolumeGUID"
                        $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                        $Query += " vb_label,vb_state "
                        $Query += " FROM vbs "
                        $Query += " WHERE vb_guid LIKE '%{$VolumeGUID}%'"                        

                        break
                    }
                    default {
                        Write-Error -ErrorAction STOP -Message "A GUID was supplied as a parameter, but was not found in the input"
                    }
                }
            }
            'UseLetter'{
                
                $Letter = "$($T4WRegexDriveFormat.Match($Letter).Groups['driveletter'].Value.ToLower()):\"
                $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                $Query += " vb_label,vb_state "
                $Query += " FROM vbs "
                $Query += " WHERE LOWER(vb_letter) = '$($Letter.ToLower())'"
                break;
            }
            'UseShare'{ 

                        $Query  = " SELECT vb_id,vb_serial,vb_guid,vb_machine,vb_share,vb_letter,"
                        $Query += " vb_label,vb_state "
                        $Query += " FROM vbs "
                        $Query += " WHERE LOWER(vb_share) = '$($Share.ToLower())' AND LOWER(vb_machine) = '$($Machine.ToLower())'"
                        break;
            }
            'Default'{ 
                        $Query  = " SELECT * "
                        $Query += " FROM vbs"
                        break;
            }
        }

        Write-Verbose "`r`n######Volume Query#######`r`n$Query`r`n##########################"

        $Results = Invoke-SqliteQuery -DataSource $Database -Query $Query  
        
        return @( foreach($Result in $Results){

            [PSCustomObject]@{
                id      =   $Result.vb_id
                serial  =   $Result.vb_serial
                guid    =   $Result.vb_guid
                machine =   $Result.vb_machine
                share   =   $Result.vb_share
                letter  =   $Result.vb_letter
                label   =   $Result.vb_label
                state   =   $Result.vb_state
            }
        })
    }
    end{
        Write-Verbose "=============  Ending Get-T4WVolumes"
    }
}

Function Add-T4WVolume{
<#
.SYNOPSIS
Adds Volume to Database

.DESCRIPTION
Add Volume to the database

.PARAMETER Volume
Supply the Path to the volume.  This can be a Drive Letter, UNC NAme or VolumeGUID path

.PARAMETER Serial
Find's volume by it's serial number

.PARAMETER Database
Add-T4WVolume -Volume \\?\Volume{0a83034e-e083-43b5-8a28-9ba7cc2c7819}\

.EXAMPLE
Database Path

.NOTES
You should add the database through the normal means of publishing a tag/comment to it, but this is the programatic approach.
#>    
    [alias("Add-Volume",'Add-T4WVolumes')]
    [cmdletbinding(DefaultParameterSetName = 'VolumeName')]

    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="VolumeName")]
        [alias("VolumeName","PATH")]
        [string]$Volume,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="Serial")]
        [alias("SerialNumber")]
        [string]$Serial,



        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    begin { 


        $VolumeRecord = [PSCustomObject]@{
            serial = ''
            guid   = ""
            machine = ''
            share   = ''
            letter  = ''
            label   = ''
        }



        
        $INSERTQuery =  " INSERT INTO vbs ( vb_serial,vb_guid,vb_machine,vb_share,vb_letter,vb_label,vb_state)"
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','2');"
        
        switch($PsCmdlet.ParameterSetName){
            'VolumeName'{ 
                switch -regex ($Volume){
                    #UNC MATCH
                    $T4WRegexUNC.ToString() {
                        $RegPath = $T4WREgexUNC.Match($Volume)

                        $VolumeRecord.machine = $RegPath.Groups['server'].value
                        $VolumeRecord.share   = $RegPath.Groups['share'].Value

                        #Is the Drive Mapped?
                        $PSDrive = Get-PSDrive | Where-Object {$_.DisplayRoot -eq "\\$($VolumeRecord.machine)\$($VolumeRecord.share)"}

                        if($Null -ne $PSDrive){
                            $VolumeRecord.letter = "$($T4WRegexDriveFormat.match($PSDrive.Root).Groups['driveletter'].value):\"
                            
                            $MappedLogicalDrive = get-ciminstance -Query  "select * FROM Win32_MappedLogicalDisk WHERE ProviderName LIKE '\\\\$($VolumeRecord.machine)\\$($VolumeRecord.share)'" | Select-Object -First 1 *
                            if($Null -ne $MappedLogicalDrive){
                                $VolumeRecord.serial = $MappedLogicalDrive.VolumeSerialNumber
                            }
                        }
                        Write-Verbose "`r`n##################`r`n$($VolumeRecord|Out-string)`r`n##################`r`n"
                        $INSERT =  $INSERTQuery -f $VolumeRecord.serial, $VolumeRecord.guid,$VolumeRecord.machine,$VolumeRecord.share,$VolumeRecord.letter,$VolumeRecord.label


                        break
                    }
                    #VOLUME PATH MATCH
                    $T4WRegexVolume.ToString() {
                        
                        $VolumeRecord.guid = "$($T4WRegexVolume.match($Volume).Groups['volume'].Value)\"

                        $CimWhere = $VolumeRecord.guid -replace '\\','\\'
                        $CimVolume = Get-Ciminstance -query "Select * FROM win32_Volume WHERE DeviceID LIKE '$CimWhere'" | Select-Object *

                        #Set the Volume Letter
                        $VolumeRecord.letter = "$($CimVolume.DriveLetter)\"

                        $CIMSMBShare   = Get-CimInstance -Namespace "Root/Microsoft/Windows/SMB" -Query "SELECT * FROM MSFT_Smbshare WHERE Path='$($VolumeRecord.Letter)\'"  | Select-Object *
                        $CimLogicalDisk = Get-Ciminstance -query "Select * FROM win32_Volume WHERE Name='$($VolumeRecord.Letter)\'" | Select-Object *

                        #Supply Serial Number if Available in HEX format
                        if(("$($CimVolume.SerialNumber)").length -gt 0){
                            $VolumeRecord.Serial = "{0:X}" -f $CimVolume.SerialNumber
                        }elseif(("$($CimLogicalDisk.SerialNumber)").length -gt 0){
                            $VolumeRecord.Serial = "{0:X}" -f $CimLogicalDisk.SerialNumber
                        }

                        #Add the Share and MAchine Name if found
                        if($Null -ne $CIMSMBShare){
                            $VolumeRecord.share = $CIMSMBShare.Name
                            $VolumeRecord.machine = "${env:COMPUTERNAME}"
                        }

                        #Add the Disk Label
                        if(("$($CimVolume.Label)").length -gt 0){
                            $VolumeRecord.Label = "{0:X}" -f $CimVolume.Label
                        }elseif(("$($CimLogicalDisk.Label)").length -gt 0){
                            $VolumeRecord.Label = "{0:X}" -f $CimLogicalDisk.Label
                        }

                        Write-Verbose "`r`n##################`r`n$($VolumeRecord|Out-string)`r`n##################`r`n"

                       
                        break
                    }
                    $T4WRegexDriveFormat.ToString(){
                        
                        #Set the Volume Letter
                        $VolumeRecord.letter="$($T4wREGEXdriveFormat.Match($Volume).Groups['driveletter'].value):\"
                        
                        #Run CIM Queries
                        $CimVolume = Get-Ciminstance -query "Select * FROM win32_Volume WHERE DriveLetter LIKE '$($VolumeRecord.letter -replace "\\",'')'" | Select-Object *
                        $CIMSMBShare   = Get-CimInstance -Namespace "Root/Microsoft/Windows/SMB" -Query "SELECT * FROM MSFT_Smbshare WHERE Path='$($VolumeRecord.Letter)\'"  | Select-Object *
                        $CimLogicalDisk = Get-Ciminstance -query "Select * FROM win32_Volume WHERE Name='$($VolumeRecord.Letter)\'" | Select-Object *

                        #Supply GUID
                        if(("$($CimVolume.DeviceID)").length -gt 0){
                            $VolumeRecord.guid = $CimVolume.DeviceID
                        }elseif(("$($CimLogicalDisk.DeviceID)").length -gt 0){
                            $VolumeRecord.Serial =  $CimLogicalDisk.DeviceID
                        }

                        #Supply Serial Number if Available in HEX format
                        if(("$($CimVolume.SerialNumber)").length -gt 0){
                            $VolumeRecord.Serial = "{0:X}" -f $CimVolume.SerialNumber
                        }elseif(("$($CimLogicalDisk.SerialNumber)").length -gt 0){
                            $VolumeRecord.Serial = "{0:X}" -f $CimLogicalDisk.SerialNumber
                        }

                        #Add the Share and MAchine Name if found
                        if($Null -ne $CIMSMBShare){
                            $VolumeRecord.share = $CIMSMBShare.Name
                            $VolumeRecord.machine = "${env:COMPUTERNAME}"
                        }

                        #Add the Disk Label
                        if(("$($CimVolume.Label)").length -gt 0){
                            $VolumeRecord.Label = "{0:X}" -f $CimVolume.Label
                        }elseif(("$($CimLogicalDisk.Label)").length -gt 0){
                            $VolumeRecord.Label = "{0:X}" -f $CimLogicalDisk.Label
                        }


                        Write-Verbose "`r`n##################`r`n$($VolumeRecord|Out-string)`r`n##################`r`n"
                        break
                    }


                }

                break
            }
            'Serial' {
                try{           
                    #Unsure if this is a Hex Value or not, so we will try to convert it from Hex to an Integer
                    $WHERE = "WHERE SerialNumber LIKE '{0}'" -f [int32]"0x$Serial"
                }catch{
                    #If it failed to convert, it is unlikely a hexvalue
                    $WHERE = "WHERE SerialNumber LIKE '{0}'" -f $Serial
                }

               $VolumeData = Get-CimInstance -query  "select DeviceID,DriveLetter,Label,SerialNumber,SystemName,Caption,Name FROM win32_volume $WHERE" | Select-Object DriveLetter,Label,SerialNumber,@{Label="SerialHex";Expression={'{0:X}' -f $_.SerialNumber}}, SystemName,Caption,Name,DeviceID

               Write-Verbose "`r`n##################`r`n$( ($VolumeData | Out-String).Trim())`r`n##################"
                break
            }
        }
    }

    process {
        
        if($VolumeRecord.serial -eq '' -and $VolumeRecord.guid -eq ''  -and $VolumeRecord.machine -eq ''  -and $VolumeRecord.share -eq ''  -and $VolumeRecord.letter -eq ''  -and $VolumeRecord.label -eq '' ){
            Write-Error -erroraction Stop -Message "No Volume Data was found"
        }

        #Build our Check Query Based on what we have Collected
        [string[]]$WhereClause = @()

        
        $Query += "vb_guid = '$($VolumeRecord.guid)' "
        if("$($VolumeRecord.guid)".Length -gt 0){
            $WhereClause += "vb_guid LIKE '$($VolumeRecord.Serial)'"
        }
        if("$($VolumeRecord.Serial)".Length -gt 0){
            $WhereClause += "vb_serial LIKE '$($VolumeRecord.Serial)'"
        }
        if("$($VolumeRecord.Letter)".Length -gt 0){
            $WhereClause += "(vb_letter = '$($VolumeRecord.Letter)' AND vb_label = '$($VolumeRecord.Label)') "
        }
        if("$($VolumeRecord.Share)".Length -gt 0){
            $WhereClause += "(vb_machine = '$($VolumeRecord.machine)' AND vb_share = '$($VolumeRecord.share)') "
        }
        $Query  = " SELECT vb_id as id,vb_serial as serial,vb_guid as guid,vb_machine as machine,vb_share as share,vb_letter as letter,vb_label as label,vb_state as state FROM vbs "
        $QUERY += "WHERE $($WhereClause -join " OR ")"
        $QUERY

        $QueryResult = Invoke-SqliteQuery -DataSource $Database -Query $Query -ErrorAction STOP

        
        
        #If we found something, we exit and return what we found
        if($Null -ne $QueryResult ){
         Write-Error "Record Already Exists" -ErrorAction Continue
         return $QueryResult
        }

        $INSERT =  $INSERTQuery -f $VolumeRecord.serial, $VolumeRecord.guid,$VolumeRecord.machine,$VolumeRecord.share,$VolumeRecord.letter,$VolumeRecord.label

        Write-Verbose $INSERT

        Invoke-SqliteQuery -DataSource $Database -Query $INSERT

        #Return the created object
        return Invoke-SqliteQuery -DataSource $Database -Query $Query
        
    }
}
#endregion


#region Tags
function get-T4WTags {
    <#
    .SYNOPSIS
    Gets the Tags Available in the database
    
    .DESCRIPTION
    Gets the Tags Avilable in the database, this is a multi use function that can gather all tags, or specific tags based on parameters
    
    .PARAMETER ID
    ID of the tag you are looking for
    
    .PARAMETER GUID
    GUID of the tag
    
    .PARAMETER ParentID
    ParentID Used to find all tags with the supplied parent ID
    
    .PARAMETER Name
    Tag Name
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
get-T4WTags -id 746

id                : 746
guid              : {b544c089-0607-459a-ab8b-182a8260c933}
parent            : 652
name              : STLMaster
fullname          : STL\ARTIST\STLMaster
created           : 133884328397839041
modified          : 133884380242534239
type              : 1
subtype           : 0
icon              : 3
state             : 0
parent_timestamp  : 133884328397839041
parent_time       : 4/6/2025 1:07:19 PM
name_timestamp    : 133884328397839041
name_time         : 4/6/2025 1:07:19 PM
icon_timestamp    : 133884380242534239
icon_time         : 4/6/2025 2:33:44 PM
subtype_timestamp : 133884328397839041
subtype_tim       : 4/6/2025 1:07:19 PM
state_timestamp   : 133884328397839041
state_time        : 4/6/2025 1:07:19 PM
    
#>
    [cmdletbinding(DefaultParameterSetName = 'Default')]
    param( 
        [Parameter(Mandatory=$False, ParameterSetName="UseID")]
        [int]$ID,
        
        [Parameter(Mandatory=$False, ParameterSetName="UseGUID")]
        [string]$GUID,
        [Parameter(Mandatory=$False, ParameterSetName="UseParentID")]
        [int]$ParentID,
        [Parameter(Mandatory=$False, ParameterSetName="UseName")]
        [string]$Name,

        [Parameter(Mandatory=$False, ParameterSetName="UseGUID")]
        [Parameter(Mandatory=$False, ParameterSetName="Default")]
        [Parameter(Mandatory=$False, ParameterSetName="UseName")]
        [Parameter(Mandatory=$False, ParameterSetName="UseID")]
        [Parameter(Mandatory=$False, ParameterSetName="UseParentID")]
        [string]$Database=$TagDBPath)
    begin{

        
    }
    process {


        switch($PsCmdlet.ParameterSetName){
            'UseID'{ 
                        $Query  = " SELECT to_id,to_guid,to_parent,to_name,to_created,to_modified,to_type,to_subtype,"
                        $Query += " to_icon,to_state,to_ts_parent,to_ts_name,to_ts_icon,to_ts_subtype,to_ts_state "
                        $Query += " FROM tos"
                        $Query += " WHERE to_id = '$ID'"

                        break;
            }
            'UseGUID'{ 
                        $Query  = " SELECT to_id,to_guid,to_parent,to_name,to_created,to_modified,to_type,to_subtype,"
                        $Query += " to_icon,to_state,to_ts_parent,to_ts_name,to_ts_icon,to_ts_subtype,to_ts_state "
                        $Query += " FROM tos"
                        $Query += " WHERE to_guid = '$GUID'"

                        break;
            }
            'UseParentID'{ 
                        $Query  = " SELECT to_id,to_guid,to_parent,to_name,to_created,to_modified,to_type,to_subtype,"
                        $Query += " to_icon,to_state,to_ts_parent,to_ts_name,to_ts_icon,to_ts_subtype,to_ts_state "
                        $Query += " FROM tos "
                        $Query += " WHERE to_parent = '$ParentID'"

                        break;
            }
            'UseName'{ 
                        $Name = $Name -replace "'","''"
                        $Query  = " SELECT to_id,to_guid,to_parent,to_name,to_created,to_modified,to_type,to_subtype,"
                        $Query += " to_icon,to_state,to_ts_parent,to_ts_name,to_ts_icon,to_ts_subtype,to_ts_state "
                        $Query += " FROM tos "
                        $Query += " WHERE to_name LIKE '%$Name%'"
                        break;
            }
            'Default'{ 
                        $Query  = " SELECT * "
                        $Query += " FROM tos"
                        break;
            }
        }

        $AllTags = Invoke-SqliteQuery -DataSource $Database -Query "Select to_id,to_parent,to_name FROM tos"

        $Results = Invoke-SqliteQuery -DataSource $Database -Query $Query  



        return @( foreach($Result in $Results){
            $Tagfullname= New-Object System.Collections.Stack
            $TargetTag=$Result.to_id
            While($TargetTag -gt 0){
                $TAG = $AllTags | Where-Object {$_.to_id -eq $TargetTag}
                $TargetTag = $TAG.to_parent
                $Tagfullname.push($TAG.to_name)
            }        
            [PSCustomObject]@{
                id         = $Result.to_id
                guid       = $Result.to_guid
                parent     = $Result.to_parent
                name       = $Result.to_name
                fullname   = @($Tagfullname.ToArray()) -join "\"
                created    = $Result.to_created
                modified   = $Result.to_modified
                type       = $Result.to_type
                subtype    = $Result.to_subtype
                icon       = $Result.to_icon
                state      = $Result.to_state
                parent_timestamp  = $Result.to_ts_parent
                parent_time       = $(try{[datetime]::FromFileTime($Result.to_ts_parent)}catch{""})
                name_timestamp    = $Result.to_ts_name
                name_time         = $(try{[datetime]::FromFileTime($Result.to_ts_name)}catch{""})
                icon_timestamp    = $Result.to_ts_icon
                icon_time         = $(try{[datetime]::FromFileTime($Result.to_ts_icon)}catch{""})
                subtype_timestamp = $Result.to_ts_subtype
                subtype_tim       = $(try{[datetime]::FromFileTime($Result.to_ts_subtype)}catch{""})
                state_timestamp   = $Result.to_ts_state
                state_time        = $(try{[datetime]::FromFileTime($Result.to_ts_state)}catch{""})
            }
        })
    }
}

Function Add-T4WTag{
    <#
    .SYNOPSIS
    Adds Tag to the database
    
    .DESCRIPTION
    Adds Tag to the database
    
    .PARAMETER ParentID
    ParentID for the tag, Must be a valid tagid
    
    .PARAMETER Name
    Name of the Tag
    
    .PARAMETER Type
    Type of Tag, Either Tag or TagGroup
    
    .PARAMETER IconID
    IconID
    
    .PARAMETER IconShape
    Shape for Icon
    
    .PARAMETER IconColor
    Color for Icon
    
    .PARAMETER State
    State,  Should be 1 for TagGroups and 0 for Tags to be enabled
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Add-T4WTag -parentid 369 -name "Purple Heart Tag" -Type Tag -IconShape Heart -IconColor Purple -State Enable

    #>
    [cmdletbinding(DefaultParameterSetName = 'UseShapeInfo',SupportsShouldProcess=$true)]
    param(
        [Parameter(Mandatory=$True)]
        [int]$ParentID,


        [Parameter(Mandatory=$True)]
        
        [string]$Name,

        [Parameter(Mandatory=$True)]
        [T4W_TagType]$Type,

        [Parameter(Mandatory=$True,ParameterSetName="UseIconID")]
        [int]$IconID,
        
        [Parameter(Mandatory=$True,ParameterSetName="UseShapeInfo")]
        [T4W_IconShape]$IconShape,
        
        [Parameter(Mandatory=$True,ParameterSetName="UseShapeInfo")]
        [T4W_IconColor]$IconColor,
        
        [Parameter(Mandatory=$False)]
        [T4W_TagGroupState]$State=[T4W_TagGroupState]::Enable ,

        
        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    begin {

        $Name=$Name -replace "'","''"  #Need to make sure any single quotes are escaped

        
        
        if($Type -eq [T4W_TagType]::TagGroup   ){
            $StateID = [int]$State
            if($ParentID -ne -2){
                Write-Verbose "Type is TagGroup, Changing the ParentId to -2"
                $ParentID = -2
            }
        }else{
            $StateID = 0
        }

        if($Null -eq $(get-t4wTags -ID $ParentID)){
            Write-Error -ErrorAction Stop -Message "Supplied ParentId is invalid"
        }

        $TimeStamp = (get-date).ToFileTime()

        switch($PsCmdlet.ParameterSetName){
            'UseShapeInfo'{ $IconID = [int]$IconShape+[int]$IconColor ;break}
        }
    }
    process {
        $GUID = Create-T4WGUID -Type Tag

        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:"
        Write-Verbose $UpdateInfoQuery

        $InsertQuery = "INSERT INTO tos ( to_guid,to_parent,to_name,to_created,to_modified,to_type,to_subtype,to_icon,to_state,to_ts_parent,to_ts_name,to_ts_icon,to_ts_subtype,to_ts_state)
                VALUES ('{0}',{1},'{2}','{3}','{4}','{5}','{6}','{7}','{8}','{9}','{10}','{11}','{12}','{13}');" -f
                    "{$Guid}",$ParentID,$Name,$TimeStamp,$TimeStamp,[int]$Type,"0",$IconID,$StateID,$TimeStamp,$TimeStamp,$TimeStamp,$TimeStamp,$TimeStamp
        Write-Verbose "INSERT Query: $InsertQuery"


        write-verbose "TOS Table Update Command:"
        Write-Verbose $InsertQuery

        Write-Verbose "Database: $Database"

        if($PSCmdlet.ShouldProcess("insert record", "Invoke-SQLLiteQuery")){
            Invoke-SqliteQuery -DataSource $Database -query $InsertQuery -ErrorAction STOP
        
            Invoke-SqliteQuery -DataSource $Database -query $UpdateInfoQuery 

        }else{
            Write-Host "Would Run SQL Query: $InsertQuery"
            Write-Host "Would Run SQL Query: $UpdateInfoQuery"
        }

        return $(Get-t4wtags -GUID $GUID)
         
    }
}

Function Update-T4WTag{
    <#
    .SYNOPSIS
    Updates the Tag
    
    .DESCRIPTION
    Updates either the ParentID, Name, State or Icon of the tag
    
    .PARAMETER ID
    ID of the tag
    
    .PARAMETER ParentID
    New ParentId
    
    .PARAMETER Name
    New Name
    
    .PARAMETER IconID
    New IconID
    
    .PARAMETER IconShape
    New Icon Shape
    
    .PARAMETER IconColor
    New Icon Color
    
    .PARAMETER TagGroupState
    New Tag Group State
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Update-T4WTag -id 771 -TagGroupState Enable

    .EXAMPLE
    Update-T4WTag -id 771 -TagGroupState Enable -Name "NewTagname"-Verbose
    
    .NOTES
    General notes
    #>
    [cmdletbinding(DefaultParameterSetName = 'UseIconID')]
    param(

        [Parameter(Mandatory=$True)]
        [string]$ID,

        
        [Parameter(Mandatory=$False)]
        [int]$ParentID,

        [Parameter(Mandatory=$False)]
        [string]$Name,


        [Parameter(Mandatory=$False,ParameterSetName="UseIconID")]
        [int]$IconID="-1",
        
        [Parameter(Mandatory=$True,ParameterSetName="UseShapeInfo")]
        [T4W_IconShape]$IconShape,
        
        [Parameter(Mandatory=$True,ParameterSetName="UseShapeInfo")]
        [T4W_IconColor]$IconColor,
        
        [Parameter(Mandatory=$False)]
        [T4W_TagGroupState]$TagGroupState,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    begin {

        $Tag = get-T4WTags -ID $ID
        if($Null -eq $TAG){
            Write-Error -ErrorAction Stop -Message "Tag with id: $ID , was not found in the database"
        }

        [int64]$TimeStamp = (get-date).ToFileTime()

        [string[]]$SETCommands = @()

        #Add Set Commands if we are doing updates to the Name
        if($Null -ne $Name -and "$Name".Length -gt 0 ){
            $Name=$Name -replace "'","''"  #Need to make sure any single quotes are escaped
            Write-Verbose "Adding Name=$Name"
            $SETCommands += "to_name = '$Name'"
            $SETCommands += "to_ts_name = '$TimeStamp'"
            $SETCommands += "to_modified = '$TimeStamp'"
        }

        #Add Set Commands if we are doing updates to the ParentID
        if($Null -ne $ParentID -and ($ParentID -in (-2,-3) -or $ParentID -gt 0  )){
            $ParentTag = Get-T4WTags -id $ParentID
            if($Null -eq $ParentTag){ Write-Error -ErrorAction Stop -Message "Parent Tag ID: $ParentID , was not found in the database" }

            Write-Verbose "Adding ParentID=$ParentID"
            $SETCommands +="to_parent = '$ParentID'"
            $SETCommands += "to_ts_parent = '$TimeStamp'"
            $SETCommands += "to_modified = '$TimeStamp'"
        }

        #Add Set Commands if we are doing updates to the Icon
        switch($PsCmdlet.ParameterSetName){
            'UseShapeInfo'{ 
                
                Write-Verbose "Adding IconID='$([int]$IconShape+[int]$IconColor)' = '$([int]$IconShape)'+'$([int]$IconColor)'"
                $IconID = [int]$IconShape+[int]$IconColor ;
                $SETCommands +="to_icon = '$IconID'"
                $SETCommands += "to_ts_icon = '$TimeStamp'"
                $SETCommands += "to_modified = '$TimeStamp'"
                break
            }
            'UseIconID' {
                if($IconID -ge 0){
                    Write-Verbose "Adding IconID=$IconID"
                    $SETCommands +="to_icon = '$IconID'"
                    $SETCommands += "to_ts_icon = '$TimeStamp'"
                    $SETCommands += "to_modified = '$TimeStamp'"
                }
                break
            }
        }

        
        if($Null -ne $TagGroupState -and $TAG.type -eq [int]([T4W_TagType]::TagGroup)){
            write-verbose "Modifying the State of the Tag, TagGroup='$TagGroup' "
            $SETCommands +="to_state = '$([int]$TagGroupState)'"
            $SETCommands += "to_ts_state = '$TimeStamp'"
            $SETCommands += "to_modified = '$TimeStamp'"
        }else{
            Write-Verbose "Ignoring TagGroup State, item is not a TagGroup "
        }

        
        $SETCommands = $SETCommands | Select-Object -Unique
    }
    process {
        if($SETCommands.Count -gt 0 ){

            $UpdateInfoQuery = "Update info`r`n  SET info_value_i = '$TimeStamp'`r`n  WHERE info_key = '-7001';"
            write-verbose "Info Table Update Command:"
            Write-Verbose $UpdateInfoQuery

            $UpdateTOS  =  " UPDATE tos "
            $UpdateTOS  += "`r`n SET "
            $UpdateTOS  += $SETCommands -join ",`r`n     "
            $UpdateTOS  += "`r`n WHERE to_id = '$ID'"

            write-verbose "Tos Table Update Command:"
            Write-Verbose $UpdateTOS

            Invoke-SqliteQuery -DataSource $TagDBPath -Query $UpdateInfoQuery
            Invoke-SqliteQuery -DataSource $TagDBPath -Query $UpdateTOS
        }else{
            Write-host "No Changes were specified"
        }
    }
}
#endregion


#region Folder
Function Get-T4WFolder{
    <#
    .SYNOPSIS
    Get List of Folders in the Database
    
    .DESCRIPTION
    Can be used to Get specific folders
    
    .PARAMETER FolderPath
    Path to the Folder, this must be a full path and in the form of Volume, Drive or UNC Path
    
    .PARAMETER FolderID
    The Folder ID
    
    .PARAMETER FolderGuid
    Folder Guid
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Get-T4WFolder 'G:\SFTP_Root\'   

id               : 8
guid             : {35657514-80b7-4d20-9cbe-2250f47d29c8}
path             : \SFTP_Root\
volume           : 5
state            : 0
path_timestamp   : 133893145186016985
path_time        : 4/16/2025 6:01:58 PM
volume_timestamp : 133893145186016985
volume_time      : 4/16/2025 6:01:58 PM
state_timestamp  : 133893012561961217
state_time       : 4/16/2025 2:20:56 PM


    .EXAMPLE
Get-T4WFolder -GUID '{35657514-80b7-4d20-9cbe-2250f47d29c8}'

id               : 8
guid             : {35657514-80b7-4d20-9cbe-2250f47d29c8}
path             : \SFTP_Root\
volume           : 5
state            : 0
path_timestamp   : 133893145186016985
path_time        : 4/16/2025 6:01:58 PM
volume_timestamp : 133893145186016985
volume_time      : 4/16/2025 6:01:58 PM
state_timestamp  : 133893012561961217
state_time       : 4/16/2025 2:20:56 PM
    
    .NOTES
    CANNOT be used to get ALL folders
    #>
    [cmdletbinding(DefaultParameterSetName = 'FolderPath')]
    param( 
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath",Position=0)]
        [alias("FolderName","Name","Fullname","Path","FilePath")]
        [string]$FolderPath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderID")]
        [alias("ID")]
        [int]$FolderID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderGUID")]
        [alias("GUID")]
        [GUID]$FolderGuid,

        [Parameter(Mandatory=$False,Position=1)]
        [string]$Database=$TagDBPath
    )


    begin{
        $FolderPath = $FolderPath -replace [regex]::Escape("Microsoft.PowerShell.Core\FileSystem::"),""
    }
    process {
        
        switch($PsCmdlet.ParameterSetName){
            'FolderPath'{ 
                        
                $T4WVolume=$null
                $RelativeFolderPath = $Null
                $FolderPath = $FolderPath -replace "'","''"  #Lets escape any single quotes in case there are any

                #Determine what Drive\Share\Volume format the path is in
                switch -Regex ($FolderPath){
                    #UNC Format  (Network path)
                    $T4WRegexUNC.ToString() {
                        $Folder = $T4WRegexUNC.Match($FolderPath) 
                        $Machine=$Folder.Groups["server"].Value.toLower()                
                        $Share=$Folder.Groups["share"].Value.toLower()
                        $RelativeFolderPath=$Folder.Groups["path"].Value.toLower()

                        Write-Verbose "Machine: $Machine"
                        Write-Verbose "Share: $Share"

                        $T4WVolume = Get-T4WVolumes -Share $Share -Machine $Machine

                        if($Null -eq $T4WVolume){ Write-Error -ErrorAction Stop -Message "A Volume matching the following could not be found `r`nMachine: $Machine  `r`nShare: $Share `r`n\\$machine\$Share" }
                        Write-Verbose "`r`n######Volume Record#######`r`n$(($T4WVolume | Out-String).Trim())`r`n##########################"

                        break
                    }

                    #Drive Letter Format
                    $T4WRegexDriveFormat.ToString() {
                        $Folder = $T4WRegexDriveLetter.Match($FolderPath)
                        $DriveLetter=$Folder.Groups["driveletter"].Value.toLower()                
                        $RelativeFolderPath=$Folder.Groups["path"].Value.toLower()

                        Write-Verbose "Drive: $DriveLetter"
                        Write-Verbose "RelativeFolderPath: $RelativeFolderPath"

                        $PSDrive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue  | Select-Object -First 1
                        
                        #Lets stop in our tracks if the drive letter isn't active.
                        if($Null -eq $PSDrive){Write-Error -ErrorAction STOP -Message "This is not an active Drive Letter"  }


                        #Mapped Drive Letter to UNC Path
                        switch -Regex ($PSdrive.DisplayRoot){
                            $T4WRegexUNC.ToString() {
                                Write-Verbose "This is a remote Path"

                                $RemoteFolder = $T4WRegexUNC.Match($PSDrive.DisplayRoot)
                                $Machine=$RemoteFolder.Groups["server"].Value.toLower()                
                                $Share=$RemoteFolder.Groups["share"].Value.toLower()

                                $T4WVolume = Get-T4WVolumes -Share $Share -Machine $Machine
                                if($Null -eq $T4WVolume){ Write-Error -ErrorAction Stop -Message "A Volume matching the following could not be found `r`nMachine: $Machine  `r`nShare: $Share `r`n\\$machine\$Share" }
                                
                                break
                            }

                            #Local Drive
                            default{
                                Write-Verbose "This is a Local Path"
                                $VolumeUniqueID = Get-Volume -DriveLetter $DriveLetter -ErrorAction SilentlyContinue | Select-Object UniqueId
                                if($Null -ne $VolumeUniqueID){
                                    #Get the volume from the database by VolumeID
                                    $T4WVolume = Get-T4WVolumes -GUID $VolumeUniqueID.UniqueId

                                    if($Null -eq $T4WVolume){
                                        Write-Verbose "Failed to find the volume by the GUID, Trying by Drive Letter"
                                        #Get the volume from the database by Drive Letter
                                        $T4WVolume = Get-T4WVolumes -Letter $DriveLetter 
                                    }
                                }else{
                                    #Get the volume from the database by Drive Letter
                                    $T4WVolume = Get-T4WVolumes -Letter $DriveLetter

                                }
                                break
                            }
                        }
                        break
                    }
                    #VOLUMEID format
                    $T4WRegexVolume.ToString() {
                        $VolumeREGEX =  $T4WRegexVolume.match($FolderPath)
                        $VolumeUniqueID = $VolumeREGEX.Groups['volume'].value

                    
                        $RelativeFolderPath = "\$($VolumeREGEX.Groups["path"].Value)"
                        $RelativeFolderPath = $RelativeFolderPath -replace "[\\]+","\"
                        Write-Verbose "VolumeID: $VolumeUniqueID"
                        Write-Verbose "Relative Folder: $RelativeFolderPath"

                        $T4WVolume = Get-T4WVolumes -GUID $VolumeUniqueID
                        break
                    }


                }

                if($RelativeFolderPath -notmatch "\\$"){Write-Error -ErrorAction Stop  -Message "FolderPaths must end in a '\'"}
                if($Null -eq $T4WVolume){Write-Error -ErrorAction Stop -Message "Failed to determine the Drive Volume from database"}

                if("$($VolumeUniqueID)".Length -gt 0){ Write-Verbose "VolumeID: $VolumeUniqueID" }
                if("$($Machine)".Length -gt 0){ Write-Verbose "Machine: $Machine" }
                if("$($Share)".Length -gt 0){ Write-Verbose "Share: $Share" }
                if("$($DriveLetter)".Length -gt 0){ Write-Verbose "Drive Letter: $DriveLetter" }

                Write-Verbose "RelativeFolderPath: $RelativeFolderPath"
                Write-Verbose "`r`n######Volume Record#######`r`n$(($T4WVolume | Out-String).Trim())`r`n##########################"

                #Build the Query with the Volume data we've collected.
                #NOTE the fo_path is intentially a LIKE statement because this is not case sensitive.
                $FolderQuery  =  " SELECT fo_id, fo_guid, fo_path, fo_volume, fo_state, fo_ts_path, fo_ts_volume, fo_ts_state "
                $FolderQuery  += " FROM fos"
                $FolderQuery  += " WHERE fo_volume = '$($T4WVolume.id)' AND fo_path LIKE '$($RelativeFolderPath)' "
                $FolderQuery  += " LIMIT 1;"

                break
            }
            'FolderID'{
                $FolderQuery  =  " SELECT fo_id, fo_guid, fo_path, fo_volume, fo_state, fo_ts_path, fo_ts_volume, fo_ts_state "
                $FolderQuery  += " FROM fos"
                $FolderQuery  += " WHERE fo_id = '$($FolderID)' "
                $FolderQuery  += " LIMIT 1;"
                break
            }
            'FolderGUID'{
                $FolderQuery  =  " SELECT fo_id, fo_guid, fo_path, fo_volume, fo_state, fo_ts_path, fo_ts_volume, fo_ts_state "
                $FolderQuery  += " FROM fos"
                $FolderQuery  += " WHERE fo_guid = '{$($FolderGUID.Guid)}' "
                $FolderQuery  += " LIMIT 1;"

                break
            }
            

        }

        #Execute the Database Query
        $Result = Invoke-SqliteQuery -DataSource $Database -Query $FolderQuery -ErrorAction Stop

        #Return a PSCustom Object of the Database Record
        if($Null -ne $Result ){
            return [PSCustomObject]@{
                id = $Result.fo_id
                guid = $Result.fo_guid
                path = $Result.fo_path
                volume = $Result.fo_volume
                state = $Result.fo_state
                path_timestamp = $Result.fo_ts_path
                path_time        = $(try{[datetime]::FromFileTime($Result.fo_ts_path)}catch{""})
                volume_timestamp = $Result.fo_ts_volume
                volume_time        = $(try{[datetime]::FromFileTime($Result.fo_ts_volume)}catch{""})
                state_timestamp = $Result.fo_ts_state
                state_time        = $(try{[datetime]::FromFileTime($Result.fo_ts_state)}catch{""})

            }
        }else{
            return $Null
        }
    }
}

Function Add-T4WFolder{
<#
.SYNOPSIS
Adds Folder to the Database

.DESCRIPTION
Adds Folder to the Database, and returns the created record.

.PARAMETER FolderPath
The Path to be created

.PARAMETER Database
Database Path

.EXAMPLE
Get-T4WFolder -folderpath G:\stl\

id               : 1
guid             : {2b0b6fde-fae0-4450-8df8-023bcbb5ff02}
path             : \STL\
volume           : 5
state            : 0
path_timestamp   : 133887003860014039
path_time        : 4/9/2025 3:26:26 PM
volume_timestamp : 133887003860014039
volume_time      : 4/9/2025 3:26:26 PM
state_timestamp  : 133887003860014039
state_time       : 4/9/2025 3:26:26 PM

.NOTES
General notes
#>
    [alias("Add-Folder")]
    [cmdletbinding(DefaultParameterSetName = 'UseFolderName')]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath")]
        [alias("FolderName","Name","Fullname","Path","FilePath")]
        [string]$FolderPath,


        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath



    )
    begin { 

        #CHECK the Guid for a collision
        $GUID = Create-T4WGUID -Type Folder


        [int64]$TimeStamp = (get-date).ToFileTime()
        $FileRecord = [PSCustomObject]@{
            guid = "{$GUID}"
            path = ''
            volumeid = -1
            state = 0
            path_time = $TimeStamp
            volume_time = $TimeStamp
            state_time = $TimeStamp
        }

        switch -Regex ($Folderpath){
            $T4WRegexVolume.ToString(){
                $FolderRegexResult = $T4WRegexVolume.Match($FolderPath)
                $Volume = Get-T4WVolumes -GUID   $FolderRegexResult.Groups['volume'].Value
                if($Null -eq $Volume){
                    Write-Error -ErrorAction Stop -Message "Volume is not present in database, Create the volume record before proceeding"
                }
                $FileRecord.volumeid = $Volume.id
                $FileRecord.path=$FolderRegexResult.Groups['path'].Value
                break
            }
            $T4WRegexUNC.ToString(){
                $FolderRegexResult = $T4WRegexUNC.Match($FolderPath)
                $Volume = Get-T4WVolumes -Machine   $FolderRegexResult.Groups['server'].Value -Share $FolderRegexResult.Groups['share'].Value
                if($Null -eq $Volume){
                    Write-Error -ErrorAction Stop -Message "Volume is not present in database, Create the volume record before proceeding"
                }
                $FileRecord.volumeid = $Volume.id
                $FileRecord.path=$FolderRegexResult.Groups['path'].Value

                break
            }
            $T4WRegexDriveLetter.ToString(){
                $FolderRegexResult = $T4WRegexDriveLetter.Match($FolderPath)
                $Volume = Get-T4WVolumes -Letter   $FolderRegexResult.Groups['driveletter'].Value
                if($Null -eq $Volume){
                    Write-Error -ErrorAction Stop -Message "Volume is not present in database, Create the volume record before proceeding"
                }
                $FileRecord.volumeid = $Volume.id
                $FileRecord.path=$FolderRegexResult.Groups['path'].Value
                break
            }
        }
    }

    process {

        if($FileRecord.path -notmatch "\\$"){
            Write-Error "Folder must end in a trailing ""\""  `r`n`tPATH presented was ""$($FileRecord.path)""" -ErrorAction Stop
        }

        $UpdateInfoQuery = "Update info SET info_value_i = '$($FileRecord.fo_ts_path)' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:"
        Write-Verbose $UpdateInfoQuery

        $INSERTQuery =  " INSERT INTO fos (fo_guid,fo_path,fo_volume,fo_state,fo_ts_path,fo_ts_volume,fo_ts_state) "
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}'); "

        $InsertStatement = $INSERTQuery -f $FileREcord.guid,$FileRecord.path,$FileRecord.volumeid,$FileRecord.state,$FileRecord.path_time,$FileRecord.volume_time,$FileRecord.state_time

        $CheckQuery = "SELECT * FROM fos WHERE fo_volume ='$($FileRecord.volumeid)' AND LOWER(fo_path) ='$($FileREcord.path.ToLower())'"

        $CheckResults = Invoke-SqliteQuery -DataSource $Database -Query $CheckQuery
        if($Null -ne $CheckResults){
            Write-Warning  -Message "This Record already Exists"
            return $CheckResults
        }

        Write-Verbose "Preparing to Insert Record"
        Write-Verbose "INSERT Query: $InsertStatement"

        Invoke-SqliteQuery -DataSource $Database -Query $InsertStatement -ErrorAction Stop
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

        return  Get-T4WFolder -FolderPath $FolderPath
    }
}

Function Update-T4WFolder {
    <#
    .SYNOPSIS
    Allows the updating of folder items.
    
    .DESCRIPTION
    Lets you change the volume and relative path of a file, can be used when moving files outside of windows explorer.
    
    .PARAMETER ID
    Folder ID  (Required)
    
    .PARAMETER NewRelativePath
    New Relative Path
    
    .PARAMETER NewParent
    New Volume ID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Update-T4WFolder -ID 8 -NewParent 5 -NewRelativePath '\SFTP_Root\' 
    

    #>
    [cmdletbinding()]
    param(

        [Parameter(Mandatory=$True)]
        [string]$ID,

        [Parameter(Mandatory=$False)]
        [ValidatePattern("[^<>:""|?*//`t`r`n`e]")]
        [string]$NewRelativePath="",

        [Parameter(Mandatory=$False)]
        [int]$NewParent=-99,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath


    )
    begin {

        [int64]$Timestamp = (get-date).ToFileTime()


    }
    process {
        $Folder = Get-T4WFolder -id $ID
        if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "Invalid FolderID"}

        [string[]] $SQLUpdates = @()
        if($NewParent -ne -99){
            $ParentRecord=Get-T4WVolumes -id $NewParent
            if($Null -eq $ParentRecord){ Write-Error -ErrorAction Stop -Message "Invalid New Parent ID" }
            $SQLUpdates += " fo_volume = '{0}' " -f $ParentRecord.id
            $SQLUpdates += " fo_ts_volume = '{0}' " -f $Timestamp
            Write-Verbose "Volume Changing to $($ParentRecord.id)"
            Write-Verbose "####################`r`n$($ParentREcord|Out-string)`r`n####################"

        }

        if($T4WRegexDriveLetter.match( $NewRelativePath).success){
            Write-Error -ErrorAction Stop -Message "New Relative Path should not be a full path.  It must not start with a drive letter and must begin with a '\'"  

        }elseif($T4WRegexUNC.match( $NewRelativePath).success){
            Write-Error -ErrorAction Stop -Message "New Relative Path should not be a UNC path.  It must not start with a UNC Server\Share Reference and must begin with a '\'"  

        }elseif($T4WRegexVolume.match( $NewRelativePath).success){
            Write-Error -ErrorAction Stop -Message "New Relative Path should not be a full path.  It must not start with a Volume Mapping Reference and must begin with a '\'"  
        }elseif([regex]::match( $NewRelativePath,"^\\.*[^\\]$").success){
            Write-Error -ErrorAction Stop -Message "New Relative Path MUST have a trailing '\', This indicates it is a File and not a Folder."  
        }elseif([regex]::match( $NewRelativePath,"^\\.*[\\]$").success){
            $SQLUpdates += " fo_path = '{0}' " -f $($NewRelativePath -replace "'","''")
            $SQLUpdates += " fo_ts_path = '{0}' " -f $Timestamp
            Write-Verbose "Updating Path"
        }    

        $SETCommands = $SQLUpdates | Select-Object -Unique
        
        if($SETCommands.Count -eq 0){
            write-warning "No Valid Changes were provided"
            return  $Null
        }
        
        $UpdateInfoQuery = "Update info`r`n  SET info_value_i = '$Timestamp'`r`n  WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        $UpdateQuery = "UPDATE fos SET $( $SETCommands -join " , ") WHERE fo_id = '$($Folder.id)'"
        write-verbose "UPDATE QUERY: $UpdateQuery"

        Invoke-SqliteQuery -DataSource $Database -Query $UpdateQuery -ErrorAction Stop
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

        return Get-T4WFolder -FolderID $Folder.ID

    }
}

function Remove-T4WFolder{
    <#
    .SYNOPSIS
    Removes the Folder from Databse
    
    .DESCRIPTION
    Copies Duplicate to Archive Table, then deletes original
    
    .PARAMETER FolderPath
    Folder Path
    
    .PARAMETER FolderID
    Folder ID
    
    .PARAMETER FolderGuid
    Folder GUID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Remove-T4WFolder -FolderPath 'G:\SFTP_Root\t1\'
    
    .NOTES
    This will also delete any FolderComments and FolderTags associated with this record.
    #>
    [cmdletbinding()]
    param(
        
    [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath",Position=0)]
    [alias("FolderName","Name","Fullname","Path","FilePath")]
    [string]$FolderPath,

    [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderID")]
    [alias("ID")]
    [int]$FolderID,

    [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderGUID")]
    [alias("GUID")]
    [GUID]$FolderGuid,

    [Parameter(Mandatory=$False,Position=1)]
    [string]$Database=$TagDBPath
    )
    process{
        #Get the Timestamp fir updating records
        [int64]$TimeStamp = (get-date).ToFileTime()
        
        switch($PsCmdlet.ParameterSetName){
            'FolderPath'{ 
                $Folder = Get-T4WFolder -Path $FolderPath
                if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "FolderPath is not present in the database."}
                break
            }
            'FolderGUID'{ 
                $Folder = Get-T4WFolder -FolderGuid $FolderGuid
                if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "FolderGuid is not present in the database."}
                break
            }
            'FolderID'{ 
                $Folder = Get-T4WFolder -FolderID $FolderID
                if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "FolderID is not present in the database."}
                break
            }
        }

        $Volume = Get-T4WVolumes -ID $Folder.volume

        
        #We need to clean up Tags so they aren't orphaned
        foreach($Tag in (Get-T4WFolderTag -FolderID $Folder.id)){
            Remove-T4WFolderTag -GUID $Tag.guid 
        }

        #We need to clean up Comments so they aren't orphaned
        foreach($Comment in (Get-T4WFolderComments -FolderID $Folder.id)){
            #Remove-T4WFileComment -CommentID $Comment.id
            Remove-T4WFolderComment -ID $Comment.id
            
        }
        $INSERT = " INSERT INTO dfos (dfo_guid,dfo_path,dfo_vb_serial,dfo_vb_guid,dfo_vb_letter,dfo_state,dfo_ts) "
        $INSERT +=" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ;" -f `
                    $Folder.guid,$Folder.path,$Volume.serial,$volume.guid,$Volume.letter,$Folder.state,$Timestamp
        
        Write-Verbose "INSERT:  $INSERT"

        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"
        

        $DeleteQuery = "DELETE FROM fos WHERE fo_id ='{0}'" -f $Folder.id
        write-verbose "Delete Command:  $UpdateInfoQuery"


        Invoke-SQLiteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SQLiteQuery -DataSource $Database -Query $DeleteQuery -ErrorAction Continue
        Invoke-SQLiteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction SilentlyContinue


    }
}

#endregion


#region File


Function Get-T4WFile{
    <#
    .SYNOPSIS
    Gets File Records
    
    .DESCRIPTION
    Collects the file database records.
    
    .PARAMETER FilePath
    Filepath of the file
    
    .PARAMETER FileID
    File ID
    
    .PARAMETER FileGuid
    File GUID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Get-T4WFile -id 4761

id               : 4761
guid             : {d305b8ac-4735-46fc-b4b2-327322030ea9}
path             : \SQL-after-delete-tag.html
volume           : 5
state            : 0
path_timestamp   : 133893222953480194
path_time        : 4/16/2025 8:11:35 PM
volume_timestamp : 133893222953480194
volume_time      : 4/16/2025 8:11:35 PM
state_timestamp  : 133893222953480194
state_time       : 4/16/2025 8:11:35 PM

    .EXAMPLE
Get-T4WFile -GUID '{d305b8ac-4735-46fc-b4b2-327322030ea9}'

id               : 4761
guid             : {d305b8ac-4735-46fc-b4b2-327322030ea9}
path             : \SQL-after-delete-tag.html
volume           : 5
state            : 0
path_timestamp   : 133893222953480194
path_time        : 4/16/2025 8:11:35 PM
volume_timestamp : 133893222953480194
volume_time      : 4/16/2025 8:11:35 PM
state_timestamp  : 133893222953480194
state_time       : 4/16/2025 8:11:35 PM
    
    #>
    [cmdletbinding(DefaultParameterSetName = 'FilePath')]
    param( 
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath",Position=0)]
        [alias("FileName","Name","Fullname","Path")]
        [string]$FilePath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileID")]
        [alias("ID")]
        [int]$FileID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileGUID")]
        [alias("GUID")]
        [GUID]$FileGuid,

        [Parameter(Mandatory=$False,Position=1)]
        [string]$Database=$TagDBPath
    )


    begin{
        $FilePath = $FilePath -replace [regex]::Escape("Microsoft.PowerShell.Core\FileSystem::"),""
    }
    process {
        
        switch($PsCmdlet.ParameterSetName){
            'FilePath'{ 
                        
                $T4WVolume=$null
                $RelativeFilePath = $Null
                $FilePath = $FilePath -replace "'","''"  #Lets escape any single quotes in case there are any

                #Determine what Drive\Share\Volume format the path is in
                switch -Regex ($FilePath){
                    #UNC Format  (Network path)
                    $T4WRegexUNC.ToString() {
                        $File = $T4WRegexUNC.Match($FilePath) 
                        $Machine=$File.Groups["server"].Value.toLower()                
                        $Share=$File.Groups["share"].Value.toLower()
                        $RelativeFilePath=$File.Groups["path"].Value.toLower()

                        Write-Verbose "Machine: $Machine"
                        Write-Verbose "Share: $Share"

                        $T4WVolume = Get-T4WVolumes -Share $Share -Machine $Machine

                        if($Null -eq $T4WVolume){ Write-Error -ErrorAction Stop -Message "A Volume matching the following could not be found `r`nMachine: $Machine  `r`nShare: $Share `r`n\\$machine\$Share" }
                        Write-Verbose "`r`n######Volume Record#######`r`n$(($T4WVolume | Out-String).Trim())`r`n##########################"

                        break
                    }

                    #Drive Letter Format
                    $T4WRegexDriveFormat.ToString() {
                        $File = $T4WRegexDriveLetter.Match($FilePath)
                        $DriveLetter=$File.Groups["driveletter"].Value.toLower()                
                        $RelativeFilePath=$File.Groups["path"].Value.toLower()

                        Write-Verbose "Drive: $DriveLetter"
                        Write-Verbose "RelativeFilePath: $RelativeFilePath"

                        $PSDrive = Get-PSDrive -Name $DriveLetter -ErrorAction SilentlyContinue  | Select-Object -First 1
                        
                        #Lets stop in our tracks if the drive letter isn't active.
                        if($Null -eq $PSDrive){Write-Error -ErrorAction STOP -Message "This is not an active Drive Letter"  }


                        #Mapped Drive Letter to UNC Path
                        switch -Regex ($PSdrive.DisplayRoot){
                            $T4WRegexUNC.ToString() {
                                Write-Verbose "This is a remote Path"

                                $RemoteFile = $T4WRegexUNC.Match($PSDrive.DisplayRoot)
                                $Machine=$RemoteFile.Groups["server"].Value.toLower()                
                                $Share=$RemoteFile.Groups["share"].Value.toLower()

                                $T4WVolume = Get-T4WVolumes -Share $Share -Machine $Machine
                                if($Null -eq $T4WVolume){ Write-Error -ErrorAction Stop -Message "A Volume matching the following could not be found `r`nMachine: $Machine  `r`nShare: $Share `r`n\\$machine\$Share" }
                                
                                break
                            }

                            #Local Drive
                            default{
                                Write-Verbose "This is a Local Path"
                                $VolumeUniqueID = Get-Volume -DriveLetter $DriveLetter -ErrorAction SilentlyContinue | Select-Object UniqueId
                                if($Null -ne $VolumeUniqueID){
                                    #Get the volume from the database by VolumeID
                                    $T4WVolume = Get-T4WVolumes -GUID $VolumeUniqueID.UniqueId

                                    if($Null -eq $T4WVolume){
                                        Write-Verbose "Failed to find the volume by the GUID, Trying by Drive Letter"
                                        #Get the volume from the database by Drive Letter
                                        $T4WVolume = Get-T4WVolumes -Letter $DriveLetter 
                                    }
                                }else{
                                    #Get the volume from the database by Drive Letter
                                    $T4WVolume = Get-T4WVolumes -Letter $DriveLetter

                                }
                                break
                            }
                        }
                        break
                    }
                    #VOLUMEID format
                    $T4WRegexVolume.ToString() {
                        $VolumeREGEX =  $T4WRegexVolume.match($FilePath)
                        $VolumeUniqueID = $VolumeREGEX.Groups['volume'].value

                    
                        $RelativeFilePath = "\$($VolumeREGEX.Groups["path"].Value)"
                        $RelativeFilePath = $RelativeFilePath -replace "[\\]+","\"
                        Write-Verbose "VolumeID: $VolumeUniqueID"
                        Write-Verbose "Relative File: $RelativeFilePath"

                        $T4WVolume = Get-T4WVolumes -GUID $VolumeUniqueID
                        break
                    }


                }

                if($Null -eq $T4WVolume){Write-Error -ErrorAction Stop -Message "Failed to determine the Drive Volume from database"}

                if("$($VolumeUniqueID)".Length -gt 0){ Write-Verbose "VolumeID: $VolumeUniqueID" }
                if("$($Machine)".Length -gt 0){ Write-Verbose "Machine: $Machine" }
                if("$($Share)".Length -gt 0){ Write-Verbose "Share: $Share" }
                if("$($DriveLetter)".Length -gt 0){ Write-Verbose "Drive Letter: $DriveLetter" }

                Write-Verbose "RelativeFilePath: $RelativeFilePath"
                Write-Verbose "`r`n######Volume Record#######`r`n$(($T4WVolume | Out-String).Trim())`r`n##########################"

                #Build the Query with the Volume data we've collected.
                #NOTE the fo_path is intentially a LIKE statement because this is not case sensitive.
                $fileQuery  =  " SELECT fi_id, fi_guid, fi_path, fi_volume, fi_state, fi_ts_path, fi_ts_volume, fi_ts_state "
                $fileQuery  += " FROM fis"
                $fileQuery  += " WHERE fi_volume = '$($T4WVolume.id)' AND fi_path LIKE '$($RelativeFilePath)'; "
                break
            }
            'FileID'{
                $fileQuery  =  " SELECT fi_id, fi_guid, fi_path, fi_volume, fi_state, fi_ts_path, fi_ts_volume, fi_ts_state "
                $fileQuery  += " FROM fis"
                $fileQuery  += " WHERE fi_id = '$($FileID)' ;"
                break
            }
            'FileGUID'{
                $fileQuery  =  " SELECT fi_id, fi_guid, fi_path, fi_volume, fi_state, fi_ts_path, fi_ts_volume, fi_ts_state "
                $fileQuery  += " FROM fis"
                $fileQuery  += " WHERE fi_guid = '{$($FileGUID.Guid)}' ;"
                break
            }
            

        }

        #Execute the Database Query
        $Results = Invoke-SqliteQuery -DataSource $Database -Query $fileQuery -ErrorAction Stop

        #Return a PSCustom Object of the Database Record
        if($Null -ne $Results ){
            return @(foreach($Result in $Results){
            [PSCustomObject]@{
                id = $Result.fi_id
                guid = $Result.fi_guid
                path = $Result.fi_path
                volume = $Result.fi_volume
                state = $Result.fi_state
                path_timestamp = $Result.fi_ts_path
                path_time        = $(try{[datetime]::FromFileTime($Result.fi_ts_path)}catch{""})
                volume_timestamp = $Result.fi_ts_volume
                volume_time        = $(try{[datetime]::FromFileTime($Result.fi_ts_volume)}catch{""})
                state_timestamp = $Result.fi_ts_state
                state_time        = $(try{[datetime]::FromFileTime($Result.fi_ts_state)}catch{""})

            }})
        }else{
            return $Null
        }
    }
}

Function Add-T4WFile{
<#
.SYNOPSIS
Adds File to database

.DESCRIPTION
Adds File to database

.PARAMETER FilePath
Filepath to be added

.PARAMETER Database
Database Path

.EXAMPLE
Add-T4WFile -Path '\\?\Volume{b5b9fcbf-d037-411d-a0c4-7f40fd8062ee}\STL\SomeFile.lys'
WARNING: This Record already Exists

fi_id        : 233
fi_guid      : {7ded430f-74aa-4d02-b6ff-70fbd78ca76d}
fi_path      : \STL\SomeFile.lys
fi_volume    : 5
fi_state     : 0
fi_ts_path   : 133883924143104789
fi_ts_volume : 133883924143104789
fi_ts_state  : 133883924143104789

.NOTES
Will prevent the creation of duplicate records.  A warning will be given and a return of the current record will be returned.
#>
    [alias("Add-File")]
    [cmdletbinding(DefaultParameterSetName = 'UseFileName')]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath")]
        [alias("FileName","Name","Fullname","Path")]
        [string]$FilePath,


        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath



    )
    begin { 

        #CHECK the Guid for a collision
        $GUID = Create-T4WGUID -Type File


        [int64]$TimeStamp = (get-date).ToFileTime()
        $FileRecord = [PSCustomObject]@{
            guid = "{$GUID}"
            path = ''
            volumeid = -1
            state = 0
            path_time = $TimeStamp
            volume_time = $TimeStamp
            state_time = $TimeStamp
        }

        switch -Regex ($Filepath){
            $T4WRegexVolume.ToString(){
                $FileRegexResult = $T4WRegexVolume.Match($FilePath)
                $Volume = Get-T4WVolumes -GUID   $FileRegexResult.Groups['volume'].Value
                if($Null -eq $Volume){
                    Write-Error -ErrorAction Stop -Message "Volume is not present in database, Create the volume record before proceeding"
                }
                $FileRecord.volumeid = $Volume.id
                $FileRecord.path=$FileRegexResult.Groups['path'].Value
                break
            }
            $T4WRegexUNC.ToString(){
                $FileRegexResult = $T4WRegexUNC.Match($FilePath)
                $Volume = Get-T4WVolumes -Machine   $FileRegexResult.Groups['server'].Value -Share $FileRegexResult.Groups['share'].Value
                if($Null -eq $Volume){
                    Write-Error -ErrorAction Stop -Message "Volume is not present in database, Create the volume record before proceeding"
                }
                $FileRecord.volumeid = $Volume.id
                $FileRecord.path=$FileRegexResult.Groups['path'].Value

                break
            }
            $T4WRegexDriveLetter.ToString(){
                $FileRegexResult = $T4WRegexDriveLetter.Match($FilePath)
                $Volume = Get-T4WVolumes -Letter   $FileRegexResult.Groups['driveletter'].Value
                if($Null -eq $Volume){
                    Write-Error -ErrorAction Stop -Message "Volume is not present in database, Create the volume record before proceeding"
                }
                $FileRecord.volumeid = $Volume.id
                $FileRecord.path=$FileRegexResult.Groups['path'].Value
                break
            }
        }
    }

    process {

        if($FileRecord.path -match "\\$"){
            Write-Error "File must NOT end in a trailing ""\""  `r`n`tPATH presented was ""$($FileRecord.path)""" -ErrorAction Stop
        }

        $UpdateInfoQuery = "Update info SET info_value_i = '$($FileRecord.fo_ts_path)' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:"
        Write-Verbose $UpdateInfoQuery

        $INSERTQuery =  " INSERT INTO fis (fi_guid,fi_path,fi_volume,fi_state,fi_ts_path,fi_ts_volume,fi_ts_state) "
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}'); "

        $InsertStatement = $INSERTQuery -f $FileREcord.guid,$($FileRecord.path -replace "'","''" ),$FileRecord.volumeid,$FileRecord.state,$FileRecord.path_time,$FileRecord.volume_time,$FileRecord.state_time

        $CheckQuery = "SELECT * FROM fis WHERE fi_volume ='$($FileRecord.volumeid)' AND LOWER(fi_path) ='$($FileREcord.path.ToLower() -replace "'","''")'"

        $CheckResults = Invoke-SqliteQuery -DataSource $Database -Query $CheckQuery
        if($Null -ne $CheckResults){
            Write-Warning  -Message "This Record already Exists"
            return Get-T4WFile -FileGuid $CheckResults.fi_guid
        }

        Write-Verbose "Preparing to Insert Record"
        Write-Verbose "INSERT Query: $InsertStatement"

        Invoke-SqliteQuery -DataSource $Database -Query $InsertStatement -ErrorAction Stop 
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

        return  Get-T4WFile -FileGuid $FileRecord.guid
    }
}

Function Update-T4WFile {
    <#
.SYNOPSIS
Updates File

.DESCRIPTION
Can be used to move files and then update recoreds to reflect new location.

.PARAMETER ID
File ID

.PARAMETER NewRelativePath
New Relative Path

.PARAMETER NewParent
New Volume ID

.PARAMETER Database
Database Path

.EXAMPLE
Update-T4WFile -ID 4760 -NewParent  5 -NewRelativePath "\kb16b-02 V2.json2"

id               : 4760
guid             : {6e21e076-7eb5-4a67-b12c-5327ebb94216}
path             : \kb16b-02 V2.json2
volume           : 5
state            : 0
path_timestamp   : 133894274666230655
path_time        : 4/18/2025 1:24:26 AM
volume_timestamp : 133894274666230655
volume_time      : 4/18/2025 1:24:26 AM
state_timestamp  : 133892905815199601
state_time       : 4/16/2025 11:23:01 AM

#>
    [cmdletbinding()]
    param(

        [Parameter(Mandatory=$True)]
        [string]$ID,

        [Parameter(Mandatory=$False)]
        [ValidatePattern("[^<>:""|?*//`t`r`n`e]")]
        [string]$NewRelativePath="",

        [Parameter(Mandatory=$False)]
        [int]$NewParent=-99,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath


    )
    begin {

        [int64]$TimeStamp = (get-date).ToFileTime()


    }
    process {
        $File = Get-T4WFile -id $ID
        if($Null -eq $File){Write-Error -ErrorAction Stop -Message "Invalid FileID"}

        [string[]] $SQLUpdates = @()
        if($NewParent -ne -99){
            $ParentRecord=Get-T4WVolumes -id $NewParent
            if($Null -eq $ParentRecord){ Write-Error -ErrorAction Stop -Message "Invalid New Parent ID" }
            $SQLUpdates += " fi_volume = '{0}' " -f $ParentRecord.id
            $SQLUpdates += " fi_ts_volume = '{0}' " -f $TimeStamp
            Write-Verbose "Volume Changing to $($ParentRecord.id)"
            Write-Verbose "####################`r`n$($ParentREcord|Out-string)`r`n####################"

        }

        if($T4WRegexDriveLetter.match( $NewRelativePath).success){
            Write-Error -ErrorAction Stop -Message "New Relative Path should not be a full path.  It must not start with a drive letter and must begin with a '\'"  

        }elseif($T4WRegexUNC.match( $NewRelativePath).success){
            Write-Error -ErrorAction Stop -Message "New Relative Path should not be a UNC path.  It must not start with a UNC Server\Share Reference and must begin with a '\'"  

        }elseif($T4WRegexVolume.match( $NewRelativePath).success){
            Write-Error -ErrorAction Stop -Message "New Relative Path should not be a full path.  It must not start with a Volume Mapping Reference and must begin with a '\'"  
        }elseif([regex]::match( $NewRelativePath,"^\\.*[\\]$").success){
            Write-Error -ErrorAction Stop -Message "New Relative Path cannot have a trailing '\', This indicates it is a folder and not a file."  



        }elseif([regex]::match( $NewRelativePath,"^\\.*[^\\]$").success){
            $SQLUpdates += " fi_path = '{0}' " -f $($NewRelativePath -replace "'","''")
            $SQLUpdates += " fi_ts_path = '{0}' " -f $TimeStamp
            Write-Verbose "Updating Path"
            }    
        $SETCommands = $SQLUpdates | Select-Object -Unique
        
        if($SETCommands.Count -eq 0){
            write-warning "No Valid Changes were provided"
            return  $Null
        }
        
        $UpdateInfoQuery = "Update info`r`n  SET info_value_i = '$TimeStamp'`r`n  WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"



        $UpdateQuery = "UPDATE fis SET $( $SETCommands -join " , ") WHERE fi_id = '$($File.id)'"
        write-verbose "UPDATE QUERY: $UpdateQuery"

        Invoke-SqliteQuery -DataSource $Database -Query $UpdateQuery -ErrorAction Stop
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

        return Get-T4WFile -FileID $File.ID
    }
}


function Remove-T4WFile{
    <#
    .SYNOPSIS
    Remove File from Database
    
    .DESCRIPTION
    Copies duplicate to archive table and deletes original record.
    
    .PARAMETER FilePath
    File Path
    
    .PARAMETER FileID
    File ID
    
    .PARAMETER FileGuid
    File GUID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Remove-T4WFile -FilePath 'G:\SFTP_Root\t1\newfile.txt'
    
    .NOTES
    This will delete any FileComments or FileTags records associated with this record.

    #>
    [cmdletbinding()]
    param(
        
    [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath",Position=0)]
    [alias("FileName","Name","Fullname","Path")]
    [string]$FilePath,

    [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileID")]
    [alias("ID")]
    [int]$FileID,

    [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileGUID")]
    [alias("GUID")]
    [GUID]$FileGuid,

    [Parameter(Mandatory=$False,Position=1)]
    [string]$Database=$TagDBPath
    )
    process{
        #Get the Timestamp fir updating records
        [int64]$TimeStamp = (get-date).ToFileTime()
        
        switch($PsCmdlet.ParameterSetName){
            'FilePath'{ 
                $File = Get-T4WFile -Path $FilePath
                if($Null -eq $File){Write-Error -ErrorAction Stop -Message "FilePath is not present in the database."}
                break
            }
            'FileGUID'{ 
                $File = Get-T4WFile -FileGuid $FileGuid
                if($Null -eq $File){Write-Error -ErrorAction Stop -Message "FileGuid is not present in the database."}
                break
            }
            'FileID'{ 
                $File = Get-T4WFile -FileID $FileID
                if($Null -eq $File){Write-Error -ErrorAction Stop -Message "FileID is not present in the database."}
                break
            }
        }

        $Volume = Get-T4WVolumes -ID $File.volume

        
        #We need to clean up Tags so they aren't orphaned
        foreach($Tag in (Get-T4WFileTag -FileID $File.id)){
            Remove-T4WFileTag -GUID $Tag.guid 
        }

        #We need to clean up Comments so they aren't orphaned
        foreach($Comment in (Get-T4WFileComments -FileID $File.id)){
            #Remove-T4WFileComment -CommentID $Comment.id
            Remove-T4WFileComment -ID $Comment.id
            
        }
        $INSERT = " INSERT INTO dfis (dfi_guid,dfi_path,dfi_vb_serial,dfi_vb_guid,dfi_vb_letter,dfi_state,dfi_ts) "
        $INSERT +=" VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}') ;" -f `
                    $File.guid, $($File.path -replace  "'","''" ) ,$Volume.serial,$volume.guid,$Volume.letter,$File.state,$Timestamp

        
        Write-Verbose "INSERT:  $INSERT"

        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"
        

        $DeleteQuery = "DELETE FROM fis WHERE fi_id ='{0}'" -f $File.id
        write-verbose "Delete Command:  $DeleteQuery"


        Invoke-SQLiteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SQLiteQuery -DataSource $Database -Query $DeleteQuery -ErrorAction Continue
        Invoke-SQLiteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction SilentlyContinue


    }
}

#endregion


#region FolderComments
Function Get-T4WFolderComments{
    <#
    .SYNOPSIS
    Gets array of Folder Comments
    
    .DESCRIPTION
    Collects the list of fodler comments for a specified folder
    
    .PARAMETER FolderPath
    Folder Path
    
    .PARAMETER FolderID
    Folder ID
    
    .PARAMETER FolderGuid
    Folder GUID
    
    .PARAMETER CommentID
    CommentID from this table (Use if trying to pull single record)
    
    .PARAMETER CommentGUID
    CommentGUID From this table  (Use if trying to pull single record)
    
    .PARAMETER Database
    Parameter description
    
    .EXAMPLE
Get-T4WFolderComment

id             : 2
guid           : {ded5f988-5bcf-4976-a028-421e7a9631ab}
folderid       : 3
text           : Chocolatey Repository Folder
timestamp      : 133889452165248771
time           : 4/12/2025 11:26:56 AM
text_timestamp : 133889452165248771
text_time      : 4/12/2025 11:26:56 AM
user           : YourUsername
last_user      : YourUsername

    #>
    [alias("Get-T4WFolderComment","Get-FolderComment")]
    [cmdletbinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath")]
        [alias("FolderName","Name","Fullname","Path","FilePath")]
        [string]$FolderPath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderID")]
        [int]$FolderID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderGUID")]
        [alias("GUID")]
        [GUID]$FolderGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="CommentID")]
        [int]$CommentID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="CommentGUID")]
        [GUID]$CommentGUID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    begin { }
    process {
        

        switch($PsCmdlet.ParameterSetName){
            'FolderPath'{ 
                $T4WFolder = Get-T4WFolder -FolderPath $FolderPath

                if($Null -eq $T4WFolder){Write-Error -ErrorAction STOP -Message "Folder Path is not present in the database"}

                $CommentQuery  = " SELECT cfo_id,cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last "
                $CommentQuery += " FROM cfos "
                $CommentQuery += " WHERE cfo_fo = '$($T4WFolder.id)' ;"

                break
            }
            'FolderGUID'{ 
                $T4WFolder = Get-T4WFolder -FolderGuid $FolderGuid

                if($Null -eq $T4WFolder){Write-Error -ErrorAction STOP -Message "Folder GUID is not present in the database"}

                $CommentQuery  = " SELECT cfo_id,cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last "
                $CommentQuery += " FROM cfos "
                $CommentQuery += " WHERE cfo_fo = '$($T4WFolder.id)' ;"
                break
            }
            'FolderID'{ 
                $T4WFolder = Get-T4WFolder -id $FolderID

                if($Null -eq $T4WFolder){Write-Error -ErrorAction STOP -Message "Folder ID is not present in the database"}

                $CommentQuery  = " SELECT cfo_id,cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last "
                $CommentQuery += " FROM cfos "
                $CommentQuery += " WHERE cfo_fo = '$($T4WFolder.id)' ;"
                break
            }
            'CommentID'{ 
                $CommentQuery  = " SELECT cfo_id,cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last "
                $CommentQuery += " FROM cfos "
                $CommentQuery += " WHERE cfo_id = '$CommentID' ;"
                break
            }
            'CommentGUID'{ 
                $CommentQuery  = " SELECT cfo_id,cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last "
                $CommentQuery += " FROM cfos "
                $CommentQuery += " WHERE cfo_guid = '{$($CommentGuid.Guid)}' ;"
                break
            }
            default{
                $CommentQuery  = " SELECT cfo_id,cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last "
                $CommentQuery += " FROM cfos "
            }
        }

        
        Write-Verbose "`r`nQuery Being Executed:`r`n $CommentQuery"
        
        $Results = Invoke-SqliteQuery -DataSource $Database -Query $CommentQuery -ErrorAction Stop
        if($Null -eq $Results){ 
            return $Null
        }else{
            
            return @(foreach($Result in $Results){
                [PSCustomObject]@{
                    id               =   $Result.cfo_id
                    guid             =   $Result.cfo_guid
                    folderid         =   $Result.cfo_fo
                    text             =   $Result.cfo_text
                    timestamp        =   $Result.cfo_timestamp
                    time             =   $(try{[datetime]::FromFileTime($Result.cfo_timestamp)}catch{""})
                    text_timestamp   =   $Result.cfo_ts_text
                    text_time        =   $(try{[datetime]::FromFileTime($Result.cfo_ts_text)}catch{""})
                    user             =   $Result.cfo_user
                    last_user        =   $Result.cfo_user_last
                }
            })
        }
    }
}

Function Add-T4WFolderComment{
    <#
    .SYNOPSIS
    Adds Folder Comments to database
    
    .DESCRIPTION
    Adds Folder Comments to database
    
    .PARAMETER FolderPath
    Folder Path, this must end in a "/" or there will be an exception.
    
    .PARAMETER FolderID
    Folder ID
    
    .PARAMETER FolderGuid
    Folder GUID
    
    .PARAMETER FolderComment
    Comment to be applied
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Add-T4WFolderComment -GUID '{2b0b6fde-fae0-4450-8df8-023bcbb5ff02}' -Comment "I Love to make Comments"
    
id               : 11
guid             : {838b7adb-54c4-4be2-b9d0-3835a6149a0c}
folder           : 1
comment          : I Love to make Comments
modified         : 133894279745394833
comment_modified : 133894279745394833
user             : Exodu
last_user        : Exodu 

    #>
    [Alias("Add-FolderComment")]
    [cmdletbinding(DefaultParameterSetName = 'UseFolderName')]
    param(

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath")]
        [alias("FolderName","Name","Fullname","Path","FilePath")]
        [string]$FolderPath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderID")]
        [alias("ID","fos_id")]
        [int]$FolderID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderGUID")]
        [alias("GUID","fos_guid")]
        [GUID]$FolderGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [alias("Comment")]
        [string]$FolderComment,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    begin {
        #Get the Timestamp for updating records
        $TimeStamp = (get-date).ToFileTime()
        
        #Need to make sure any single quotes are escaped
        $FolderComment=$FolderComment -replace "[']","''"  
        

        #CHECK the Guid for a collision
        $GUID = Create-T4WGUID -Type Folder_Comment


        [int64]$TimeStamp = (get-date).ToFileTime()
        $CommentRecord = [PSCustomObject]@{
            guid = "{$GUID}"
            folder_id = -1
            comment = "$FolderComment"
            modified_time = $TimeStamp
            comment_time = $TimeStamp
            user = "${env:USERNAME}"
            user_last = "${env:USERNAME}"
        }

        switch($PsCmdlet.ParameterSetName){
            'FolderPath'{ 
                $Folder = Get-T4WFolder -Path $FolderPath
                $CommentRecord.folder_id = $Folder.id
                break
            }
            'FolderGUID'{ 
                $Folder = Get-T4WFolder -FolderGuid $FolderGuid
                $CommentRecord.folder_id = $Folder.id
                break
            }
            'FolderID'{ 
                $Folder = Get-T4WFolder -FolderID $FolderID
                $CommentRecord.folder_id = $Folder.id
            }
        }

        if($CommentRecord.folder_id -eq -1){
            Write-Error -ErrorAction Stop -Message "Was Unable to find the folder"
        }


        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command: `r`n`t $UpdateInfoQuery"

        $INSERTQuery  = " INSERT INTO cfos (cfo_guid,cfo_fo,cfo_text,cfo_timestamp,cfo_ts_text,cfo_user,cfo_user_last) "
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}');"
        $INSERT = $INSERTQuery -f $CommentRecord.guid,$CommentRecord.folder_id,$CommentRecord.comment,$CommentRecord.modified_time,$CommentRecord.comment_time,$CommentRecord.user,$CommentRecord.user_last
        write-verbose "Insert Command: `r`n`t $UpdateInfoQuery"

        Invoke-SqliteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction STOP
        
        return Invoke-SqliteQuery -DataSource $Database -Query "SELECT cfo_id as id,cfo_guid as guid,cfo_fo as folder,cfo_text as comment,cfo_timestamp as modified,cfo_ts_text as comment_modified,cfo_user as user,cfo_user_last as last_user FROM cfos WHERE cfo_guid = '$($CommentRecord.guid)'"

    }
}

Function Update-T4WFolderComment {
    <#
    .SYNOPSIS
    Updates the Folder Comments
    
    .DESCRIPTION
    Allows you to update the Comment Text
    
    .PARAMETER CommentID
    Comment ID
    
    .PARAMETER Comment
    Text to be replaced
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Update-T4WFolderComment -ID 11 -Comment "I Really Really Like making Comments'"

id             : 11
guid           : {838b7adb-54c4-4be2-b9d0-3835a6149a0c}
folderid       : 1
text           : I Love to make Comments
timestamp      : 133894279745394833
time           : 4/18/2025 1:32:54 AM
text_timestamp : 133894279745394833
text_time      : 4/18/2025 1:32:54 AM
user           : Exodu
last_user      : Exodu

    #>
    [cmdletbinding()]
    param(

        [Parameter(Mandatory=$True)]
        [alias('ID')]
        [string]$CommentID,

        [Parameter(Mandatory=$True)]
        [alias('NewComment')]
        [string]$Comment,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath


    )
    begin {

        [int64]$Timestamp = (get-date).ToFileTime()


    }
    process {
        $FolderComment = Get-T4WFolderComment -CommentID $CommentID
        if($Null -eq $FolderComment){Write-Error -ErrorAction Stop -Message "Invalid FolderID"}

        $UpdateInfoQuery = "Update info`r`n  SET info_value_i = '$Timestamp'`r`n  WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        $UpdateQuery = "UPDATE cfos SET cfo_text = '$( $Comment -replace "'","''")' WHERE cfo_id = '$($FolderComment.folderid)'"
        write-verbose "UPDATE QUERY: $UpdateQuery"

        
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateQuery -ErrorAction Stop
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

        return Get-T4WFolderComment -FolderID $FolderComment.folderid  

    }
}

function Remove-T4WFolderComment{
    <#
    .SYNOPSIS
    Removes Folder Comment
    
    .DESCRIPTION
    Inserts Folder Comment into the Archive Table, then deletes the Comment
    
    .PARAMETER CommentID
    Comment ID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Remove-T4WFolderComment -CommentID 11

    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [alias('ID')]
        [int]$CommentID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath
    )

    process {
        $CommentRecord = Get-T4WFolderComment -CommentID $CommentID

        [int64]$Timestamp = (get-date).ToFileTime()

        $INSERT =  " INSERT INTO dcfos ( dcfo_guid, dcfo_fo, dcfo_text, dcfo_timestamp, dcfo_ts_text, dcfo_user, dcfo_user_last, dcfo_ts) "
        $INSERT += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}');" 
        $INSERT = $INSERT -f $CommentRecord.Guid, `
                             $commentRecord.FolderId, `
                             ($CommentRecord.text -replace "'","''"),`
                             $CommentRecord.timestamp,`
                             $CommentRecord.text_timestamp,`
                             $CommentRecord.user,`
                             $CommentRecord.last_user,`
                             $Timestamp

        Write-Verbose "Deletion Record Query: $Insert"

        $DeleteQuery = "DELETE FROM cfos WHERE cfo_id = '{0}' " -f $CommentRecord.id
        Write-Verbose "Deletion Query: $DeleteQuery"

        $UpdateInfoQuery = "Update info SET info_value_i = '{0}' WHERE info_key = '-7001';" -f $Timestamp
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        Invoke-SQLiteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SQLiteQuery -DataSource $Database -Query $DeleteQuery -ErrorAction Continue
        Invoke-SQLiteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction SilentlyContinue
    }
}
#endregion


#region FileComments
Function Get-T4WFileComments{
    <#
    .SYNOPSIS
    Gets File Comments
    
    .DESCRIPTION
    Gets all file comments OR specific comments depending on parameters supplied.
    
    .PARAMETER FilePath
    FilePath (Cannot end in a \)
    
    .PARAMETER FileID
    File ID
    
    .PARAMETER FileGuid
    FileGUID
    
    .PARAMETER CommentID
    Comment ID
    
    .PARAMETER CommentGUID
    Comment GUID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Get-T4WFileComment -CommentID 3

id             : 3
guid           : {3fa4c51d-6c57-490b-ad27-785509c4c60f}
Fileid         : 4752
text           : This is a Modified Comment Enjoy🍆
timestamp      : 133889512430183355
time           : 4/12/2025 1:07:23 PM
text_timestamp : 133889512430183355
text_time      : 4/12/2025 1:07:23 PM
user           : Exodu
last_user      : Exodu
    
    .NOTES
    General notes
    #>
    [alias("Get-T4WFileComment")]
    [cmdletbinding(DefaultParameterSetName = 'UseFileName')]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath")]
        [alias("FileName","Name","Fullname","Path")]
        [string]$FilePath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileID")]
        [int]$FileID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileGUID")]
        [GUID]$FileGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="CommentID")]
        [alias("ID")]
        [int]$CommentID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="CommentGUID")]
        [GUID]$CommentGUID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    process {
        switch($PsCmdlet.ParameterSetName){
            'FilePath'{ 
                $T4WFile = Get-T4WFile -FilePath $FilePath

                if($Null -eq $T4WFile){Write-Error -ErrorAction STOP -Message "File Path is not present in the database"}

                $CommentQuery  = " SELECT cfi_id,cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last "
                $CommentQuery += " FROM cfis "
                $CommentQuery += " WHERE cfi_fi = '$($T4WFile.id)' ;"

                break
            }
            'FileGUID'{ 
                $T4WFile = Get-T4WFile -FileGuid $FileGuid

                if($Null -eq $T4WFile){Write-Error -ErrorAction STOP -Message "File GUID is not present in the database"}

                $CommentQuery  = " SELECT cfi_id,cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last "
                $CommentQuery += " FROM cfis "
                $CommentQuery += " WHERE cfi_fi = '$($T4WFile.id)' ;"
                break
            }
            'FileID'{ 
                $T4WFile = Get-T4WFile -id $FileID

                if($Null -eq $T4WFile){Write-Error -ErrorAction STOP -Message "File ID is not present in the database"}

                $CommentQuery  = " SELECT cfi_id,cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last "
                $CommentQuery += " FROM cfis "
                $CommentQuery += " WHERE cfi_fi = '$($T4WFile.id)' ;"
                break
            }
            'CommentID'{ 
                $CommentQuery  = " SELECT cfi_id,cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last "
                $CommentQuery += " FROM cfis "
                $CommentQuery += " WHERE cfi_id = '$CommentID' ;"
                break
            }
            'CommentGUID'{ 
                $CommentQuery  = " SELECT cfi_id,cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last "
                $CommentQuery += " FROM cfis "
                $CommentQuery += " WHERE cfi_guid = '{$($CommentGuid.Guid)}' ;"
                break
            }
            default{
                $CommentQuery  = " SELECT cfi_id,cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last "
                $CommentQuery += " FROM cfis "
                $CommentQuery += " WHERE cfi_guid = '{$($CommentGuid.Guid)}' ;"
            }
        }

        
        Write-Verbose "`r`nQuery Being Executed:`r`n $CommentQuery"
        
        $Results = Invoke-SqliteQuery -DataSource $Database -Query $CommentQuery -ErrorAction Stop
        if($Null -eq $Results){ 
            return $Null
        }else{
            
            return @(foreach($Result in $Results){
                [PSCustomObject]@{
                    id               =   $Result.cfi_id
                    guid             =   $Result.cfi_guid
                    Fileid           =   $Result.cfi_fi
                    text             =   $Result.cfi_text
                    timestamp        =   $Result.cfi_timestamp
                    time             =   $(try{[datetime]::FromFileTime($Result.cfi_timestamp)}catch{""})
                    text_timestamp   =   $Result.cfi_ts_text
                    text_time        =   $(try{[datetime]::FromFileTime($Result.cfi_ts_text)}catch{""})
                    user             =   $Result.cfi_user
                    last_user        =   $Result.cfi_user_last
                }
            })
        }
    }
}

Function Add-T4WFileComments{
    <#
    .SYNOPSIS
    Adds File Comments
    
    .DESCRIPTION
    Adds File Comments
    
    .PARAMETER FilePath
    File Path
    
    .PARAMETER FileID
    File ID
    
    .PARAMETER FileGuid
    File GUID
    
    .PARAMETER FileComment
    The Comment to be applied
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Add-T4WFileComments -Filepath 'G:\STL\Gridfity.stl' -comment "So Much Gridfinity, So WoW" 

id               : 13
guid             : {3119c539-c08d-4963-a4a2-5ce325d5ed1b}
File             : 1348
comment          : So Much Gridfinity, So WoW
modified         : 133894287765676646
comment_modified : 133894287765676646
user             : YourUser
last_user        : YourUser
    

    #>
    [Alias("Add-FileComment")]
    [cmdletbinding(DefaultParameterSetName = 'UseFileName')]
    param(

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath")]
        [alias("FileName","Name","Fullname","Path")]
        [string]$FilePath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileID")]
        [alias("ID","fis_id")]
        [int]$FileID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileGUID")]
        [alias("GUID","fis_guid")]
        [GUID]$FileGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [alias("Comment")]
        [string]$FileComment,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    begin {
        #Get the Timestamp for updating records
        $TimeStamp = (get-date).ToFileTime()
        
        #Need to make sure any single quotes are escaped
        $FileComment=$FileComment -replace "[']","''"  
        

        #CHECK the Guid for a collision
        $GUID = Create-T4WGUID -Type File_Comment


        [int64]$TimeStamp = (get-date).ToFileTime()
        $CommentRecord = [PSCustomObject]@{
            guid = "{$GUID}"
            File_id = -1
            comment = "$FileComment"
            modified_time = $TimeStamp
            comment_time = $TimeStamp
            user = "${env:USERNAME}"
            user_last = "${env:USERNAME}"
        }

        switch($PsCmdlet.ParameterSetName){
            'FilePath'{ 
                $File = Get-T4WFile -Path $FilePath
                $CommentRecord.File_id = $File.id
                break
            }
            'FileGUID'{ 
                $File = Get-T4WFile -FileGuid $FileGuid
                $CommentRecord.File_id = $File.id
                break
            }
            'FileID'{ 
                $File = Get-T4WFile -FileID $FileID
                $CommentRecord.File_id = $File.id
            }
        }

        if($CommentRecord.File_id -eq -1){
            Write-Error -ErrorAction Stop -Message "Was Unable to find the File"
        }

        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command: `r`n`t $UpdateInfoQuery"

        $INSERTQuery  = " INSERT INTO cfis (cfi_guid,cfi_fi,cfi_text,cfi_timestamp,cfi_ts_text,cfi_user,cfi_user_last) "
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}');"
        $INSERT = $INSERTQuery -f $CommentRecord.guid,$CommentRecord.File_id,$CommentRecord.comment,$CommentRecord.modified_time,$CommentRecord.comment_time,$CommentRecord.user,$CommentRecord.user_last
        write-verbose "Insert Command: `r`n`t $INSERT"

        Invoke-SqliteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery
        
        return Invoke-SqliteQuery -DataSource $Database -Query "SELECT cfi_id as id,cfi_guid as guid,cfi_fi as File,cfi_text as comment,cfi_timestamp as modified,cfi_ts_text as comment_modified,cfi_user as user,cfi_user_last as last_user FROM cfis WHERE cfi_guid = '$($CommentRecord.guid)'"
    }
}

Function Update-T4WFileComment {
    <#
    .SYNOPSIS
    Updates File Comment
    
    .DESCRIPTION
    Allows you to change the Comment Text
    
    .PARAMETER CommentID
    Comment ID
    
    .PARAMETER Comment
    Comment Text to be replaced
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Update-T4WFileComment -CommentID 13 -Comment "🍆 Are Delish"                              

id             : 13
guid           : {3119c539-c08d-4963-a4a2-5ce325d5ed1b}
Fileid         : 1348
text           : 🍆 Are Delish
timestamp      : 133894287765676646
time           : 4/18/2025 1:46:16 AM
text_timestamp : 133894287765676646
text_time      : 4/18/2025 1:46:16 AM
user           : Exodu
last_user      : Exodu
    

    #>
    [cmdletbinding()]
    param(

        [Parameter(Mandatory=$True)]
        [alias('ID')]
        [int]$CommentID,

        [Parameter(Mandatory=$True)]
        [alias('NewComment')]
        [string]$Comment,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath


    )
    begin {
        [int64]$TimeStamp = (get-date).ToFileTime()
    }
    process {
        $FileComment = Get-T4WFileComment -CommentID $CommentID
        if($Null -eq $FileComment){Write-Error -ErrorAction Stop -Message "Invalid CommentID"}

        $UpdateInfoQuery = "Update info`r`n  SET info_value_i = '$TimeStamp'`r`n  WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        $UpdateQuery = "UPDATE cfis SET cfi_text = '$( $Comment -replace "'","''")' WHERE cfi_id = '$($CommentID)'"
        write-verbose "UPDATE QUERY: $UpdateQuery"

        
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateQuery -ErrorAction Stop
        Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

        Write-Verbose $FileComment
        return Get-T4WFileComment -CommentID $CommentID
    }
}

function Remove-T4WFileComment{
    <#
    .SYNOPSIS
    Remove File Comment
    
    .DESCRIPTION
    Inserts a replica record into the Archive Table, then deletes the original.
    
    .PARAMETER CommentID
    Comment ID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Remove-T4WFileComment -CommentID 13

    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [alias('ID')]
        [int]$CommentID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath
    )

    process {
        $CommentRecord = Get-T4WFileComment -CommentID $CommentID

        [int64]$TimeStamp = (get-date).ToFileTime()

        $INSERT =  " INSERT INTO dcfis (dcfi_id, dcfi_guid, dcfi_fi, dcfi_text, dcfi_timestamp, dcfi_ts_text, dcfi_user, dcfi_user_last, dcfi_ts) "
        $INSERT += " VALUES ('{0}','{1}','{2}','{3}','{4}','{5}','{6}','{7}','{8}');" 
        $INSERT = $INSERT -f $CommentRecord.id, `
                             $CommentRecord.Guid, `
                             $commentRecord.FileId, `
                             ($CommentRecord.text -replace "'","''"),`
                             $CommentRecord.timestamp,`
                             $CommentRecord.text_timestamp,`
                             $CommentRecord.user,`
                             $CommentRecord.last_user,`
                             $TimeStamp

        Write-Verbose "Deletion Record Query: $Insert"

        $DeleteQuery = "DELETE FROM cfis WHERE cfi_id = '{0}' " -f $CommentRecord.id
        Write-Verbose "Deletion Query: $DeleteQuery"

        $UpdateInfoQuery = "Update info SET info_value_i = '{0}' WHERE info_key = '-7001';" -f $TimeStamp
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        Invoke-SQLiteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SQLiteQuery -DataSource $Database -Query $DeleteQuery -ErrorAction Continue
        Invoke-SQLiteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction SilentlyContinue
    }


}
#endregion


#region FolderTag
function Get-T4WFolderTag {
    <#
    .SYNOPSIS
    Gets list of Tags associated with a file
    
    .DESCRIPTION
    Gets list of Tags associated with a file
    
    .PARAMETER FolderPath
    Folder Path
    
    .PARAMETER FolderID
    Folder ID
    
    .PARAMETER FolderGuid
    Folder GUID
    
    .PARAMETER GUID
    GUID of the FolderTag
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Get-T4WFolderTag -FolderPath G:\STL\ 

guid            : {ae7baf33-cc88-41b2-8673-460db708d96e}
tag_id          : 746
Folder_id       : 1
state           : 0
state_timestamp : 133892545252381794
state_time      : 4/16/2025 1:22:05 AM
    

    #>
    [cmdletbinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath")]
        [alias("FolderName","Name","Fullname","Path")]
        [string]$FolderPath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderID")]
        [int]$FolderID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderGUID")]
        [GUID]$FolderGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="LINKGUID")]
        [GUID]$GUID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    process {
        $Query = ""
        switch($PsCmdlet.ParameterSetName){
            'FolderPath' {
                $Folder = Get-T4wFolder -FolderPath $FolderPath
                if($Null -eq $Folder){
                    Write-Error -ErrorAction Stop -Message "Invalid Folder Supplied"
                }
                
                $Query = "SELECT * from xfos WHERE xfo_fo = '{0}'" -f $Folder.id

                break
            }
            'FolderID'   {
                $Folder = Get-T4wFolder -FolderID $FolderID
                if($Null -eq $Folder){
                    Write-Error -ErrorAction Stop -Message "Invalid Folder ID Supplied"
                }
                
                $Query = "SELECT * from xfos WHERE xfo_fo = '{0}'" -f $Folder.id

                break
            }
            'FolderGUID' {
                $Folder = Get-T4wFolder -FolderGuid $FolderGuid
                if($Null -eq $Folder){
                    Write-Error -ErrorAction Stop -Message "Invalid Folder Supplied"
                }
                
                $Query = "SELECT * from xfos WHERE xfo_fo = '{0}'" -f $Folder.id

                break
            }
            'LINKGUID' {
                
                $Query = "SELECT *  from xfos WHERE xfo_guid = '{0}'" -f "{$($GUID.guid)}"

                break
            }
            'Default'   {
                $Query = "SELECT *  from xfos;"
                break
            }
        }

        Write-Verbose "Query: $Query"
        $Results = Invoke-SqliteQuery -DataSource $Database -Query $Query
        if($Null -ne $Results){
            return @(foreach($Result in $Results){
                [PSCustomObject]@{
                    guid = $Result.xfo_guid
                    tag_id = $Result.xfo_to
                    Folder_id = $Result.xfo_fo
                    state  =    $Result.xfo_state
                    state_timestamp = $Result.xfo_ts_state
                    state_time =       $(try{[datetime]::FromFileTime($Result.xfo_ts_state)}catch{""})

                }
            })
        }else{
            write-warning "Nothing collected"
            return $Null
        }
    }
}    

function Add-T4WFolderTag{
    <#
    .SYNOPSIS
    Adds a Foldertag
    
    .DESCRIPTION
    Adds a Tag to a folder
    
    .PARAMETER FolderPath
    Path to File
    
    .PARAMETER FolderID
    Folder ID
    
    .PARAMETER FolderGuid
    Folder GUID
    
    .PARAMETER TagID
    TAG ID to be applied
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Add-T4WFolderTag -FolderPath G:\STL\ -TagID 746 
WARNING: Tag and Folder Combination Already Exists

guid            : {5df21138-cb08-4bb5-93f3-d6a5ccf4f3fb}
tag_id          : 746
file_id         : 1
state           : 0
state_timestamp : 133894295837767996
    
    .NOTES
    General notes
    #>
    [Alias("Add-FolderTag")]
    [cmdletbinding(DefaultParameterSetName = 'FolderName')]
    param(

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderPath")]
        [alias("FolderName","Name","Fullname","Path","FilePath")]
        [string]$FolderPath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderID")]
        [alias("ID","fos_id")]
        [int]$FolderID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FolderGUID")]
        [alias("GUID","fos_guid")]
        [GUID]$FolderGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [int]$TagID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )

    begin {
        #Get the Timestamp for updating records
        [int64]$TimeStamp = (get-date).ToFileTime()
        
        #Need to make sure any single quotes are escaped
        $FolderComment=$FolderComment -replace "[']","''"  
        
        #CHECK the Guid for a collision
        $GUID = Create-T4WGUID -Type Folder_Tag

        $TagRecord = [PSCustomObject]@{
            guid = "{$($GUID)}"
            tag_id = ''
            folder_id = ''
            state = 0
            state_time = $TimeStamp

        }



        switch($PsCmdlet.ParameterSetName){
            'FolderPath'{ 
                $Folder = Get-T4WFolder -Path $FolderPath
                if($Null -eq $Folder){
                    $Folder = Add-T4WFolder -FilePath $FolderPath
                }
                if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "FolderPath is not present in the database."}
                $TagRecord.folder_id = $Folder.id
                break
            }
            'FolderGUID'{ 
                $Folder = Get-T4WFolder -FolderGuid $FolderGuid
                if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "FolderGuid is not present in the database."}
                $TagRecord.folder_id = $Folder.id
                break
            }
            'FolderID'{ 
                $Folder = Get-T4WFolder -FolderID $FolderID
                if($Null -eq $Folder){Write-Error -ErrorAction Stop -Message "FolderID is not present in the database."}
                $TagRecord.folder_id = $Folder.id
            }
        }

        if($TagRecord.folder_id -eq -1){
            Write-Error -ErrorAction Stop -Message "Was Unable to find the folder"
        }

        $TAGInfo = get-T4WTags -ID $TagID
        if($null -eq $TAGInfo){
            Write-Error -ErrorAction STOP -Message 'TagID is invalid'
        }else{
            $TagRecord.tag_id = $Taginfo.id
            Write-Verbose "TAG Name: '$($Taginfo.Name)'"
        }

        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command: `r`n`t $UpdateInfoQuery"

        $INSERTQuery =  " INSERT INTO xfos (xfo_guid,xfo_to,xfo_fo,xfo_state,xfo_ts_state) "
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}');"
        $Insert = $INSERTQuery  -f $TagRecord.guid,$TagRecord.tag_id,$TagRecord.folder_id,$TagRecord.state,$TagRecord.state_time
        Write-verbose "INSERT Statement: $Insert"

        $Check = "SELECT xfo_guid as guid,xfo_to as tag_id,xfo_fo as file_id,xfo_state as state,xfo_ts_state as state_timestamp FROM xfos WHERE xfo_to = '{0}' AND xfo_fo = '{1}'" -f $TagRecord.tag_id,$TagRecord.folder_id
        $CheckResult = Invoke-SqliteQuery -DataSource $Database -Query $Check
        if($Null -ne $CheckREsult){   
            Write-Warning "Tag and Folder Combination Already Exists"
            return $CheckResult
            break
        }else{
            Invoke-SqliteQuery -DataSource $Database -Query $Insert -ErrorAction STOP
            Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery

            return Get-T4WFolderTag -GUID $TagRecord.guid
        }

    }
}

function Remove-T4WFolderTag{
    <#
    .SYNOPSIS
    Removes Folder Tag
    
    .DESCRIPTION
    Copies duplicate to Archive table and deletes original Record
    
    .PARAMETER GUID
    Guid for this record
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Remove-T4WFolderTag -GUID '{5df21138-cb08-4bb5-93f3-d6a5ccf4f3fb}'
    

    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [GUID]$GUID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath
    )

    process {
        $TAG = Get-T4WFolderTag -GUID $GUID
        if($Null -eq $TAG){WRite-Error -ErrorAction Stop -Message "Unable to Retrieve Folder Tag, Verify this is correct with Get-T4WFolderTag"}


        [int64]$Timestamp = (get-date).ToFileTime()

        $INSERT =  " INSERT INTO dxfos ( dxfo_guid, dxfo_to, dxfo_fo, dxfo_state, dxfo_ts) "
        $INSERT += " VALUES ('{0}','{1}','{2}','{3}','{4}');" 
        $INSERT = $INSERT -f $TAG.guid, `
                             $TAG.tag_id,`
                             $TAG.Folder_id, `
                             $TAG.state,`
                             $TAG.state_timestamp

        Write-Verbose "Deletion Record Query: $Insert"

        $DeleteQuery = "DELETE FROM xfos WHERE xfo_guid = '{0}' " -f $TAG.guid
        Write-Verbose "Deletion Query: $DeleteQuery"

        $UpdateInfoQuery = "Update info SET info_value_i = '{0}' WHERE info_key = '-7001';" -f $Timestamp
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        Invoke-SQLiteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SQLiteQuery -DataSource $Database -Query $DeleteQuery -ErrorAction Continue
        Invoke-SQLiteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction SilentlyContinue
    }
}
#endregion


#region FileTag
function Get-T4WFileTag {
    <#
    .SYNOPSIS
    Returns File Tags
    
    .DESCRIPTION
    Returns all the file tags associated with a file or a specific unique ID
    
    .PARAMETER FilePath
    File Path
    
    .PARAMETER FileID
    File ID
    
    .PARAMETER FileGuid
    File GUID
    
    .PARAMETER GUID
    FileTag Guid
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Get-T4WFileTag -fileID 216                                       

guid            : {7032e0c1-7d46-483e-860b-01a3efb2a7d1}
tag_id          : 253
File_id         : 216
state           : 0
state_timestamp : 133883921804366664
state_time      : 4/6/2025 1:49:40 AM

guid            : {08f047b3-4906-4c73-8deb-1dd356c48522}
tag_id          : 470
File_id         : 216
state           : 0
state_timestamp : 133885420657718534
state_time      : 4/7/2025 7:27:45 PM

guid            : {23698727-50ff-4ad9-b701-315765ed513b}
tag_id          : 498
File_id         : 216
state           : 0
state_timestamp : 133885420657728537
state_time      : 4/7/2025 7:27:45 PM

guid            : {cec28836-c9a5-4410-9b86-5847a2c1852e}
tag_id          : 499
File_id         : 216
state           : 0
state_timestamp : 133885420657728537
state_time      : 4/7/2025 7:27:45 PM

guid            : {e5c48173-0107-4581-8dec-f33a0e1a8114}
tag_id          : 658
File_id         : 216
state           : 0
state_timestamp : 133885405963082767
state_time      : 4/7/2025 7:03:16 PM
    

    #>
    [cmdletbinding(DefaultParameterSetName = 'Default')]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath")]
        [alias("FileName","Name","Fullname","Path")]
        [string]$FilePath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileID")]
        [int]$FileID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileGUID")]
        [GUID]$FileGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="LINKGUID")]
        [GUID]$GUID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )
    process  {
        $Query = ""
        switch($PsCmdlet.ParameterSetName){
            'FilePath' {
                $File = Get-T4wFile -FilePath $FilePath
                if($Null -eq $File){
                    Write-Error -ErrorAction Stop -Message "Invalid File Supplied"
                }
                
                $Query = "SELECT * from xfis WHERE xfi_fi = '{0}'" -f $File.id

                break
            }
            'FileID'   {
                $File = Get-T4wFile -FileID $FileID
                if($Null -eq $File){
                    Write-Error -ErrorAction Stop -Message "Invalid File ID Supplied"
                }
                
                $Query = "SELECT * from xfis WHERE xfi_fi = '{0}'" -f $File.id

                break
            }
            'FileGUID' {
                $File = Get-T4wFile -FileGuid $FileGuid
                if($Null -eq $File){
                    Write-Error -ErrorAction Stop -Message "Invalid File Supplied"
                }
                
                $Query = "SELECT * from xfis WHERE xfi_fi = '{0}'" -f $File.id

                break
            }
            'LINKGUID' {
                
                $Query = "SELECT * from xfis WHERE xfi_guid = '{0}'" -f "{$($GUID.Guid)}"

                break
            }
            'Default' {
                $Query = "SELECT * from xfis"
                break
            }
        }

        Write-Verbose "Query: $Query"
        $Results = Invoke-SqliteQuery -DataSource $Database -Query $Query

        if($Null -ne $Results){
            return @(foreach($Result in $Results){
                [PSCustomObject]@{
                    guid = $Result.xfi_guid
                    tag_id = $Result.xfi_to
                    File_id = $Result.xfi_fi
                    state  =    $Result.xfi_state
                    state_timestamp = $Result.xfi_ts_state
                    state_time =       $(try{[datetime]::FromFileTime($Result.xfi_ts_state)}catch{""})

                }
            })
        }else{
            write-warning "Nothing collected"
            return $Null
        }
    }
}    

function Add-T4WFileTag{
    <#
    .SYNOPSIS
    Adds a File Tag
    
    .DESCRIPTION
    Adds a File Tag
    
    .PARAMETER FilePath
    File Path
    
    .PARAMETER FileID
    File ID
    
    .PARAMETER FileGuid
    File GUID
    
    .PARAMETER TagID
    Tag ID
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
Add-T4WFileTag -Path 'G:\kb16b-02 V2.json2'  -TagID 386
WARNING: Tag and File Combination Already Exists

xfi_guid                               xfi_to xfi_fi
--------                               ------ ------
{7e3cf0cb-3e29-443b-bf95-7e72ff81d327}    386   4760
    
    .NOTES
    Will prevent the creation of duplicate tags.
    #>
    [Alias("Add-FileTag")]
    [cmdletbinding(DefaultParameterSetName = 'FileName')]
    param(

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FilePath")]
        [alias("FileName","Name","Fullname","Path")]
        [string]$FilePath,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileID")]
        [alias("ID","fis_id")]
        [int]$FileID,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="FileGUID")]
        [alias("GUID","fis_guid")]
        [GUID]$FileGuid,

        [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
        [int]$TagID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath

    )

    begin {
        #Get the Timestamp fir updating records
        [int64]$TimeStamp = (get-date).ToFileTime()
        

        
        #CHECK the Guid fir a collision
        $GUID = Create-T4WGUID -Type File_Tag

        $TagRecord = [PSCustomObject]@{
            guid = "{$($GUID)}"
            tag_id = -1
            File_id = -1
            state = 0
            state_time = $TimeStamp

        }



        switch($PsCmdlet.ParameterSetName){
            'FilePath'{ 
                $File = Get-T4WFile -Path $FilePath
                if($Null -eq $File){
                    $File = Add-T4WFile -FilePath $FilePath
                }

                $TagRecord.File_id = $File.id
                break
            }
            'FileGUID'{ 
                $File = Get-T4WFile -FileGuid $FileGuid
                if($Null -eq $File) { Write-Error -ErrorAction Stop -Message "FileGUID is not Present in the database.  GUID:'$FileGUID'"}
                $TagRecord.File_id = $File.id

                break
            }
            'FileID'{ 
                $File = Get-T4WFile -FileID $FileID
                if($Null -eq $File) { Write-Error -ErrorAction Stop -Message "FileID is not Present in the database.  GUID:'$FileID'"}
                $TagRecord.File_id = $File.id
            }
        }

        if($TagRecord.File_id -eq -1){
            Write-Error -ErrorAction Stop -Message "Was Unable to find the File"
        }

        $TAGInfo = get-T4WTags -ID $TagID
        if($null -eq $TAGInfo){
            Write-Error -ErrorAction STOP -Message 'TagID is invalid'
        }else{
            $TagRecord.tag_id = $TAGInfo.id
            Write-Verbose "TAG Name: '$($TAGInfo.Name)'"
        }

        

        $UpdateInfoQuery = "Update info SET info_value_i = '$TimeStamp' WHERE info_key = '-7001';"
        write-verbose "Info Table Update Command: `r`n`t $UpdateInfoQuery"

        $INSERTQuery =  " INSERT INTO xfis (xfi_guid,xfi_to,xfi_fi,xfi_state,xfi_ts_state) "
        $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}');"

        $Insert = $INSERTQuery  -f $TagRecord.guid,$TagRecord.tag_id,$TagRecord.File_id,$TagRecord.state,$TagRecord.state_time
        Write-verbose "INSERT Statement: $Insert"

        $Check = "SELECT xfi_guid as guid,xfi_to as tag_id,xfi_fi as File_id FROM xfis WHERE xfi_to = '{0}' AND xfi_fi = '{1}'" -f $TagRecord.tag_id,$TagRecord.File_id
        $CheckResult = $(Invoke-SqliteQuery -DataSource $Database -Query $Check)
        if($Null -ne $CheckResult){   
            Write-Warning "Tag and File Combination Already Exists"
            return $CheckResult
            break
        }else{
            Invoke-SqliteQuery -DataSource $Database -Query $Insert -ErrorAction STOP
            Invoke-SqliteQuery -DataSource $Database -Query $UpdateInfoQuery
            
        }
    }
}

function Remove-T4WFileTag{
    <#
    .SYNOPSIS
    Remove File Tag
    
    .DESCRIPTION
    Makes duplicate record in the Archive Table and Deletes the original record.
    
    .PARAMETER GUID
    GUID of the FileTag
    
    .PARAMETER Database
    Database Path
    
    .EXAMPLE
    Remove-T4WFileTag -GUID '{7e3cf0cb-3e29-443b-bf95-7e72ff81d327}'
    
    #>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True)]
        [GUID]$GUID,

        [Parameter(Mandatory=$False)]
        [string]$Database=$TagDBPath
    )

    process {
        $TAG = Get-T4WFileTag -GUID $GUID
        if($Null -eq $TAG){WRite-Error -ErrorAction Stop -Message "Unable to Retrieve File Tag, Verify this is correct with Get-T4WFileTag"}


        [int64]$Timestamp = (get-date).ToFileTime()

        $INSERT =  " INSERT INTO dxfis ( dxfi_guid, dxfi_to, dxfi_fi, dxfi_state, dxfi_ts) "
        $INSERT += " VALUES ('{0}','{1}','{2}','{3}','{4}');" 
        $INSERT = $INSERT -f $TAG.guid, `
                             $TAG.tag_id,`
                             $TAG.File_id, `
                             $TAG.state,`
                             $TAG.state_timestamp

        Write-Verbose "Deletion Record Query: $Insert"

        $DeleteQuery = "DELETE FROM xfis WHERE xfi_guid = '{0}' " -f $TAG.guid
        Write-Verbose "Deletion Query: $DeleteQuery"

        $UpdateInfoQuery = "Update info SET info_value_i = '{0}' WHERE info_key = '-7001';" -f $Timestamp
        write-verbose "Info Table Update Command:  $UpdateInfoQuery"

        Invoke-SQLiteQuery -DataSource $Database -Query $INSERT -ErrorAction STOP
        Invoke-SQLiteQuery -DataSource $Database -Query $DeleteQuery -ErrorAction Continue
        Invoke-SQLiteQuery -DataSource $Database -Query $UpdateInfoQuery -ErrorAction SilentlyContinue
    }
}
#endregion



#region Auto_Tagging_DB

Function Get-AutoTagFolderGroup{
    [cmdletbinding(DefaultParameterSetName = 'Default')]
    param(

       [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="GroupID")]
       [int] $FolderGroupID,
       
       [Parameter(Mandatory=$True,ValueFromPipeline=$true,ParameterSetName="GroupName")]
       [string] $FolderGroupName,
       
        [Parameter(Mandatory=$False)]
        [string]$Database=$RulesDBPath
    )
process{
    $Query = ""
    switch($PsCmdlet.ParameterSetName){
        'GroupName' { 
            $Query =  " SELECT adb_fg_id as Group_ID, adb_fg_name as Group_Name, adb_fg_folders as Group_Folders, "
            $Query += " adb_fg_filters as Group_Filters , adb_fg_timestamp as Group_Timestamp, adb_fg_process as Group_Process, adb_fg_state as Group_State "
            $Query += " FROM adb_fgs "
            $Query += " WHERE Group_Name LIKE '%$($FolderGroupName)%';"
        }
        'GroupID' { 
            $Query =  " SELECT adb_fg_id as Group_ID, adb_fg_name as Group_Name, adb_fg_folders as Group_Folders, "
            $Query += " adb_fg_filters as Group_Filters , adb_fg_timestamp as Group_Timestamp, adb_fg_process as Group_Process, adb_fg_state as Group_State "
            $Query += " FROM adb_fgs "
            $Query += " WHERE Group_ID = '$FolderGroupID';"
        }        
        default  {
            $Query =  " SELECT adb_fg_id as Group_ID, adb_fg_name as Group_Name, adb_fg_folders as Group_Folders, "
            $Query += " adb_fg_filters as Group_Filters , adb_fg_timestamp as Group_Timestamp, adb_fg_process as Group_Process, adb_fg_state as Group_State "
            $Query += " FROM adb_fgs ;"
        }
    }

    Invoke-SqliteQuery -DataSource $Database -Query $Query
    <#
    SELECT adb_fg_id, adb_fg_name, adb_fg_folders, adb_fg_filters, adb_fg_timestamp, adb_fg_process, adb_fg_state FROM adb_fgs;    
    #>
}    

}

Function Add-AutoTagFolderGroup{
    [cmdletbinding()]
    param(
       
       [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
       [string] $FolderGroupName,

       [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
       [AutoTagging_State]$Schedule,

       [Parameter(Mandatory=$False,ValueFromPipeline=$true)]
       [switch]$Disabled,
       [Parameter(Mandatory=$False,ValueFromPipeline=$true)]
       [switch]$NoMonitor,

       [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
       [AutoTagging_SearchIn]$SearchIn,

       [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
       [string[]]$Folders,

       [Parameter(Mandatory=$True,ValueFromPipeline=$true)]
       [string[]]$FolderGroupFilters,
         
        [Parameter(Mandatory=$False)]
        [string]$Database=$RulesDBPath
    )
process{
    [string] $Filters = $FolderGroupFilters -join "`r`n"
    $Filters = $Filters -replace "'","''"

    [string] $Folders = $Folders -join "`r`n"
    $Folders = $Folders -replace "'","''"

    $FolderGroupName = $FolderGroupName -replace "'","''"


    $StateSchedule = [int]$Schedule
    if(!($Disabled)){
        $StateSchedule+=1
    }
    if(!($NoMonitor)){
        $StateSchedule+=2
    }    

    $INSERTQuery =  " INSERT INTO adb_fgs ( adb_fg_name,adb_fg_folders,adb_fg_filters,adb_fg_process,adb_fg_state ) "
    $INSERTQuery += " VALUES ('{0}','{1}','{2}','{3}','{4}');"
    $Insert = $INSERTQuery -f $FolderGroupName, $Folders,$Filters,[int]$SearchIn,$StateSchedule

    
    write-verbose "INSERT QUERY: $Insert"

    Invoke-SqliteQuery -DataSource $Database -Query $Insert

    
    <#
    $Query = ""
    switch($PsCmdlet.ParameterSetName){
        'GroupName' { 
            $Query =  " SELECT adb_fg_id as Group_ID, adb_fg_name as Group_Name, adb_fg_folders as Group_Folders, "
            $Query += " adb_fg_filters as Group_Filters , adb_fg_timestamp as Group_Timestamp, adb_fg_process as Group_Process, adb_fg_state as Group_State "
            $Query += " FROM adb_fgs "
            $Query += " WHERE Group_Name LIKE '%$($FolderGroupName)%';"
        }
        'GroupID' { 
            $Query =  " SELECT adb_fg_id as Group_ID, adb_fg_name as Group_Name, adb_fg_folders as Group_Folders, "
            $Query += " adb_fg_filters as Group_Filters , adb_fg_timestamp as Group_Timestamp, adb_fg_process as Group_Process, adb_fg_state as Group_State "
            $Query += " FROM adb_fgs "
            $Query += " WHERE Group_ID = '$FolderGroupID';"
        }        
        default  {
            $Query =  " SELECT adb_fg_id as Group_ID, adb_fg_name as Group_Name, adb_fg_folders as Group_Folders, "
            $Query += " adb_fg_filters as Group_Filters , adb_fg_timestamp as Group_Timestamp, adb_fg_process as Group_Process, adb_fg_state as Group_State "
            $Query += " FROM adb_fgs ;"
        }
    }

    Invoke-SqliteQuery -DataSource $Database -Query $Query
    <#
    SELECT adb_fg_id, adb_fg_name, adb_fg_folders, adb_fg_filters, adb_fg_timestamp, adb_fg_process, adb_fg_state FROM adb_fgs;    
    #>
}    

}


Function Move-File {
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,Position=0)]
        [string[]]$Path,
        [Parameter(Mandatory=$True,ValueFromPipeline=$true,Position=1)]
        [string]$Destination,
        [switch]$Force
    )

    process{
        $Items = Get-Item $Path
        if($Null -eq $Items){
            Write-Error -ErrorAction Stop -Message "File/Folder Not Found"
        }

        foreach($Item in $Items){

            if($Item.PSIsContainer -eq $true){

                $FolderFullname = if($Item.FullName -notmatch "\\$"){"$($Item.FullName)\"}else{$Item.FullName}
                $T4WFoundFolders = Get-T4WFolder -FolderPath "$FolderFullname"


                if($Null -eq $T4WFoundFolders){
                    Write-Verbose "File is not referenced in database, No action is needed"
                }else{
                    $MovedItems = Move-Item -Path $Item.FullName -Destination $Destination -Force:$Force -ErrorAction Stop -PassThru

                    foreach($MovedItem in $MovedItems){
                        $movedPath = Get-T4WPath -Path $MovedItem.fullname
                        if($MovedPath.driveletter -ne ''){
                            $Volume = Get-T4WVolumes -Letter $MovedPath.DriveLetter
                        }elseif($MovedPath.volume -ne ''){
                            $Volume = Get-T4WVolumes -GUID $MovedPath.volume
                        }elseif($MovedPath.server -ne '' -and $MovedPath.share -ne ''){
                            $Volume = Get-T4WVolumes -Machine $movedPath.server -Share $movedPath.share
                        }else{
                            $Volume = $Null
                        }

                        $ParentID=Get-T4WVolumes -Letter $movedPath.driveletter
                        if($Null -eq $ParentID){
                            $ParentID=Get-T4WVolumes -GUID $movedPath.volume
                        }
                        if($Null -eq $ParentID){
                            $ParentID=Get-T4WVolumes -Machine $movedPath.Server -Share $movedPath.share
                        }

                        if($Null -ne $ParentID){
                            Update-T4WFolder -id $T4WFoundFolders.id -NewRelativePath $movedPath.RelativePath -NewParent $ParentID.id | Out-Null
                        }
                    }
                }
            }else{
                
                $FileFullname = $Item.FullName
                $T4WFoundFiles = Get-T4WFile -FilePath "$FileFullname" 
                
                if($Null -eq $T4WFoundFiles){
                    Write-Verbose "File is not referenced in database, No action is needed"
                }else{
                    $MovedItems = Move-Item -Path $Item.FullName -Destination $Destination -Force:$Force -ErrorAction Stop -PassThru

                    foreach($MovedItem in $MovedItems){
                        $movedPath = Get-T4WPath -Path $MovedItem.fullname
                        if($MovedPath.driveletter -ne ''){
                            $Volume = Get-T4WVolumes -Letter $MovedPath.DriveLetter
                        }elseif($MovedPath.volume -ne ''){
                            $Volume = Get-T4WVolumes -GUID $MovedPath.volume
                        }elseif($MovedPath.server -ne '' -and $MovedPath.share -ne ''){
                            $Volume = Get-T4WVolumes -Machine $movedPath.server -Share $movedPath.share
                        }else{
                            $Volume = $Null
                        }

                        $ParentID=Get-T4WVolumes -Letter $movedPath.driveletter
                        if($Null -eq $ParentID){
                            $ParentID=Get-T4WVolumes -GUID $movedPath.volume
                        }
                        if($Null -eq $ParentID){
                            $ParentID=Get-T4WVolumes -Machine $movedPath.Server -Share $movedPath.share
                        }

                        if($Null -ne $ParentID){
                            Update-T4WFile -id $T4wFoundfiles.id -NewRelativePath $movedPath.RelativePath -NewParent $ParentID.id | Out-Null
                        }
                    }
                }
            }
        }
    }
}

#endregion

################
###   TODO   ###
################

# Add Following
# Remove-T4WTags

##  Determine if the following is needing to be added.
# Update-T4WFileTag       ###NOTHING REALLY TO CHANGE with State and I don't know what correct states are
# Update-T4WFolderTag     ###NOTHING REALLY TO CHANGE with State and I don't know what correct states are

######
#  Note:  Will probably never add remove-volume, I don't see the need in it, and it has far reaching issues if done incorrectly.  Invalidating entire drives worth of tags.


#Add Functions/Methods for dealing with Auto Tool