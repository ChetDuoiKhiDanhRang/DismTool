class WimMountInfo {
    [string]$MountDir
    [string]$ImageFile
    [int]$Index
    [string]$ProductName
    [string]$Version
    [string]$Status
    [string]$Arch
}

class WimIndexInfo {
    [int]$Index
    [string]$Name
}

class feature{
    [string]$Name
    [string]$State
}

class driver{
    [string]$PublishName
    [string]$OriginalFileName
    [String]$Inbox
    [String]$ClassName
    [String]$Provider
    [String]$Version
}

class appx {
    [string]$DisplayName
    [string]$PackageName
    [string]$Version
}

class commulativePackage {
    [string]$PackageIdentity
    [string]$State
    [string]$ReleaseType
    [datetime]$InstallTime
}

function FilesPicker {
    param (
    [bool] $MultiSelect=$false,
	[string] $Filter = 'All files (*.*)|*.*|Text files (*.txt)|*.txt',
    [string] $Title = "Select file(s)"
    )
    Add-Type -AssemblyName System.Windows.Forms
    #$currentFolder = $PSScriptRoot

    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    #$openFileDialog.InitialDirectory = $currentFolder
    $openFileDialog.Title = $Title
    $openFileDialog.Filter = $Filter
    $openFileDialog.AddExtension = $false
    $openFileDialog.CheckFileExists = $true
    $openFileDialog.CheckPathExists = $true
    $openFileDialog.DereferenceLinks = $false
    $openFileDialog.ShowHelp = $false
    $openFileDialog.Multiselect = $MultiSelect
    $openFileDialog.ValidateNames = $false
    $openFileDialog.RestoreDirectory = $true
    

    if ($openFileDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK)
    {
        $selectedFiles = $openFileDialog.FileNames
        return $selectedFiles
    }
    else
    {
        return $null
    }
}

function FolderPicker {
    param (
        [string]$Title = "Select folder"
    )
    Add-Type -AssemblyName System.Windows.Forms

    $dialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $dialog.Description = $Title    

    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $selectedPath = $dialog.SelectedPath
        return $selectedPath
    } else {
        return null
    }
}

function PrintMenu {
    param (
        [string[]] $Strings
    )

    Write-Host "==================================="
    for ($i = 1; $i -lt $Strings.Length; $i++) {
        Write-Host ("{0:D1}. {1}" -f $i.ToString().PadRight(2), $Strings[$i])
    }
    Write-Host "    -----------------"
    Write-Host ("{0:D1}. {1}" -f "0 ", $Strings[0])
    Write-Host "==================================="
    Write-Host "Input your choice and press ENTER: " -NoNewline -Foreground Cyan
}

function ListAllMountPoints{
    $output = dism /Get-MountedWimInfo | Out-String
    $blocks = $output -split "(?m)^\s*$"

    [WimMountInfo[]]$list = @()

    foreach ($block in $blocks) {
        if ($block -match "Mount Dir\s*:\s*(.+)") {

            # tạo object trống
            $obj = [WimMountInfo]::new()
            
            
            # gán từng thuộc tính
            if ($block -match "Mount Dir\s*:\s*(.+)")    { $obj.MountDir  = $matches[1].Trim() }
            if ($block -match "Image File\s*:\s*(.+)")   { $obj.ImageFile = $matches[1].Trim() }
            if ($block -match "Image Index\s*:\s*(\d+)") { $obj.Index     = [int]$matches[1] }
            if ($block -match "Status\s*:\s*(.+)") { $obj.Status = $matches[1].Trim() }

            $info = dism /Get-WimInfo /WimFile:"$($obj.ImageFile)" /index:$($obj.Index) | Select-Object -Skip 3
            foreach ($i in $info){
                if ($i -match "Name\s*:\s*(.+)") {$obj.ProductName = $matches[1].Trim()}
                if ($i -match "Version\s*:\s*(.+)") {$obj.Version= $matches[1].Trim()}
                if ($i -match "Architecture\s*:\s*(.+)") {$obj.Arch= $matches[1].Trim()}
            }

            # thêm object vào danh sách
            $list += $obj
        }
    }

    return $list
}

function WimAndMountPoint {
    $SelectedItem =""
    [string]$WimFile =""
    [string]$MountDir=""
    [int]$SelectIndex=-1
    [string]$SelectName=""

    $MenuItems =@("To Main menu...", #0
    "Mount WIM file and index", #1
    "Select mounted point`n", #2
    "Listing/Uninstall 3rd drivers", #3
    "Listing/Uninstall Provisioned appx packages", #4
    "Listing/Enable/Disable Features", #5
    "Listing/Uninstall commulative updated`n", #6
    "Install drivers from directory", #7
    "Install appx and depenency packages (appx, msix, appxbundle, msixbundle) from directory", #8
    "Install commulative packages from directory`n", #9
    "Unmount (Commit or Discard)" #10
    )

    while ($SelectedItem -ne 0) {
        cls
        if ($WimFile -ne "") {
            Write-Host ("{0}: {1}" -f "Wim file".PadRight(13),$WimFile) -ForegroundColor Green
            Write-Host ("{0}: {1} ({2})" -f "Index".PadRight(13),$SelectIndex,$SelectName) -ForegroundColor Green
        } else {
            Write-Host ("{0}: " -f "Wim file".PadRight(13)) -ForegroundColor Green
            Write-Host ("{0}: " -f "Index".PadRight(13)) -ForegroundColor Green
        }
        if (($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
            Write-Host ("{0}: {1}" -f "Mounted point".PadRight(13),$MountDir) -ForegroundColor Green
        } else {
            Write-Host ("{0}: " -f "Mounted point".PadRight(13)) -ForegroundColor Green
        }

        PrintMenu $MenuItems
        $SelectedItem = Read-Host
        switch ($SelectedItem) {
            #Select WIM file and Index
            1 {
                $WimFile = FilesPicker -MultiSelect $false -Filter "WIM file (*.wim)|*.wim" -Title "Select WIM file"
                if ($WimFile -eq $null) {break}
                $SelectIndex = -1
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                if (Test-Path -Path $WimFile) {
                    $WimInfo = dism /Get-WimInfo /WimFile:"$($WimFile)" | Out-String    
                    $info = $WimInfo -split "(?m)^\s*$"
                    [WimIndexInfo[]]$list = @()
                    
                    foreach ($i in $info) {
                        if ($i -match "Index\s*:\s*(.+)"){
                            $obj = [WimIndexInfo]::new()
                            if ($i -match "Name\s*:\s*(.+)")    { $obj.Name  = $matches[1].Trim() }
                            if ($i -match "Index\s*:\s*(.+)")   { $obj.Index  = [int]$matches[1].Trim() }
                            $list += $obj
                        }
                    }

                    if ($list.Count -gt 0) {
                        while (-not ((($SelectIndex -as [int]) -gt 0) -and (($SelectIndex -as [int]) -le $list.Length))) {
                            cls
                            Write-Host $WimFile
                            Write-Host Index list:
                            $list | Format-Table
                            Write-Host -NoNewline "Input index and Enter: "
                            $SelectIndex = Read-Host
                        }
                        $SelectName = $($list[$SelectIndex-1].Name)
                    }                    
                    
                    #chose mount folder
                    Write-Host "Select folder to mount wim/index..." -ForegroundColor Cyan
                    $tMountDir = FolderPicker
                    if ($tMountDir -eq $null) {break}
                    $MountDir = $tMountDir
                    [WimMountInfo[]]$listMounted = ListAllMountPoints
                    
                    if ($listMounted.Count -gt 0){
                        $mountedPoint = $listMounted | Where-Object {$_.MountDir.Contains($MountDir)}
                    } else {
                        $mountedPoint =$null
                    }

                    if ($mountedPoint -ne $null) {
                        Write-Host "This dir has been used as a mount point.`nCheck list mounted points below:"; 
                        $MountDir =""
                        $WimFile =""
                        $SelectIndex = -1
                        $mountedPoint | Format-Table; 
                        Read-Host; 
                        Break
                    }

                    if ($WimFile.Contains($MountDir)) {
                        Write-Host "This dir is folder or sub folder of WIM file's location, chose another.";
                        $MountDir =""
                        $WimFile =""
                        $SelectIndex = -1
                        Read-Host
                        Break
                    }

                    Write-Host Mount to: $MountDir
                    if (($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                        dism /Mount-Wim /WimFile:"$($WimFile)" /MountDir:"$($MountDir)" /index:$SelectIndex
                        if ($?) {
                            Write-Host ("{0}`nhas been mounted to:`n{1}" -f $WimFile, $MountDir)
                            Read-Host "Press Enter to continue..."
                        }
                    }
                }
                break
            }

            #select mount point
            2 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                [WimMountInfo[]]$listMounted = ListAllMountPoints

                if ($listMounted.Count -gt 0){
                    
                    $selectMountPoint = 0
                    while (-not (($selectMountPoint -gt 0) -and ($selectMountPoint -le $listMounted.Count)))
                    {
                        cls
                        for ($i = 1; $i -le $listMounted.Count; $i++) {
                            Write-Host "Mounted point $($i):"
                            Write-Host " - Path: $($listMounted[$($i)-1].MountDir)"
                            Write-Host " - File: $($listMounted[$($i)-1].ImageFile)"
                            Write-Host " - Product: $($listMounted[$($i)-1].ProductName)"
                            Write-Host
                        }
                    
                        Write-Host "Chose mounted point or input ""e"" to exit: " -NoNewline -ForegroundColor Cyan
                        $selectMountPoint = Read-Host
                        if (($selectMountPoint -gt 0) -and ($selectMountPoint -le $listMounted.Count)) {
                            $MountDir = $listMounted[$selectMountPoint-1].MountDir
                            $WimFile = $listMounted[$selectMountPoint-1].ImageFile
                            $SelectIndex = $listMounted[$selectMountPoint-1].Index
                            $SelectName = $listMounted[$selectMountPoint-1].ProductName
                            
                        } elseif ($selectMountPoint -eq "e") {break}
                    }
                } else { 
                    Write-Host No mounted point!
                    Read-Host
                }
                break
            }

            #List third-party drivers
            3 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                if (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                    $drivers = dism /Image:"$($MountDir)" /get-Drivers | Out-String
                } else {
                    $MountDir = FolderPicker -Title "Select mounted folder:"
                    if (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                        $drivers = dism /Image:"$($MountDir)" /get-Drivers | Out-String
                    } else { break }
                }

                $driversBlocks = $drivers -split "(?m)^\s*$"
                [driver[]]$listDrv = @()
                foreach ($block in $driversBlocks){
                    if ($block -match "Published Name\s*:\s*(.+)") {
                        [driver]$d = [driver]::new()
                        if ($block -match "Published Name\s*:\s*(.+)") {$d.PublishName = $matches[1].Trim()}
                        if ($block -match "Original File Name\s*:\s*(.+)") {$d.OriginalFileName = $matches[1].Trim()}
                        if ($block -match "Inbox\s*:\s*(.+)") {$d.Inbox = $matches[1].Trim()}
                        if ($block -match "Class Name\s*:\s*(.+)") {$d.ClassName = $matches[1].Trim()}
                        if ($block -match "Provider Name\s*:\s*(.+)") {$d.Provider = $matches[1].Trim()}
                        if ($block -match "Version\s*:\s*(.+)") {$d.Version = $matches[1].Trim()}
                        $listDrv += $d
                    } 
                }
                if ($listDrv.Count -eq 0) {Write-Host No 3rd driver found.} else { $listDrv | Format-Table }

                $driverPublishName = ""
                while ($driverPublishName -ne "e") {
                    Write-host "Input driver's published name to uninstall.`nInput ""All"" to clear all 3rd driver, ""e"" to exit.`nInput: " -NoNewline -ForegroundColor Cyan
                    $driverPublishName = Read-Host
                    $driverPublishName = $driverPublishName.Trim()
                    if ($driverPublishName -eq "All") {
                        foreach ($d in $listDrv) {
                            dism /Image:"$($MountDir)" /Remove-Driver /Driver:$($d.PublishName)
                            if ($?) { Write-host $d.PublishName removed. }
                        }
                    } elseif ($driverPublishName -eq "e") { break} else {
                        $driverEntry = $listDrv | Where-Object {$_.PublishName -eq $driverPublishName}
                        if ($driverEntry -ne $null) {
                            dism /Image:"$($MountDir)" /Remove-Driver /Driver:$($driverEntry.PublishName)
                            if ($?) { Write-host $driverEntry.PublishName removed. }
                        }
                    }
                }
                break
            }

            #List ProvisionedAppxPackages
            4 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                if (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                    $appx = dism /Image:"$($MountDir)" /Get-ProvisionedAppxPackages | Out-String
                } else {
                    $MountDir = FolderPicker -Title "Select mounted folder:"
                    if (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                        $appx = dism /Image:"$($MountDir)" /Get-ProvisionedAppxPackages | Out-String
                    } else {break}
                }

                $appxBlocks = $appx -split "(?m)^\s*$"
                [appx[]]$listAppx = @()
                foreach ($block in  $appxBlocks) {
                    if ($block -match "DisplayName\s*:\s*(.+)"){
                        $app = [appx]::new()
                        if ($block -match "DisplayName\s*:\s*(.+)") { $app.DisplayName = $matches[1].trim()}
                        if ($block -match "PackageName\s*:\s*(.+)") { $app.PackageName = $matches[1].trim()}
                        if ($block -match "Version\s*:\s*(.+)") { $app.Version = $matches[1].trim()}
                        $listAppx += $app
                    }
                }
                if ($listAppx.Count -gt 0){
                    $listAppx | Format-List} else { Write-Host No package found.}

                $selectAppxPackageName = ""
                while ($selectAppxPackageName -ne "e"){
                    Write-Host "Input appx package name to remove.`nInput ""All"" to remove all appx. Input ""e"" to exit.`nInput: " -NoNewline -ForegroundColor Cyan
                    $selectAppxPackageName = Read-Host
                    $selectAppxPackageName =$selectAppxPackageName.trim()

                    if ($selectAppxPackageName -eq "All") {
                        foreach ($package in $listAppx){
                            dism /Image:"$($MountDir)" /Remove-ProvisionedAppxPackage /PackageName:$($package.PackageName)
                            if ($?) {Write-Host $package.PackageName removed.}
                        }
                    } elseif (($selectAppxPackageName.Length -gt 0) -and ($selectAppxPackageName -ne "e")) {
                        $package = $listAppx | Where-Object {$_.PackageName -eq $selectAppxPackageName}
                        if ($package -ne $null) {
                            dism /Image:"$($MountDir)" /Remove-ProvisionedAppxPackage /PackageName:$($package.PackageName)
                            if ($?) {Write-Host $package.PackageName removed.}
                        } else { Write-Host "Package not found."}
                    }

                }

                break
            }

            #List features, enable/disable
            5 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                [feature[]]$listF =@()
                
                if (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                    $features = dism /Image:"$($MountDir)" /Get-Features | Out-String
                } else {
                    $MountDir = FolderPicker -Title "Select mounted folder:"
                    if (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir)) {
                        $features = dism /Image:"$($MountDir)" /Get-Features | Out-String
                        Write-Host 
                    } else {break}
                }
                $fearureBlocks = $features -split "(?m)^\s*$"


                foreach ($block in $fearureBlocks) {
                    if ($block -match "Feature Name\s*:\s*(.+)") { 
                        $f = [feature]::new()
                        
                        if ($block -match "Feature Name\s*:\s*(.+)") { $f.Name = $matches[1].trim()}
                        if ($block -match "State\s*:\s*(.+)") { $f.State = $matches[1].trim()}

                        $listF += $f
                    }
                }

                $listF | Format-table

                $featureName = ""
                
                while ($featureName -ne "e") {
                    Write-Host "Input feature name to switch enable/disable. Input ""e"" to exit): " -NoNewline
                    $featureName = Read-Host
                    $featureName = $featureName.Trim()
                    [feature]$featureEntry = $listF | Where-Object {$_.Name -eq $featureName}
                    if ($featureEntry -ne $null) {
                        if($featureEntry.State -eq "Disabled") {
                            dism /Image:"$($MountDir)" /Enable-Feature /FeatureName:$($featureEntry.Name) /All
                            if ($?) {
                                $featureEntry.State = "Enabled"
                                Write-Host $featureEntry.Name enabled! -ForegroundColor Cyan
                                    
                            }
                        } elseif (($featureEntry.State -eq "Disabled with Payload Removed") -and ($featureEntry.Name -eq "netFx3") ){
                            Write-Host "Enable this feature need ISO file.`nSpecify the Windows ISO file..." -ForegroundColor Cyan
                            $ISOfile = FilesPicker -MultiSelect $false -Title "Windows ISO file" -Filter "ISO files (*.iso)|*.iso"
                            if ($ISOfile -ne $null -and (Test-Path -Path $ISOfile)) {
                                EnableNetFx3FromISO -ISOPath "$($ISOfile)" -MountPoint $MountDir
                                $featureEntry.State = "Enabled"
                            } 
                        } elseif ($featureEntry.State -eq "Enabled"){
                            dism /Image:"$($MountDir)" /Disable-Feature /FeatureName:$($featureEntry.Name)
                            if ($?) {
                                $featureEntry.State = "Disable"
                                Write-Host $featureEntry.Name disabled! -ForegroundColor Cyan
                            }
                        }

                    } else { Write-Host Feature not found! }

                }
                break
            }

            #Commulative update
            6 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                if (-not (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir))) {
                    Write-Host "Mounted point not selected. Select mounted point folder..."
                    $MountDir = FolderPicker -Title "Select mounted folder:"
                    if  (($MountDir -eq $null) -or ($MountDir.Length -eq 0)) {break}
                }

                $listPackage = dism /Image:"$($MountDir)" /Get-Packages | Out-String
                $blocks = $listPackage -split "(?m)^\s$"

                [commulativePackage[]]$cPackages = @()
                foreach ($block in $blocks){
                    if ($block -match "Package Identity\s*:\s*(.+)"){
                        [commulativePackage]$package = [commulativePackage]::new()
                        if ($block -match "Package Identity\s*:\s*(.+)") {$package.PackageIdentity = $matches[1].trim()}
                        if ($block -match "State\s*:\s*(.+)") {$package.State = $matches[1].trim()}
                        if ($block -match "Release Type\s*:\s*(.+)") {$package.ReleaseType = $matches[1].trim()}
                        if ($block -match "Install Time\s*:\s*(.+)") {$package.InstallTime = (Get-Date $matches[1])}
                        $package
                        if ($package.InstallTime -gt (Get-Date "2021-12-31")){
                            $cPackages += $package
                        }
                    }
                }

                if ($cPackages.Count -eq 0) {Write-Host "No commulative package after 2021-12-31 installed"} else { $cPackages | Format-Table }

                $packageName = " "
                while ($packageName -ne "e"){
                    Write-Host "Input package name to remove or ""e"" to exit: " -NoNewline -ForegroundColor Cyan
                    $packageName = Read-Host
                    [commulativePackage]$p = $cPackages | Where-Object {$_.PackageIdentity -eq $packageName}
                    if ($p -ne $null) {
                        dism /Image:"$($MountDir)" /Remove-Package /PackageName:$($p.PackageIdentity)
                        if ($?) {Write-Host $p.PackageIdentity has been remove.}

                    } else { Write-Host "$packageName" not found.}
                }



            }
            
            #install driver
            7 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                
                if (-not (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir))) {
                    Write-Host "Mounted point not selected. Select mounted point folder..."
                    $MountDir = FolderPicker -Title "Select mounted folder:"
                    if  (($MountDir -eq $null) -or ($MountDir.Length -eq 0)) {break}
                }

                Write-Host "Mounted point: $($MountDir). Select driver folder..."
                
                [string]$driverFolder = FolderPicker -Title "Select drivers folder"
                if (($driverFolder -eq $null) -or ($driverFolder.Length -eq 0)) { break }
                Write-Host "Drivers folder: $($driverFolder)"
                dism /Image:"$($MountDir)" /Add-Driver /Driver:"$($driverFolder)" /recurse
                if ($?) {Write-Host Done.}
                

                Read-Host
                break
            }

            #Appx and Depenencies
            8 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                
                if (-not (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir))) {
                    Write-Host "Mounted point not selected. Select mounted point folder..."
                    $MountDir = FolderPicker -Title "Select mounted folder:"
                    if  (($MountDir -eq $null) -or ($MountDir.Length -eq 0)) {break}
                }
                Write-Host "Mounted point: $($MountDir). Select appx folder..." -ForegroundColor Cyan
                
                [string]$appxFolder = FolderPicker -Title "Select appx folder..."
                if (($appxFolder -eq $null) -or ($appxFolder.Length -eq 0)) { break }
                Write-Host "Appx folder: $($appxFolder)" -ForegroundColor Cyan

                $packages = Get-ChildItem -Recurse $appxFolder -File | Select-Object -Property Name,FullName | Where-Object {$_.Name -imatch "[*.appx,*.appxbundle,*.msix,*.msixbundle]$"}
                $appxs = @()
                $depends = @()
                foreach ($pf in $packages) {
                    if (Test-PackageIsFramework -PackagePath $pf.FullName) { $depends += $pf } else { $appxs += $pf }
                }
                
                Write-Host "Appx packages:"
                $appxs | Format-Table
                Write-Host "Dependency packages:"
                $depends | Format-Table
                
                foreach ($d in $depends) {
                    Write-Host Install $d.Name: -ForegroundColor Cyan
                    dism /Image:"$($MountDir)" /Add-ProvisionedAppxPackage /PackagePath:"$($d.FullName)" /SkipLicense
                }
                foreach ($d in $appxs) {
                    Write-Host Install $d.Name: -ForegroundColor Cyan
                    dism /Image:"$($MountDir)" /Add-ProvisionedAppxPackage /PackagePath:"$($d.FullName)" /SkipLicense
                }

                Write-Host "Press Enter to continue..." -ForegroundColor Cyan
                Read-Host
                break
            }
            
            #Install commulative update packages
            9 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                $packageFolder = FolderPicker -Title "Select commulative package folder"
                if (($packageFolder -ne $null) -and ($packageFolder.Length -gt 0) -and (Test-Path -Path $packageFolder)){
                    dism /Image:"$($MountDir)" /Add-Package /PackagePath:"$($packageFolder)"
                }

                Read-Host
                break
            }

            #Unmount & Commit or Discard
            10 {
                Write-Host("{0} - {1}" -f $SelectedItem,$MenuItems[$SelectedItem]) -ForegroundColor Cyan
                
                if (-not (($MountDir -ne $null) -and ($MountDir.Length -gt 0) -and (Test-Path -Path $MountDir))) {
                    Write-Host "Mounted point not selected." -ForegroundColor Cyan
                    Read-Host
                    break
                } else {
                    $UnmountMode = "x"
                    while ($UnmountMode -match "[^cde]") {
                        Write-Host "Commit (c)/Discard (d)/Exit (e): " -NoNewline -ForegroundColor Cyan
                        $UnmountMode = Read-Host
                        switch ($UnmountMode) {
                            "c" {
                                dism /Unmount-Wim /MountDir:"$($MountDir)" /Commit
                                if ($?) {
                                    $WimFile = ""
                                    $SelectIndex = -1
                                    $MountDir = ""
                                    dism /Cleanup-Wim
                                    
                                }
                                Read-Host
                                break
                            }
                            "d" {
                                dism /Unmount-Wim /MountDir:"$($MountDir)" /Discard
                                if ($?) {
                                    $WimFile = ""
                                    $SelectIndex = -1
                                    $MountDir = ""
                                    dism /Cleanup-Wim
                                }
                                Read-Host
                                break
                            }
                            "e" {
                                break
                            }
                        }
                    }
                }
            }
        }
    }
}

function EnableNetFx3FromISO {
param (
    [string]$ISOpath,
    [string]$MountPoint
)
    if (Test-Path -Path $ISOpath) {
        Write-Host "Mount ISO file..."
        $result = Mount-DiskImage -ImagePath "$($ISOpath)" -PassThru -ErrorAction Stop
        if ($result -ne $null) {
            $DriveLetter = (Get-Volume -DiskImage $result).DriveLetter
            Write-Host $ISOpath has been mount to $DriveLetter :\
            $sxspath = "$($DriveLetter):\sources\sxs"
            if (Test-Path -Path $sxspath) {
                dism /Image:"$($MountPoint)" /Enable-Feature /FeatureName:NetFx3 /All /Source:$sxspath
                Write-Host "Dismount ISO file..."
                Dismount-DiskImage -ImagePath $ISOpath -ErrorAction SilentlyContinue
                Write-Host "Done."
            } else {
                Write-Host "No SXS folder found!"
            }
        }
    }
}

#ChatGPT hân hạnh tài trợ đoạn code này :)))))
function Test-PackageIsFramework {
    <#
        .SYNOPSIS
            Kiểm tra xem một gói AppX/MSIX có phải là gói framework (dependency) hay không.

        .DESCRIPTION
            Hàm sẽ mở file .appx/.appxbundle/.msix/.msixbundle (ZIP container),
            đọc file AppxManifest.xml bên trong, rồi kiểm tra xem nó có thẻ <Framework>True</Framework> không.

        .PARAMETER PackagePath
            Đường dẫn đầy đủ tới file gói cần kiểm tra.

        .OUTPUTS
            [bool] True nếu là gói framework, False nếu không.

        .EXAMPLE
            Test-PackageIsFramework -PackagePath "C:\Appx\Microsoft.VCLibs.x64.14.00.appx"
    #>
    
    [CmdletBinding()]
    param(
        [Parameter(Mandatory, ValueFromPipeline, ValueFromPipelineByPropertyName)]
        [string]$PackagePath
    )

    begin {
        Add-Type -AssemblyName System.IO.Compression.FileSystem -ErrorAction SilentlyContinue
    }

    process {
        if (-not (Test-Path $PackagePath)) {
            Write-Error "File không tồn tại: $PackagePath"
            return $false
        }

        $tempDir = [System.IO.Path]::Combine([System.IO.Path]::GetTempPath(), [System.IO.Path]::GetRandomFileName())
        try {
            [System.IO.Compression.ZipFile]::ExtractToDirectory($PackagePath, $tempDir)
            $manifestPath = Join-Path $tempDir "AppxManifest.xml"

            if (-not (Test-Path $manifestPath)) {
                Write-Verbose "Không tìm thấy AppxManifest.xml trong gói $PackagePath"
                return $false
            }

            [xml]$manifest = Get-Content -Path $manifestPath -Raw
            $frameworkNode = $manifest.Package.Properties.Framework
            return ($frameworkNode -eq 'true')
        }
        catch {
            Write-Verbose "Không thể đọc gói: $($_.Exception.Message)"
            return $false
        }
        finally {
            if (Test-Path $tempDir) {
                Remove-Item $tempDir -Recurse -Force -ErrorAction SilentlyContinue
            }
        }
    }
}

#--Main---------------------
$SelectedItem = ""
$MenuItems = @("EXIT.", 
    "Modify WIM file and mounted point", 
    "List all mounted points", 
    "Commit & unmount all points", 
    "Discard & unmount all points")

while ($SelectedItem -ne 0)
{
    cls
    #$PSScriptRoot
    Write-Host $PSCommandPath -ForegroundColor Green	
    PrintMenu $MenuItems
    $SelectedItem = Read-Host
    if (($SelectedItem -as [int]) -ne $null) {
        cls
        Write-Host ("{0:D1} - {1}" -f $SelectedItem, $MenuItems[$SelectedItem]) -ForegroundColor Green
        switch ($SelectedItem) {
            1 {
                WimAndMountPoint
                break
            }
            
            #List all mounted points
            2 {
                $moutPoints = ListAllMountPoints
                if ($moutPoints.Count -gt 0){
                    $moutPoints | Format-Table -AutoSize
                } else { Write-Host "No mount point"}
                Write-Host -NoNewline Press ENTER to return to menu...
                Read-Host
                break
            }

            #Commit & Unmount
            3 {
                dism /Cleanup-Wim

                $moutPoints = ListAllMountPoints
                if ($moutPoints.Count -gt 0){
                    foreach ($point in $moutPoints) {
                    Write-Host Unmount: $($point.MountDir)
                    dism /Unmount-Wim /MountDir:"$($point.MountDir)" /Commit
                    }
                } else { Write-Host "No mount point"}
                dism /Cleanup-Wim
                Write-Host -NoNewline Press ENTER to return to menu...
                Read-Host
                break
            }

            #Discard & Unmount
            4 {
                dism /Cleanup-Wim
                $moutPoints = ListAllMountPoints
                if ($moutPoints.Count -gt 0){
                    foreach ($point in $moutPoints) {
                    Write-Host Unmount: $($point.MountDir)
                    dism /Unmount-Wim /MountDir:"$($point.MountDir)" /Discard
                    }
                } else { Write-Host "No mount point"}
                dism /Cleanup-Wim
                Write-Host -NoNewline Press ENTER to return to menu...
                Read-Host
                break
            }

            0 {cls; break}
        }
    }   

}


