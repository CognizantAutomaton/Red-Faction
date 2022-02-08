Add-Type -AssemblyName System.Windows.Forms
<#

usage:
For example, to find VPP files containing the text "L19S3", use the following command:

& "C:\Path\to\script\VPP_searcher.ps1" L19S3

or

cd "C:\Path\to\script"
.\VPP_searcher.ps1 L19S3

#>

[string]$RF_path = [string]::Empty

#region attempt to retrieve RF root folder from registry
[Array]$RegistryPaths = @(
    @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Volition\Red Faction"; AccessKey = "InstallPath" }, # legacy 64-bit
    @{ Path = "HKLM:\SOFTWARE\Volition\Red Faction"; AccessKey = "InstallPath" }, # legacy 32-bit
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 20530"; AccessKey = "InstallLocation" }, # steam
    @{ Path = "HKLM:\SOFTWARE\GOG.com\GOGREDFACTION"; AccessKey = "PATH" } # GOG
)

foreach ($r in $RegistryPaths) {
    if (Test-Path $r.Path) {
        $RF_path = (Get-ItemProperty -Path $r.Path)."$($r.AccessKey)"
        if ($RF_path.Length -gt 0) { break }
    }
}
#endregion

if (($RF_path.Length -eq 0) -or (-not (Test-Path -LiteralPath $RF_path))) {
    [System.Windows.Forms.MessageBox]::Show("Unable to find RF folder.`nYou will be prompted to browse for the path yourself.", "Unable to find default RF folder", "OK", "Information")

    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = "Browse to RF root folder"
    }

    if ($fbd.ShowDialog() -eq "OK") {
        $RF_path = $fbd.SelectedPath
    }

    if ($RF_path.Length -eq 0) {
        Write-Host "Unable to run the program unless you select your RF root folder"
    }
}

[string]$search_text = [string]::Join(" ", $Args).ToUpper()

class VPPPosition {
    static [int]Calculate([int]$inpt) {
        [int]$output = $inpt

        if ($inpt % 2048 -ne 0) {
            $output = $inpt - ($inpt % 2048)
            $output += 2048
        }

        return $output
    }
}

if (Test-Path -LiteralPath $RF_path) {
    if ($search_text.Trim().Length -eq 0) {
        $search_text = Read-Host "Enter your search text"
    }

    [System.Collections.Generic.List[HashTable]]$match_table = New-Object System.Collections.Generic.List[HashTable]

    # loop thru the list of VPP files in the RF folder
    Get-ChildItem *.vpp -LiteralPath $RF_path -File -Recurse | ForEach-Object {
        # retrieve full path to VPP file
        [string]$VPP_Filename = $_.FullName

        # read VPP data into memory
        [byte[]]$bytes = [System.IO.File]::ReadAllBytes($VPP_Filename)

        # read VPP header information
        [int]$signature = [System.BitConverter]::ToInt32($bytes, 0)
        [int]$version = [System.BitConverter]::ToInt32($bytes, 4)

        # check that VPP header is well-formed
        if (($signature -eq 0x51890ace) -and ($version -eq 1)) {
            # read file count in VPP (byte positions 8-11)
            [int]$num_files = [System.BitConverter]::ToInt32($bytes, 8)

            [int]$prev_position = 0
            [int]$prev_size = 0

            for ([int]$n = 0; $n -lt $num_files; $n++) {
                # filename in first 60 bytes of current position, replace null characters
                [string]$filename = [System.Text.Encoding]::ASCII.GetString($bytes, 2048 + $n*64, 60) -replace '\x00',''

                # file size in next 4 bytes
                [int]$filesize = [System.BitConverter]::ToInt32($bytes, 2048 + $n*64 + 60)

                # calculate position to raw file data
                [int]$curr_position = [VPPPosition]::Calculate($(if ($n -eq 0) {2048 + 64 * $num_files} else {$prev_position + $prev_size}))

                # if filename contains search text
                if ($filename.ToLower().Contains($search_text.ToLower()) -or ($filename -like $search_text)) {
                    $match_table.Add(@{
                        Filename = $filename
                        Size = $filesize
                        Position = $curr_position
                        VPP_Filename = $VPP_Filename
                    })
                    # print a match
                    Write-Host "$($match_table.Count). '$filename', $([math]::Round($filesize / 1KB, 2)) KiB in '$VPP_Filename'" -ForegroundColor Green
                }

                # track previous size and position
                $prev_size = $filesize
                $prev_position = $curr_position
            }
        }
    }

    # if matches found
    if ($match_table.Count -gt 0) {
        [bool]$prompt_extract = $true
        [int]$extract_count = 0

        # prompt for save
        [string]$file_to_extract = [string]::Empty

        while ($prompt_extract) {
            if ($extract_count -eq 0) {
                $file_to_extract = Read-Host "Would you like to extract any of the results? Type the number to extract, or X to quit"
            } else {
                $file_to_extract = Read-Host "Type the number to extract another, or X to quit"
            }

            if ((-not $file_to_extract.StartsWith("0")) -and ($file_to_extract -match "^[0-9]+$")) {
                # retrieve file record to save
                $filenum = [int]$file_to_extract

                if ($filenum -le $match_table.Count) {
                    [HashTable]$file_record = $match_table[$filenum - 1]

                    # read VPP into memory
                    [byte[]]$vpp_bytes = [System.IO.File]::ReadAllBytes($file_record.VPP_Filename)

                    # retrieve file data from VPP
                    [byte[]]$file_data = $vpp_bytes[($file_record.Position)..($file_record.Position + $file_record.Size - 1)]

                    # determine file extension
                    [int]$extension_idx = $file_record.Filename.LastIndexOf(".")
                    [string]$extension = [string]::Empty

                    if ($extension_idx -gt 0) {
                        $extension = $file_record.Filename.Substring($extension_idx)
                    }

                    # initialize save dialog
                    $FileDialog = New-Object System.Windows.Forms.SaveFileDialog -Property @{
                        Filter = @(if ($extension.Length -gt 0) {"$($extension.Substring(1).ToUpper()) files (*$extension)|*$extension"} else {"All files (*.*)|*.*"})
                        Filename = $file_record.Filename
                    }

                    if ($FileDialog.ShowDialog() -eq 'OK') {
                        # write the file to disk
                        [System.IO.File]::WriteAllBytes($FileDialog.FileName, $file_data)
                        Write-Host "Saved '$($FileDialog.FileName)'"
                        $extract_count++
                    }
                } else {
                    Write-Host "There are only $($match_table.Count) records listed above. Try again." -ForegroundColor Red
                }
            } else {
                $prompt_extract = $false
                Write-Host "Quitting"
            }
        }
    }
} else {
    Write-Host "'$RF_path' does not exist!`nBe sure to update the variable `$RF_path to your actual RF folder path." -ForegroundColor Yellow
}
