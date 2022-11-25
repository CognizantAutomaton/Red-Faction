param([Parameter(Mandatory=$true)][string]$SearchText)

Add-Type -AssemblyName System.Windows.Forms
<#

usage:
For example, to find VPP files containing the text "L19S3", use the following command:

& "C:\Path\to\script\VPP_searcher.ps1" L19S3

or

cd "C:\Path\to\script"
.\VPP_searcher.ps1 L19S3

#>

$ErrorActionPreference = "Stop"
[string]$RF_path = [string]::Empty

#region attempt to retrieve RF root folder from registry
[Array]$RegistryPaths = @(
    @{ Path = "HKLM:\SOFTWARE\WOW6432Node\Volition\Red Faction"; AccessKey = "InstallPath" }, # legacy 64-bit
    @{ Path = "HKLM:\SOFTWARE\Volition\Red Faction"; AccessKey = "InstallPath" }, # legacy 32-bit
    @{ Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Steam App 20530"; AccessKey = "InstallLocation" }, # steam
    @{ Path = "HKLM:\SOFTWARE\GOG.com\GOGREDFACTION"; AccessKey = "PATH" } # GOG
    @{ Path = "HKCU:\SOFTWARE\$($MyInvocation.MyCommand.Name)"; AccessKey = "RFPath" }
)

foreach ($r in $RegistryPaths) {
    if (Test-Path $r.Path) {
        $RF_path = (Get-ItemProperty -Path $r.Path)."$($r.AccessKey)"
        if ($RF_path.Length -gt 0) { break }
    }
}

if (($RF_path.Length -eq 0) -or (-not (Test-Path -LiteralPath $RF_path))) {
    [System.Windows.Forms.MessageBox]::Show("Unable to find RF folder.`nYou will be prompted to browse for the path yourself.", "Unable to find default RF folder", "OK", "Information")

    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = "Browse to RF root folder"
    }

    if ($fbd.ShowDialog() -eq "OK") {
        $RF_path = $fbd.SelectedPath

        if (Test-Path -Path "HKCU:\SOFTWARE\$($MyInvocation.MyCommand.Name)") {
            Set-ItemProperty -Path "HKCU:\SOFTWARE\$($MyInvocation.MyCommand.Name)" -Name "RFPath" -Value $RF_path | Out-Null
        } else {
            New-Item -Path "HKCU:\SOFTWARE" -Name $MyInvocation.MyCommand.Name | Out-Null
            New-ItemProperty -Path "HKCU:\SOFTWARE\$($MyInvocation.MyCommand.Name)" -Name "RFPath" -Value $RF_path | Out-Null
        }
    } else {
        Write-Host "Unable to find RF root folder" -ForegroundColor Red
        Exit
    }

    if ($RF_path.Length -eq 0) {
        Write-Host "Unable to run the program unless you select your RF root folder"
    }
}
#endregion

class VPPPosition {
    static [UInt32]Calculate([UInt32]$inpt) {
        [UInt32]$output = $inpt

        if ($inpt % 2048 -ne 0) {
            $output = $inpt - ($inpt % 2048)
            $output += 2048
        }

        return $output
    }
}

if (Test-Path -LiteralPath $RF_path) {
    [System.Collections.Generic.List[HashTable]]$match_table = New-Object System.Collections.Generic.List[HashTable]
    [System.IO.BinaryReader]$reader = $null

    try {
        # loop thru the list of VPP files in the RF folder
        Get-ChildItem *.vpp -LiteralPath $RF_path -File -Recurse | ForEach-Object {
            try {
                [string]$VPP_Filename = $_.FullName
                [byte[]]$buffer = [byte[]]::CreateInstance([byte], 60)
                $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($VPP_Filename, "Open")))

                # read VPP header information
                [UInt32]$signature = $reader.ReadUInt32()
                [UInt32]$version = $reader.ReadUInt32()

                # check that VPP header is well-formed
                if (($signature -eq 0x51890ace) -and ($version -eq 1)) {
                    # read file count in VPP (byte positions 8-11)
                    [UInt32]$num_files = $reader.ReadUInt32()

                    [UInt32]$prev_position = 2048
                    [UInt32]$prev_size = 64 * $num_files

                    for ([int]$n = 0; $n -lt $num_files; $n++) {
                        [void]$reader.BaseStream.Seek(2048 + $n*64, "Begin")

                        # filename in first 60 bytes of current position
                        [void]$reader.Read($buffer, 0, 60)
                        [string]$filename = [System.Text.Encoding]::ASCII.GetString($buffer).Split([byte]0)[0]

                        # file size in next 4 bytes
                        [UInt32]$filesize = $reader.ReadUInt32()

                        # calculate position to raw file data
                        [UInt32]$curr_position = [VPPPosition]::Calculate($prev_position + $prev_size)

                        # if filename contains search text
                        if ($filename.ToLower().Contains($SearchText.ToLower()) -or ($filename -like $SearchText)) {
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
            } catch {
                if ($_.Exception.Message -notlike "*Access to the path*is denied*") {
                    throw $_
                }
            }

            $reader.Close()
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
                        [Hashtable]$file_record = $match_table[$filenum - 1]

                        # read VPP into memory
                        $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($file_record.VPP_Filename, "Open")))
                        [void]$reader.BaseStream.Seek($file_record.Position, "Begin")

                        # retrieve file data from VPP
                        $buffer = [byte[]]::CreateInstance([byte], $file_record.Size)
                        [void]$reader.Read($buffer, 0, $file_record.Size)

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
                            [System.IO.File]::WriteAllBytes($FileDialog.FileName, $buffer)
                            Write-Host "Saved '$($FileDialog.FileName)'"
                            $extract_count++
                        }

                        $reader.Close()
                    } else {
                        Write-Host "There are only $($match_table.Count) records listed above. Try again." -ForegroundColor Red
                    }
                } else {
                    $prompt_extract = $false
                    Write-Host "Quitting"
                }
            }
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red

        if ($null -ne $reader) {
            $reader.Close()
        }
    }
}
