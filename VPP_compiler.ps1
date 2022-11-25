Add-Type -AssemblyName System.Windows.Forms

[string]$InstallLocation = [string]::Empty

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
        $InstallLocation = (Get-ItemProperty -Path $r.Path)."$($r.AccessKey)"
        if ($InstallLocation.Length -gt 0) { break }
    }
}

if (($InstallLocation.Length -eq 0) -or (-not (Test-Path -LiteralPath $InstallLocation))) {
    [System.Windows.Forms.MessageBox]::Show("Unable to find RF folder.`nYou will be prompted to browse for the path yourself.", "Unable to find default RF folder", "OK", "Information")

    $fbd = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
        Description = "Browse to RF root folder"
    }

    if ($fbd.ShowDialog() -eq "OK") {
        $InstallLocation = $fbd.SelectedPath

        if (Test-Path -Path "HKCU:\SOFTWARE\$($MyInvocation.MyCommand.Name)") {
            Set-ItemProperty -Path "HKCU:\Software\$($MyInvocation.MyCommand.Name)" -Name "RFPath" -Value $InstallLocation | Out-Null
        } else {
            New-Item -Path "HKCU:\Software" -Name $MyInvocation.MyCommand.Name | Out-Null
            New-ItemProperty -Path "HKCU:\Software\$($MyInvocation.MyCommand.Name)" -Name "RFPath" -Value $InstallLocation | Out-Null
        }
    } else {
        Write-Host "Unable to find RF root folder" -ForegroundColor Red
        Exit
    }

    if ($InstallLocation.Length -eq 0) {
        Write-Host "Unable to run the program unless you select your RF root folder"
    }
}
#endregion

#region functions
class VPPPosition {
    static [int]Calculate([UInt32]$inpt) {
        [UInt32]$output = $inpt

        if ($inpt % 2048 -ne 0) {
            $output = $inpt - ($inpt % 2048)
            $output += 2048
        }

        return $output
    }
}

[ScriptBlock]$LoadVPPFile = {
    param([string]$VPP_Filename, [Hashtable]$VPP_File)

    [System.IO.BinaryReader]$reader = $null
    $VPP_File.Files.Clear()
    $VPP_File.Filepath = $VPP_Filename

    try {
        $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($VPP_Filename, "Open")))
        [void]$reader.BaseStream.Seek(0, "Begin")

        [byte[]]$buffer = [byte[]]::CreateInstance([byte], 12)
        [void]$reader.Read($buffer, 0, 12)
        [UInt32]$signature = [System.BitConverter]::ToUInt32($buffer, 0)
        [UInt32]$version = [System.BitConverter]::ToUInt32($buffer, 4)

        if (($signature -eq 0x51890ace) -and ($version -eq 1)) {
            [UInt32]$num_files = [System.BitConverter]::ToUInt32($buffer, 8)
            [UInt32]$offset = 0x800

            [UInt32]$prev_position = 0x800
            [UInt32]$prev_size = 64 * $num_files

            for ([int]$n = 0; $n -lt $num_files; $n++) {
                [void]$reader.BaseStream.Seek(0x800 + $n*64, "Begin")

                # filename in first 60 bytes of current position
                [string]$filename = [System.Text.Encoding]::ASCII.GetString($reader.ReadBytes(60)).Split([byte]0)[0]
                [UInt32]$filesize = $reader.ReadUInt32()
                [UInt32]$curr_position = [VPPPosition]::Calculate($prev_position + $prev_size)

                [void]$VPP_File.Files.Add(@{
                    Filename = $filename
                    Size = $filesize
                    Position = $curr_position
                })

                $offset += 64
                $prev_position = $curr_position
                $prev_size = $filesize
            }
        } else {
            Write-Host "$($ofd.FileName) is not well-formed" -ForegroundColor Red
        }
    } catch {
        Write-Host $_.Exception.Message -ForegroundColor Red
    }

    if ($null -ne $reader) {
        $reader.Close()
    }
}

[ScriptBlock]$TriggerLoadVPPFile = {
    param([Hashtable]$VPP_File)

    $ofd = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Filter = "VPP files (*.vpp)|*.vpp"
        Title = "Open VPP file"
    }

    if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
        $ofd.InitialDirectory = $InstallLocation
    }

    if ($ofd.ShowDialog() -eq "OK") {
        Write-Host "Loading '$($ofd.FileName)'"
        $LoadVPPFile.Invoke($ofd.FileName, $VPP_File)
        Write-Host "$($VPP_File.Files.Count) files loaded from '$($ofd.FileName)'" -ForegroundColor Green
    } else {
        Write-Host "Cancelled"
    }
}

[ScriptBlock]$UnloadVPPFile = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Filepath -ne [string]::Empty) {
        $VPP_File.Files.Clear()
        Write-Host "Unloaded '$($VPP_File.Filepath)'.`nFile list is clear." -ForegroundColor Yellow
        $VPP_File.Filepath = [string]::Empty
    } else {
        Write-Host "Unloaded $($VPP_File.Files.Count) files" -ForegroundColor Yellow
        $VPP_File.Files.Clear()
    }
}

[ScriptBlock]$AddFiles = {
    param([Hashtable]$VPP_File)

    [Array]$extensions = @(
        "All files (*.*)|*.*",
        "Character files (*.v3c)|*.v3c",
        "Red Faction Levels (*.rfl)|*.rfl",
        "RFA files (*.rfa)|*.rfa",
        "Table files (*.tbl)|*.tbl",
        "Text files (*.txt)|*.txt",
        "TGA graphics (*.tga)|*.tga",
        "V3M files (*.v3m)|*.v3m",
        "VBM files (*.vbm)|*.vbm",
        ".VF fonts (*.vf)|*.vf",
        "VFX files (*.vfx)|*.vfx",
        "WAV files (*.wav)|*.wav"
    )

    $ofd = New-Object System.Windows.Forms.OpenFileDialog -Property @{
        Filter = $extensions -join '|'
        DefaultExt = 0
        Multiselect = $true
    }

    if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
        $ofd.InitialDirectory = $InstallLocation
    }

    if ($ofd.ShowDialog() -eq "OK") {
        foreach ($file in $ofd.FileNames) {
            $item = Get-Item -LiteralPath $file
            [Hashtable]$existing_file = $VPP_File.Files | Where-Object { $_.Filename -eq $item.Name } | Select-Object -First 1

            if ($existing_file -ne $null) {
                Write-Host "'$($item.Name)' already exists" -ForegroundColor Yellow
                $prompt_replace = Read-Host "Would you like to replace it? Y/n"

                if ($prompt_replace.ToLower() -eq "y") {
                    Write-Host "Replacing '$file'" -ForegroundColor Green
                    $existing_file.Size = $item.Length
                } else {
                    Write-Host "Skipping '$($item.FullName)'"
                }
            } else {
                Write-Host "Adding '$file'" -ForegroundColor Green
                $VPP_File.Files.Add(@{
                    Filename = $item.FullName
                    Size = $item.Length
                    Position = -1 # will be calculated upon VPP save
                })
            }
        }
    } else {
        Write-Host "Cancelled"
    }
}

[ScriptBlock]$RemoveFile = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Files.Count -gt 0) {
        [string]$file_to_remove = $file_to_remove = Read-Host "Enter the search terms for the file to remove (case insensitive, wildcard is *)"
        [Array]$existing_files = @($VPP_File.Files | Where-Object { $_.Filename -like $file_to_remove })

        if ($existing_files.Count -gt 0) {
            [int]$n = 1
            [bool]$removed_flag = $false

            foreach ($file in $existing_files) {
                Write-Host "$n. $($file.Filename), $([math]::Round($file.Size / 1KB, 2))KiB"
                $n++
            }

            [string]$file_num = [string]::Empty
            do {
                Write-Host ""

                if ($removed_flag) {
                    $file_num = Read-Host "Remove another? Type the file number listed above, or X to cancel"
                } else {
                    $file_num = Read-Host "Which file number listed above would you like to remove (X to cancel)"
                }

                if ($file_num.ToLower() -ne 'x') {
                    $file_num = $file_num.Trim(".")

                    if ((-not $file_num.StartsWith("0")) -and ($file_num -match '^\d+$')) {
                        [int]$file_idx = [convert]::ToInt32($file_num)

                        if ($file_idx -le $existing_files.Count) {
                            [string]$filename_to_remove = $existing_files[$file_idx - 1].Filename
                            [void]$VPP_File.Files.RemoveAll({ param($r) $r.Filename -eq $filename_to_remove })
                            Write-Host "Removed '$filename_to_remove' from this VPP." -ForegroundColor Green
                            $removed_flag = $true
                        } else {
                            Write-Host "There are only $($existing_files.Count) files listed above. Try again." -ForegroundColor Red
                        }
                    } else {
                        Write-Host "Not a valid option. Try again." -ForegroundColor Red
                    }
                }
            } until ($file_num.ToLower() -eq 'x')
        } else {
            Write-Host "No files match the search '$file_to_remove'" -ForegroundColor Red
        }
    } else {
        Write-Host "No files available to remove!" -ForegroundColor Red
    }
}

[ScriptBlock]$ListFiles15 = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Files.Count -gt 0) {
        [int]$pages = [Math]::Ceiling($VPP_File.Files.Count / 15.0)

        [int]$k = 1
        for ([int]$n = 0; $n -lt $pages; $n++) {
            $VPP_File.Files | Select-Object -Skip $($n * 15) -First 15 | ForEach-Object {
                Write-Host "$k. $($_.Filename), $([math]::Round($_.Size / 1KB, 2))KiB"
                $k++
            }
            Pause
        }

        Write-Host "End of file list"
    } else {
        Write-Host "No files available to list!" -ForegroundColor Red
    }
}

[ScriptBlock]$ListFiles = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Files.Count -gt 0) {
        [int]$n = 1
        foreach ($file in $VPP_File.Files) {
            Write-Host "$n. $($file.Filename), $([math]::Round($file.Size / 1KB, 2))KiB"
            $n++
        }

        Write-Host "End of file list"
    } else {
        Write-Host "No files available to list!" -ForegroundColor Red
    }
}

[ScriptBlock]$ExtractFile = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Files.Count -gt 0) {
        [string]$file_to_extract = Read-Host "Enter the search terms for the file to extract (case insensitive, wildcard is *)"

        [Array]$existing_files = @($VPP_File.Files | Where-Object { $_.Filename -like $file_to_extract })

        if ($existing_files.Count -gt 0) {
            [int]$n = 1
            [bool]$extracted_flag = $false
            foreach ($file in $existing_files) {
                Write-Host "$n. $($file.Filename), $([math]::Round($file.Size / 1KB, 2))KiB"
                $n++
            }

            do {
                Write-Host ""

                if ($extracted_flag) {
                    $file_to_extract = Read-Host "Extract another? Type the file number listed above, or X to cancel"
                } else {
                    $file_to_extract = Read-Host "Which file number listed above would you like to extract (X to cancel)?"
                }

                if ($file_to_extract.ToLower() -ne "x") {
                    $file_to_extract = $file_to_extract.Trim(".")

                    if ((-not $file_to_extract.StartsWith("0")) -and ($file_to_extract -match '^\d+$')) {
                        [int]$extract_idx = [Convert]::ToInt32($file_to_extract)

                        if ($extract_idx -gt $VPP_File.Files.Count) {
                            Write-Host "There are only $($existing_files.Count) files listed above. Try again."
                        } else {
                            [string]$extension = "All files (*.*)|*.*"
                            $existing_file = $existing_files[$extract_idx - 1]

                            [int]$ext_idx = $existing_file.Filename.IndexOf(".")
                            if ($ext_idx -gt 0) {
                                $extension = $existing_file.Filename.Substring($ext_idx)
                                $extension = "$($extension.Substring(1).ToUpper()) files (*$extension)|*$($extension)"
                            }

                            $sfd = New-Object System.Windows.Forms.SaveFileDialog -Property @{
                                Filter = $extension
                                Title = "Extract file"
                                Filename = $existing_file.Filename
                            }

                            if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
                                $sfd.InitialDirectory = $InstallLocation
                            }

                            [System.IO.BinaryReader]$reader = $null
                            if ($sfd.ShowDialog() -eq "OK") {
                                try {
                                    $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($VPP_File.Filepath, "Open")))
                                    [void]$reader.BaseStream.Seek($existing_file.Position, "Begin")
                                    [byte[]]$buffer = [byte[]]::CreateInstance([byte], $existing_file.Size)
                                    [void]$reader.Read($buffer, 0, $existing_file.Size)
                                    [System.IO.File]::WriteAllBytes($sfd.FileName, $buffer)
                                } catch {
                                    Write-Host $_.Exception.Message -ForegroundColor Red
                                }

                                if ($null -ne $reader) {
                                    $reader.Close()
                                }
                                Write-Host "Extracted '$($sfd.FileName)'" -ForegroundColor Green
                            } else {
                                Write-Host "Cancelled"
                            }

                            $extracted_flag = $true
                        }
                    } else {
                        Write-Host "Not a valid option. Try again." -ForegroundColor Red
                    }
                }
            } until ($file_to_extract.ToLower() -eq 'x')
        } else {
            Write-Host "No files match the search '$file_to_extract'" -ForegroundColor Red
        }
    } else {
        Write-Host "No files available to extract!" -ForegroundColor Red
    }
}

[ScriptBlock]$ExtractAllFiles = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Files.Count -gt 0) {
        # Extract all VPP files to a folder
        $fbd = New-Object System.Windows.Forms.FolderBrowserDialog -Property @{
            Description = "Select a destination folder"
        }

        if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
            $fbd.SelectedPath = $InstallLocation
        }

        if ($fbd.ShowDialog() -eq "OK") {
            [System.IO.BinaryReader]$reader = $null

            try {
                $reader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($VPP_File.Filepath, "Open")))

                # write each file to selected path
                foreach ($file in ($VPP_File.Files | Where-Object { $_.Position -gt -1 })) {
                    $savepath = Join-Path $fbd.SelectedPath $file.Filename

                    $buffer = [byte[]]::CreateInstance([byte], $existing_file.Size)
                    [void]$reader.BaseStream.Seek($existing_file.Position, "Begin")
                    [byte[]]$buffer = [byte[]]::CreateInstance([byte], $existing_file.Size)
                    [void]$reader.Read($buffer, 0, $existing_file.Size)
                    [System.IO.File]::WriteAllBytes($savepath, $buffer)
                }
            } catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            if ($null -ne $reader) {
                $reader.Close()
            }

            Write-Host "$($VPP_File.Files.Count) files have been extracted to '$($fbd.SelectedPath)'" -ForegroundColor Green
        } else {
            Write-Host "Cancelled"
        }
    } else {
        Write-Host "No files available to extract!" -ForegroundColor Red
    }
}

[ScriptBlock]$SaveVPPFile = {
    param([Hashtable]$VPP_File)

    if ($VPP_File.Files.Count -gt 0) {
        $sfd = New-Object System.Windows.Forms.SaveFileDialog -Property @{
            Filter = "VPP files (*.vpp)|*.vpp"
            Title = "Save VPP file"
            FileName = $VPP_File.Filepath
        }

        if (($InstallLocation.Length -gt 0) -and (Test-Path -LiteralPath $InstallLocation)) {
            $sfd.InitialDirectory = $InstallLocation
        }

        if ($sfd.ShowDialog() -eq "OK") {
            [System.IO.FileStream]$stream = $null
            [System.IO.StreamWriter]$writer = $null
            [System.IO.BinaryReader]$CurrentReader = $null
            [string]$tmp_filename = Join-Path ([IO.Path]::GetDirectoryName($sfd.FileName)) "$((New-Guid).Guid).tmp"
            [byte[]]$buffer = $null

            # calculate VPP save positions
            [UInt32]$curr_position = [VPPPosition]::Calculate(0x800 + 64 * $VPP_File.Files.Count)
            [UInt32]$next_position = 0
            foreach ($file in ($VPP_File.Files | Sort-Object -Property Filename)) {
                $next_position = [VPPPosition]::Calculate($curr_position + $file.Size)
                $file.Position = $curr_position
                $curr_position = $next_position
            }

            [UInt32]$vpp_total_size = $next_position
            [byte[]]@([byte]0) | Set-Content -LiteralPath $tmp_filename -AsByteStream

            try {
                $stream = New-Object IO.FileStream($tmp_filename, "Open", "ReadWrite")
                $writer = New-Object IO.StreamWriter($stream)
                $CurrentReader = New-Object System.IO.BinaryReader((New-Object System.IO.FileStream($VPP_File.Filepath, "Open")))

                [void]$writer.BaseStream.SetLength($vpp_total_size)
                [void]$writer.BaseStream.Seek(0, "Begin")

                # write VPP header
                [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes(0x51890ace), 0, 4)
                [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes(1), 0, 4)
                [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($VPP_File.Files.Count), 0, 4)
                [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($vpp_total_size), 0, 4)

                [UInt32]$offset = 0x800

                foreach ($file in ($VPP_File.Files | Sort-Object -Property Filename)) {
                    [string]$filename = $file.Filename

                    if ($file.Position -eq -1) {
                        # retrieve file data from filesystem
                        $buffer = [IO.File]::ReadAllBytes($filename)
                        $filename = [IO.Path]::GetFileName($filename)
                    } else {
                        # seek file data within vpp
                        $buffer = [byte[]]::CreateInstance([byte], $file.Size)
                        [void]$CurrentReader.BaseStream.Seek($file.Position, "Begin")
                        [void]$CurrentReader.Read($buffer, 0, $file.Size)
                    }

                    if ($filename.Length -gt 60) {
                        $filename = $filename.Substring(0, 60)
                    }

                    # write file index info
                    [void]$writer.BaseStream.Seek($offset, "Begin")
                    [void]$writer.BaseStream.Write([System.Text.Encoding]::ASCII.GetBytes($filename), 0, $filename.Length)
                    [void]$writer.BaseStream.Seek($offset+60, "Begin")
                    [void]$writer.BaseStream.Write([System.BitConverter]::GetBytes($file.Size), 0, 4)

                    # write file data
                    [void]$writer.BaseStream.Seek($file.Position, "Begin")
                    [void]$writer.BaseStream.Write($buffer, 0, $buffer.Count)

                    $offset += 64
                }
            } catch {
                Write-Host $_.Exception.Message -ForegroundColor Red
            }

            if ($null -ne $writer) {
                $writer.Close()
            }
            if ($null -ne $stream) {
                $stream.Close()
            }
            if ($null -ne $CurrentReader) {
                $CurrentReader.Close()
            }

            try {
                if (Test-Path -LiteralPath $sfd.FileName) {
                    Remove-Item -LiteralPath $sfd.FileName -Force
                }

                Rename-Item -LiteralPath $tmp_filename -NewName $sfd.FileName

                $LoadVPPFile.Invoke($sfd.FileName, $VPP_File)
            } catch {
                Write-Host $_ -ForegroundColor Red
            }

            Write-Host "Saved '$($sfd.FileName)'" -ForegroundColor Green
        } else {
            Write-Host "Cancelled"
        }
    } else {
        Write-Host "No files available to write!" -ForegroundColor Red
    }
}
#endregion

#region global variables
[string]$inpt = [string]::Empty

[Hashtable]$VPP_File = @{
    Filepath = [string]::Empty
    Files = New-Object System.Collections.ArrayList
}

[Array]$MenuOptions = @(
    @{ Key = "1"; Text = "Load a VPP file"; Action = $TriggerLoadVPPFile },
    @{ Key = "2"; Text = "Unload current VPP file"; Action = $UnloadVPPFile },
    @{ Key = "3"; Text = "Add files"; Action = $AddFiles },
    @{ Key = "4"; Text = "Remove a file"; Action = $RemoveFile },
    @{ Key = "5"; Text = "List files (15 per page)"; Action = $ListFiles15 },
    @{ Key = "6"; Text = "List all files"; Action = $ListFiles },
    @{ Key = "7"; Text = "Extract a file"; Action = $ExtractFile },
    @{ Key = "8"; Text = "Extract all files"; Action = $ExtractAllFiles },
    @{ Key = "9"; Text = "Save VPP file"; Action = $SaveVPPFile },
    @{ Key = "X"; Text = "Quit" }
)
#endregion

#region main loop
do {
    # list menu options
    Write-Host ""
    foreach ($item in $MenuOptions) {
        Write-Host "$($item.Key). $($item.Text)"
    }

    $inpt = Read-Host "Menu option"

    Write-Host ""

    if ($inpt.ToLower() -ne "x") {
        [Hashtable]$SelectedMenuOption = $MenuOptions | Where-Object { $_.Key -eq $inpt.ToLower() }

        if ($null -ne $SelectedMenuOption) {
            $SelectedMenuOption.Action.Invoke($VPP_File)
        } else {
            Write-Host "Menu item '$inpt' is not valid" -ForegroundColor Red
        }
    }
} until ($inpt.ToLower() -eq 'x')

Write-Host "Quitting"
#endregion
