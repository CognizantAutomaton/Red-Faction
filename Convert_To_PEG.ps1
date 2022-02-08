Add-Type -AssemblyName System.Windows.Forms

function Encode-Indexed {
    param(
        $filename,
        $Bitmap
    )
    [byte[]]$texture_bytes = [byte[]]::new(1024 + $Bitmap.Width * $Bitmap.Height)

    # create palette i.e. find all unique colors
    $palette_table = @{}
    foreach ($y in 1..$Bitmap.Height) {
        foreach ($x in 1..$Bitmap.Width) {
            $palette_table[$Bitmap.GetPixel($x - 1,$y - 1)] = $true
            if ($palette_table.Count -gt 256) {
                [System.Windows.Forms.MessageBox]::Show("File '$filename' has more than 256 colors! Quitting")
                Exit
            }
        }
    }

    [System.Drawing.Color[]]$palette = [System.Drawing.Color[]]::new(256)
    for ($k = 0; $k -lt 256; $k++) {
        $palette[$k] = [System.Drawing.Color]::FromArgb(0, 0, 0, 0)
    }

    [System.Drawing.Color[]]$unique_colors = $palette_table.Keys.GetEnumerator() | Select-Object
    [Array]::Copy($unique_colors, 0, $palette, 0, $unique_colors.Length)

    [int]$n = 0

    # assign palette values to first 1024 bytes
    foreach ($color in $palette) {
        $texture_bytes[$n] = $color.R
        $texture_bytes[$n+1] = $color.G
        $texture_bytes[$n+2] = $color.B
        $texture_bytes[$n+3] = $color.A

        $n += 4
    }

    # build index
    foreach ($y in 1..$Bitmap.Height){
        foreach ($x in 1..$Bitmap.Width) {
            $idx = $palette.IndexOf($Bitmap.GetPixel($x-1, $y-1))
            $idx = ($idx -band 0xE7) -bor (($idx -shr 1) -band 0x8) -bor (($idx -shl 1) -band 0x10)
            $texture_bytes[$n++] = $idx
        }
    }

    return $texture_bytes
}

function Encode-RGBA8888 {
    param(
        $Bitmap
    )

    [byte[]]$texture_bytes = [byte[]]::new(4 * $Bitmap.Width * $Bitmap.Height)
    [Array]::Clear($texture_bytes, 0, 4 * $Bitmap.Width * $Bitmap.Height)

    [int]$n = 0
    foreach ($y in 1..$Bitmap.Height) {
        foreach ($x in 1..$Bitmap.Width) {
            [System.Drawing.Color]$color = $Bitmap.GetPixel($x-1, $y-1)

            $texture_bytes[$n] = $color.R
            $texture_bytes[$n+1] = $color.G
            $texture_bytes[$n+2] = $color.B
            $texture_bytes[$n+3] = $color.A

            $n += 4
        }
    }

    return $texture_bytes
}

function Encode-RGBA5551 {
    param(
        $Bitmap
    )

    [uint16]$texture_bytes = [uint16]::new(2 * $Bitmap.Width * $Bitmap.Height)
    [Array]::Clear($texture_bytes, 0, 2 * $Bitmap.Width * $Bitmap.Height)

    [int]$n = 0
    foreach ($y in 1..$Bitmap.Height) {
        foreach ($x in 1..$Bitmap.Width) {
            [System.Drawing.Color]$color = $Bitmap.GetPixel($x-1, $y-1)

            [int]$a = $(if ($color.A -gt 0) {1} else {0})
            [int]$r = $color.R * 31 / 255
            [int]$g = $color.G * 31 / 255
            [int]$b = $color.B * 31 / 255

            $r = $r -shl 11
            $g = $g -shl 6
            $b = $b -shl 1

            [uint16]$bgra = $b -bor $g -bor $r -bor $a
            [Array]::Copy([System.BitConverter]::GetBytes($bgra), 0, $texture_bytes, $n, 2)

            $n += 2
        }
    }

    return $texture_bytes
}

# texture and frame count to be used in PEG header
[int]$texture_count = 0
[int]$frame_count = 0
[int]$total_offset = 0
[string]$presets_path = Join-Path $PSScriptRoot "texture_presets.json"
[PSCustomObject]$presets = $null

if (Test-Path -LiteralPath $presets_path) {
    # load presets
    $presets = Get-Content -LiteralPath $presets_path | ConvertFrom-Json
} else {
    # handle preset JSON file not found error
    [void][System.Windows.Forms.MessageBox]::Show("Unable to find presets at '$presets_path'", "Missing file error", "OK", "Error")
    Exit
}

# instantiate open file dialog
$ofd = New-Object System.Windows.Forms.OpenFileDialog -Property @{
    DefaultExt = ".png"
    Filter = "PNG files|*.png"
    FilterIndex = 0
    Multiselect = $true
    Title = "Select images to encode"
}

# show open file dialog
if ($ofd.ShowDialog() -eq 'OK') {
    # store texture count
    $texture_count = $ofd.Filenames.Length

    # compute offset to first texture data (32 bytes PEG header + 64 bytes per texture header)
    $total_offset = 32 + 64 * $ofd.Filenames.Length

    # collect texture info
    [Array]$texture_map_data = @()
    foreach ($filename in ($ofd.Filenames | Sort-Object)) {
        # read image file
        $Bitmap = [System.Drawing.Bitmap]::FromFile((Resolve-Path $filename).ProviderPath)
        Write-Host "Encoding '$filename'"

        # retrieve filename
        [string]$filename_to_write = [System.IO.Path]::GetFileName($filename)

        [int]$idx = $filename_to_write.IndexOf(".")
        if ($idx -gt 0) {
            # determine filename without extension
            $filename_to_write = $filename_to_write.Substring(0, $idx)
        }

        # lookup a preset to use for this filename
        [PSCustomObject]$preset_to_use = $presets | Where-Object {
            $_.Name.StartsWith($filename_to_write) -and ($Bitmap.Width -eq $_.Width) -and ($Bitmap.Height -eq $_.Height)
        } | Select-Object -First 1

        if ($preset_to_use -eq $null) {
            # handle preset not found error
            [void][System.Windows.Forms.MessageBox]::Show("Unable to find preset in texture_presets.json for texture file '$filename'`nPlease create an entry", "Preset error", "OK", "Error")
            Exit
        }

        $filename_to_write = $preset_to_use.Name

        # store frame count
        $frame_count += $preset_to_use.Frame_Count

        # store texture
        $texture_map_data += @{
            Width = $Bitmap.Width
            Height = $Bitmap.Height
            Format_1 = $preset_to_use.Format
            Format_2 = $preset_to_use.Format2
            Flags = $preset_to_use.Flags
            Frame_Count = $preset_to_use.Frame_Count
            Animation_Delay = $preset_to_use.Animation_Delay
            Mip_Count = $preset_to_use.Mip_Count
            Unknown_1 = $preset_to_use.Unk1
            Unknown_2 = $preset_to_use.Unk2
            Filename = $filename_to_write
            Offset = $total_offset
            Data = $(switch -exact ($preset_to_use.Format) {
                3 { Encode-RGBA5551 $Bitmap }
                4 { Encode-Indexed $filename $Bitmap }
                7 { Encode-RGBA8888 $Bitmap }
                default {
                    # handle incorrect format error
                    [void][System.Windows.Forms.MessageBox]::Show("Format preset for '$filename_to_write' not recognized.`nCurrent value: $($preset_to_use.Format)`nExpected possible values:`n`t3 (RGBA5551)`n`t4 (Indexed)`n`t7 (RGBA8888)", "Format error", "OK", "Error")
                    Exit
                }
            })
        }

        # compute offset to next texture data (by adding the size of current texture)
        $total_offset += $texture_map_data[-1].Data.Length
    }

    Write-Host "Ready to write PEG file"

    # instantiate save file dialog
    $sfd = New-Object System.Windows.Forms.SaveFileDialog -Property @{
        Filter = "PEG files|*.peg"
        Title = "Save PEG file"
    }

    # show save file dialog
    if ($sfd.ShowDialog() -eq 'OK') {
        # initialize PEG bytes
        [byte[]]$peg_data = [byte[]]::new($total_offset)
        [Array]::Clear($peg_data, 0, $total_offset)

        # write PEG header
        [Array]::Copy([System.BitConverter]::GetBytes(0x564B4547), 0, $peg_data, 0, 4) # signature 0-3
        [Array]::Copy([System.BitConverter]::GetBytes(6), 0, $peg_data, 4, 4) # version 4-7
        [Array]::Copy([System.BitConverter]::GetBytes(32), 0, $peg_data, 8, 4) # header size 8-11
        [Array]::Copy([System.BitConverter]::GetBytes($total_offset), 0, $peg_data, 12, 4) # PEG file size 12-15
        [Array]::Copy([System.BitConverter]::GetBytes($texture_count), 0, $peg_data, 16, 4) # texture count 16-19
        [Array]::Copy([System.BitConverter]::GetBytes(0), 0, $peg_data, 20, 4) # unknown14 20-23
        [Array]::Copy([System.BitConverter]::GetBytes($frame_count), 0, $peg_data, 24, 4) # frame count 24-27
        [Array]::Copy([System.BitConverter]::GetBytes(16), 0, $peg_data, 28, 4) # unknown1c 28-31

        # textures header starts at position 32
        [int]$position = 32
        foreach ($t in $texture_map_data.GetEnumerator()) {
            # write texture header
            [Array]::Copy([System.BitConverter]::GetBytes($t.Width), 0, $peg_data, $position, 2)
            $position += 2

            [Array]::Copy([System.BitConverter]::GetBytes($t.Height), 0, $peg_data, $position, 2)
            $position += 2

            [Array]::Copy([System.BitConverter]::GetBytes($t.Format_1), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Format_2), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Flags), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Frame_Count), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Animation_Delay), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Mip_Count), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Unknown_1), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.BitConverter]::GetBytes($t.Unknown_2), 0, $peg_data, $position, 1)
            $position++

            [Array]::Copy([System.Text.Encoding]::ASCII.GetBytes($t.Filename), 0, $peg_data, $position, $t.Filename.Length)
            $position += 48

            [Array]::Copy([System.BitConverter]::GetBytes($t.Offset), 0, $peg_data, $position, 4)
            $position += 4
            # end of texture header

            # write texture data
            [Array]::Copy($t.Data, 0, $peg_data, $t.Offset, $t.Data.Length)
        }

        # save to disk
        [System.IO.File]::WriteAllBytes($sfd.FileName, $peg_data)
        Write-Host "Saved '$($sfd.FileName)'"
    }
}
