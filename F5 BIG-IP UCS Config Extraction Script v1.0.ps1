# This script performs the following tasks

# Part 1: Extract specific logs from raw data
# 1st level decompression of ucs files using DeGZip-File function
# 2nd level decompression of ucs files using 7-zip
# Extracts bigip.conf and bigip_base.conf configuration from /config directory
# Extracts bigip.conf from /config/partitions directory
# Extracts bigip_gtm.conf from /config directory
# Extracts certificate information from \var\tmp\filestore_temp\files_d\Common_d\certificate_d directory
# Extracts ASM Policy Name from var\tmp\ts_db.save_dir_*.cstmp\ts_db.data.PLC.PL_POLICIES.cstmp file

# Part 2: Populate extracted config into excel
# Collapse multiline to single line config for bigip.conf
# Collapse multiline to single line config for bigip_base.conf
# Collapse multiline to single line config for bigip_user.conf
# Collapse multiline to single line config for BigDB.dat
# Collapse multiline to single line config for bigip_gtm.conf
# Extracts config from bipip.conf
# Extracts config from bigip_base.conf
# Extracts config from bigip_user.conf
# Extracts config from BigDB.dat
# Extracts ASM Policy Name from var\tmp\ts_db.save_dir_*.cstmp\ts_db.data.PLC.PL_POLICIES.cstmp
# Extracts config from bigip_gtm.conf
# Extracts config from /config/partitions/<partitionname>/bigip.conf
# Extracts config from /config/partitions/<partitionname>/bigip_gtm.conf

# Part 3: Extract certificate information
# Extract certificate information from \var\tmp\filestore_temp\files_d\Common_d\certificate_d directory
# Extract device certificate information from \var\tmp\cert_temp\conf\ssl.crt file
# Extract certificate information from \var\tmp\gtm_tmp\gtm directory
# Export the certificate details into Excel
# Log excel paste error to $DstDir\$path-paste_errors.log

# Prerequisites:

# Run Windows Powershell ISE as Administrator
# Download and install 7-zip from https://www.7-zip.org/
# Create SrcDir and DstDir folder
# Place ReportTemplate in SrcDir
# Place UCS files in SrcDir
# Place Extraction Template in SrcDir

Set-ExecutionPolicy -ExecutionPolicy Bypass

$7zDir = $null
$regPath = "HKLM:\Software\7-Zip"


if (Test-Path $regPath) {
    $7zDir = (Get-ItemProperty -Path $regPath).Path
}

# Check if $zipPath is still null (meaning no result was found)
if (-not $7zDir) {
    Write-Host -ForegroundColor black -BackgroundColor Yellow "7-Zip installation not found. Exiting script."
    exit
}

# Output the result if found
Write-Host -ForegroundColor black -BackgroundColor Yellow "7-Zip installation path: $7zDir"

# File Variable Control

$7zExe = "$7zDir\7z.exe"
$SrcDir = "c:\Temp\ExtractSrc"
$DstDir = "c:\Temp\ExtractDst"
$Ext = "ucs"
$ReportTemplate1 = "F5 BIG-IP UCS Config Template v1.0.xlsx"

$Date = (Get-Date -Format yyyyMMdd) + "-" + (Get-Date -Format HHmmss) + ".xlsx"

# Part 1: Extract specific logs from raw data

$StartTime = Get-Date
cd $SrcDir

# Add DeGZip-File function to do 1st level decompression

Function DeGZip-File{
    Param(
        $infile,
        $outfile = ($infile -replace '\.gz$','')
        )
    $input = New-Object System.IO.FileStream $inFile, ([IO.FileMode]::Open), ([IO.FileAccess]::Read), ([IO.FileShare]::Read)
    $output = New-Object System.IO.FileStream $outFile, ([IO.FileMode]::Create), ([IO.FileAccess]::Write), ([IO.FileShare]::None)
    $gzipStream = New-Object System.IO.Compression.GzipStream $input, ([IO.Compression.CompressionMode]::Decompress)
    $buffer = New-Object byte[](1024)
    while($true){
        $read = $gzipstream.Read($buffer, 0, 1024)
        if ($read -le 0){break}
        $output.Write($buffer, 0, $read)
        }
    $gzipStream.Close()
    $output.Close()
    $input.Close()
}

$filenames = $null
$filenames = gci -name -File *.$Ext

if (-not $filenames) {
    Write-Host -ForegroundColor black -BackgroundColor Yellow "$Ext not found. Exiting script."
    exit
}

$filenames

# Use 7-zip to do 2nd level decompression

Foreach ($path in $filenames) {
    $infile = "$SrcDir\$path"
    $outfile = "$SrcDir\$path.2"
    $outfile2 = "$SrcDir\$path.3"
    Write-Host -ForegroundColor black -BackgroundColor Yellow "Decompressing $infile"
    DeGZip-File $infile $outfile
    Write-Host -ForegroundColor black -BackgroundColor Yellow "Decompressed to $outfile"
    Write-Host -ForegroundColor black -BackgroundColor Yellow "Decompressing $SrcDir\$path.2"
    #cmd /c $7zExe x -r -y "$SrcDir\$path.2" -o"$SrcDir\$path.3"
    $Arg = "x -r -y "+"$SrcDir\$path.2"+" -o"+"$SrcDir\$path.3"
    Start-Process -FilePath $7zExe -Args $Arg -Verb RunAs
    $nid = (Get-Process 7z).id
    Wait-Process -Id $nid
    Write-Host -ForegroundColor black -BackgroundColor Yellow "Decompressed to $SrcDir\$path.3"

# Clear-Variable line2

    if (Get-Variable -Name line2 -Scope Global -ErrorAction SilentlyContinue) {
    Clear-Variable -Name line2 -Scope Global
    }

# Collapse multiline to single line config for bigip.conf

    $bigipPath = "$SrcDir\$path.3\config\bigip.conf"
    $bigipoutputPath = "$SrcDir\$path.3\bigip_output.txt"
    $lines = $null
    $line2 = $null

    $lines = gc $bigipPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing $bigipPath multiline to single line config"

    Foreach ($line in $lines) {
        if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("}")) {
            if ( $line  -match ("^[a-z].*") -and $line.EndsWith("}")) {
                $line2 += $line+"`r`n"
                } elseif ( $line  -match ("^[a-z].*")) {
                $line2 += $line -replace '\s+',' '
                } elseif ( $line  -match ("^[A-Z]")) {
                $line2 += $line -replace '\s+',' '
                } elseif ( $line.StartsWith(" ")) {
                $line2 += $line -replace '\s+',' '
                } elseif ( $line.StartsWith("}")) {
                $line2 += $line+"`r`n"
            }
        }
    }

    $line2 | Out-File -width 10000 $bigipoutputPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $bigipoutputPath" 

# Collapse multiline to single line config for bigip_base.conf

    $bigipbasePath = "$SrcDir\$path.3\config\bigip_base.conf"
    $bigipbaseoutputPath = "$SrcDir\$path.3\bigip_base_output.txt"
    $lines = $null
    $line2 = $null

    $lines = gc $bigipbasePath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing $bigipbasePath multiline to single line config"

    Foreach ($line in $lines) {
        #$line
        if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("}")) {
            if ( $line  -match ("^[a-z].*") -and $line.EndsWith("}")) {
                $line2 += $line+"`r`n"
            } elseif ( $line  -match ("^[a-z].*")) {
            $line2 += $line -replace '\s+',' '
            } elseif ( $line  -match ("^[A-Z]")) {
            $line2 += $line -replace '\s+',' '
            } elseif ( $line.StartsWith(" ")) {
            $line2 += $line -replace '\s+',' '
            } elseif ( $line.StartsWith("}")) {
            $line2 += $line+"`r`n"
            }
        }
    }

    $line2 | Out-File -width 10000 $bigipbaseoutputPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $bigipbaseoutputPath" 

# Collapse multiline to single line config for bigip_user.conf

    $bigipuserPath = "$SrcDir\$path.3\config\bigip_user.conf"
    $bigipuseroutputPath = "$SrcDir\$path.3\bigip_user_output.txt"
    $lines = $null
    $line2 = $null

    $lines = gc $bigipuserPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing $bigipuserPath multiline to single line config"

    Foreach ($line in $lines) {
        #$line
        if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("}")) {
            if ( $line  -match ("^[a-z].*") -and $line.EndsWith("}")) {
                $line2 += $line+"`r`n"
            } elseif ( $line  -match ("^[a-z].*")) {
            $line2 += $line -replace '\s+',' '
            } elseif ( $line  -match ("^[A-Z]")) {
            $line2 += $line -replace '\s+',' '
            } elseif ( $line.StartsWith(" ")) {
            $line2 += $line -replace '\s+',' '
            } elseif ( $line.StartsWith("}")) {
            $line2 += $line+"`r`n"
            }
        }
    }

    $line2 | Out-File -width 10000 $bigipuseroutputPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $bigipuseroutputPath" 

# Collapse multiline to single line config for BigDB.dat

    $bigDBPath = "$SrcDir\$path.3\config\BigDB.dat"
    $bigDBoutputPath = "$SrcDir\$path.3\BigDB_output.txt"
    $lines = $null
    $line2 = $null

    $lines = gc $bigDBPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing $bigDBPath multiline to single line config"

    Foreach ($line in $lines) {
        #$line
        if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("[")) {
            if ( $line  -match ("^[a-z].*") -and $line.EndsWith("[")) {
                $line2 += $line+"`r`n"
            } elseif ( $line  -match ("^[a-z].*")) {
            $line2 += $line+":" -replace '\s+',' '
            } elseif ( $line  -match ("^[A-Z]")) {
            $line2 += $line+":" -replace '\s+',' '
            } elseif ( $line.StartsWith(" ")) {
            $line2 += $line+":" -replace '\s+',' '
            } elseif ( $line.StartsWith("[")) {
            $line2 += "`r`n"+$line+":"
            }
        }
    }

    $line2 | Out-File -width 10000 $bigDBoutputPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $bigDBoutputPath" 

# Collapse multiline to single line config for bigip_gtm.conf

    $bigipgtmPath = "$SrcDir\$path.3\config\bigip_gtm.conf"
    $bigipgtmoutputPath = "$SrcDir\$path.3\bigip_gtm_output.txt"

if (Test-Path -Path $bigipgtmPath) {
    $lines = $null
    $line2 = $null
    $lines = Get-Content $bigipgtmPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing $bigipgtmPath multiline to single line config"

    Foreach ($line in $lines) {
        if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("}")) {
            if ( $line  -match ("^[a-z].*") -and $line.EndsWith("}")) {
                $line2 += $line+"`r`n"
                } elseif ( $line  -match ("^[a-z].*")) {
                $line2 += $line -replace '\s+',' '
                } elseif ( $line  -match ("^[A-Z]")) {
                $line2 += $line -replace '\s+',' '
                } elseif ( $line.StartsWith(" ")) {
                $line2 += $line -replace '\s+',' '
                } elseif ( $line.StartsWith("}")) {
                $line2 += $line+"`r`n"
            }
        }
    } 

    $line2 | Out-File -width 10000 $bigipgtmoutputPath

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $SrcDir\$path.3\bigip_gtm_output.txt"


    } else {
    Write-Host -ForegroundColor Yellow "Config file not found at $bigipgtmPath - skipping this section"
    }


# Extracts config from bipip.conf

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving profile data"
    gc $SrcDir\$path.3\config\bigip.conf -ReadCount 10000 | % { $_ } | sls "ltm profile" | ? {$_ -match '\S'} | Out-File -width 10000 $SrcDir\$path.3\ltmprofile1.txt
    gc $SrcDir\$path.3\ltmprofile1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 ( New-Item -Path "$DstDir\$path\ltmprofile.txt" -Force )

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving route data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "net route" | Out-File -width 10000 $SrcDir\$path.3\route1.txt
    gc $SrcDir\$path.3\route1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\route.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving node data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm node" | Out-File -width 10000 $SrcDir\$path.3\ltmnode1.txt
    gc $SrcDir\$path.3\ltmnode1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmnode.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving pool data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm pool" | Out-File -width 10000 $SrcDir\$path.3\ltmpool1.txt
    gc $SrcDir\$path.3\ltmpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmpool.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving SNAT pool data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm snatpool " | Out-File -width 10000 $SrcDir\$path.3\ltmsnatpool1.txt
    gc $SrcDir\$path.3\ltmsnatpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmsnatpool.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving virtual server data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm virtual " | Out-File -width 10000 $SrcDir\$path.3\ltmvirtualserver1.txt
    gc $SrcDir\$path.3\ltmvirtualserver1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmvirtualserver.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving virtual address data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm virtual-address " | Out-File -width 10000 $SrcDir\$path.3\ltmvirtualaddress1.txt
    gc $SrcDir\$path.3\ltmvirtualaddress1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmvirtualaddress.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving client SSL data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm profile client-ssl" | Out-File -width 10000 $SrcDir\$path.3\ltmclientssl1.txt
    gc $SrcDir\$path.3\ltmclientssl1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmclientssl.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving server SSL data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm profile server-ssl" | Out-File -width 10000 $SrcDir\$path.3\ltmserverssl1.txt
    gc $SrcDir\$path.3\ltmserverssl1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmserverssl.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving certificate data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "certificate_d" | Out-File -width 10000 $SrcDir\$path.3\cert1.txt
    gc $SrcDir\$path.3\cert1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\cert.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving key data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "certificate_key_d" | Out-File -width 10000 $SrcDir\$path.3\key1.txt
    gc $SrcDir\$path.3\key1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\key.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving auth data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "^auth " | Out-File -width 10000 $SrcDir\$path.3\auth1.txt
    gc $SrcDir\$path.3\auth1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\auth.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cipher group"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm cipher group " | Out-File -width 10000 $SrcDir\$path.3\ltmciphergroup1.txt
    gc $SrcDir\$path.3\ltmciphergroup1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmciphergroup.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cipher rule"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm cipher rule " | Out-File -width 10000 $SrcDir\$path.3\ltmcipherrule1.txt
    gc $SrcDir\$path.3\ltmcipherrule1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmcipherrule.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving monitor data"
    gc $bigipoutputPath -ReadCount 10000 | % { $_ } | sls "ltm monitor " | Out-File -width 10000 $SrcDir\$path.3\ltmmonitor1.txt
    gc $SrcDir\$path.3\ltmmonitor1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ltmmonitor.txt

# Extracts config from bigip_base.conf

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving hostname data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys global-settings " | Out-File -width 10000 $SrcDir\$path.3\hostname1.txt
    gc $SrcDir\$path.3\hostname1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000  ( New-Item -Path "$DstDir\$path\Hostname.txt" -Force )

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving mgmt ip data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys management-ip " | Out-File -width 10000 $SrcDir\$path.3\mgmtip1.txt
    gc $SrcDir\$path.3\mgmtip1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\mgmtip.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving mgmt route data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys management-route " | Out-File -width 10000 $SrcDir\$path.3\mgmtroute1.txt
    gc $SrcDir\$path.3\mgmtroute1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\mgmtroute.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving self ip data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "net self " | Out-File -width 10000 $SrcDir\$path.3\selfip1.txt
    gc $SrcDir\$path.3\selfip1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\selfip.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving trunk data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "net trunk " | Out-File -width 10000 $SrcDir\$path.3\trunk1.txt
    gc $SrcDir\$path.3\trunk1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\trunk.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving VLAN data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "net vlan " | Out-File -width 10000 $SrcDir\$path.3\vlan1.txt
    gc $SrcDir\$path.3\vlan1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\vlan.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving DNS data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys dns " | Out-File -width 10000 $SrcDir\$path.3\dns1.txt
    gc $SrcDir\$path.3\dns1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\dns.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving NTP data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys ntp " | Out-File -width 10000 $SrcDir\$path.3\ntp1.txt
    gc $SrcDir\$path.3\ntp1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\ntp.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving provisioning data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys provision " | Out-File -width 10000 $SrcDir\$path.3\provision1.txt
    gc $SrcDir\$path.3\provision1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\provision.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving SNMP data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys snmp " | Out-File -width 10000 $SrcDir\$path.3\snmp1.txt
    gc $SrcDir\$path.3\snmp1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\snmp.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving SSHD data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys sshd " | Out-File -width 10000 $SrcDir\$path.3\sshd1.txt
    gc $SrcDir\$path.3\sshd1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\sshd.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving HTTPD data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys httpd " | Out-File -width 10000 $SrcDir\$path.3\httpd1.txt
    gc $SrcDir\$path.3\httpd1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\httpd.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving Password Policy data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "auth password-policy " | Out-File -width 10000 $SrcDir\$path.3\passwordpolicy1.txt
    gc $SrcDir\$path.3\passwordpolicy1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\passwordpolicy.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cli global-settings data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "cli global-settings " | Out-File -width 10000 $SrcDir\$path.3\cliglobalsettings1.txt
    gc $SrcDir\$path.3\cliglobalsettings1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\cliglobalsettings.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving sys global-settings data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys global-settings " | Out-File -width 10000 $SrcDir\$path.3\sysglobalsettings1.txt
    gc $SrcDir\$path.3\sysglobalsettings1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\sysglobalsettings.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving syslog data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys syslog " | Out-File -width 10000 $SrcDir\$path.3\syslog1.txt
    gc $SrcDir\$path.3\syslog1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\syslog.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cm device data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "cm device " | Out-File -width 10000 $SrcDir\$path.3\cmdevice1.txt
    gc $SrcDir\$path.3\cmdevice1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\cmdevice.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cm device-group data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "cm device-group " | Out-File -width 10000 $SrcDir\$path.3\cmdevicegroup1.txt
    gc $SrcDir\$path.3\cmdevicegroup1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\cmdevicegroup.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cm traffic-group data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "cm traffic-group " | Out-File -width 10000 $SrcDir\$path.3\cmtrafficgroup1.txt
    gc $SrcDir\$path.3\cmtrafficgroup1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\cmtrafficgroup.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cm trust-domain data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "cm trust-domain " | Out-File -width 10000 $SrcDir\$path.3\cmtrustdomain1.txt
    gc $SrcDir\$path.3\cmtrustdomain1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\cmtrustdomain.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving net route-domain data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "net route-domain " | Out-File -width 10000 $SrcDir\$path.3\netroutedomain1.txt
    gc $SrcDir\$path.3\netroutedomain1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\netroutedomain.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving net tunnels data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "net tunnels " | Out-File -width 10000 $SrcDir\$path.3\nettunnels1.txt
    gc $SrcDir\$path.3\nettunnels1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\nettunnels.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving sys state-mirroring data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "sys state-mirroring " | Out-File -width 10000 $SrcDir\$path.3\sysstatemirroring1.txt
    gc $SrcDir\$path.3\sysstatemirroring1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\sysstatemirroring.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving security firewall management-ip-rules data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "security firewall management-ip-rules " | Out-File -width 10000 $SrcDir\$path.3\securityfwmgmtip1.txt
    gc $SrcDir\$path.3\securityfwmgmtip1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\securityfwmgmtip.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving auth partition data"
    gc $SrcDir\$path.3\bigip_base_output.txt -ReadCount 10000 | % { $_ } | sls "auth partition " | Out-File -width 10000 $SrcDir\$path.3\partition1.txt
    gc $SrcDir\$path.3\partition1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\partition.txt


# Extracts config from bigip_user.conf

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving user data"
    gc $SrcDir\$path.3\bigip_user_output.txt -ReadCount 10000 | % { $_ } | sls "auth user " | Out-File -width 10000 $SrcDir\$path.3\user1.txt
    gc $SrcDir\$path.3\user1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\user.txt

# Extracts config from BigDB.dat

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving UI.logAccess data"
    gc $SrcDir\$path.3\BigDB_output.txt -ReadCount 10000 | % { $_ } | sls "UI.logAccess" | Out-File -width 10000 $SrcDir\$path.3\uilogaccess1.txt
    gc $SrcDir\$path.3\uilogaccess1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\uilogaccess.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving Config.Auditing data"
    gc $SrcDir\$path.3\BigDB_output.txt -ReadCount 10000 | % { $_ } | sls "Config.Auditing" | Out-File -width 10000 $SrcDir\$path.3\configauditing1.txt
    gc $SrcDir\$path.3\configauditing1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\configauditing.txt

# Extracts ASM Policy Name from var\tmp\ts_db.save_dir_*.cstmp\ts_db.data.PLC.PL_POLICIES.cstmp

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving ASM data"

    # Check if the directory pattern exists
    $folderName = gci $SrcDir\$path.3\var\tmp\ts_db.save_dir_* -Directory | ForEach-Object { $_.Name }

    if ($folderName) {
        $policyFile = "$SrcDir\$path.3\var\tmp\$folderName\ts_db.data.PLC.PL_POLICIES.cstmp"
 
    # Check if the policy file exists
    if (Test-Path -Path $policyFile) {
        Get-Content $policyFile | Out-File -Width 10000 "$DstDir\$path\asmpolicy.txt"
    } else {
        Write-Host -ForegroundColor Yellow "ASM policy file not found at $policyFile - skipping this section"
    }
    } else {
        Write-Host -ForegroundColor Yellow "No ts_db.save_dir_* directory found in $SrcDir\$path.3\var\tmp - skipping ASM data retrieval"
    }

#    gc $SrcDir\$path.3\var\tmp\$folderName\ts_db.data.PLC.PL_POLICIES.cstmp  | Out-File -width 10000 $DstDir\$path\asmpolicy.txt

# Extracts config from bigip_gtm.conf

if (Test-Path -Path $bigipgtmPath) {

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm datacenter data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm datacenter" | Out-File -width 10000 $SrcDir\$path.3\gtmdatacenter1.txt
    gc $SrcDir\$path.3\gtmdatacenter1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000  $DstDir\$path\gtmdatacenter.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm prober-pool data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm prober-pool" | Out-File -width 10000 $SrcDir\$path.3\gtmproberltmpool1.txt
    gc $SrcDir\$path.3\gtmproberltmpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000  $DstDir\$path\gtmproberltmpool.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm region data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm region" | Out-File -width 10000 $SrcDir\$path.3\gtmregion1.txt
    gc $SrcDir\$path.3\gtmregion1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000  $DstDir\$path\gtmregion.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm server data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm server" | Out-File -width 10000 $SrcDir\$path.3\gtmserver1.txt
    gc $SrcDir\$path.3\gtmserver1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\gtmserver.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm topology data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm topology" | Out-File -width 10000 $SrcDir\$path.3\gtmtopology1.txt
    gc $SrcDir\$path.3\gtmtopology1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\gtmtopology.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm global-settings data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm global-settings" | Out-File -width 10000 $SrcDir\$path.3\gtmglobalsettings1.txt
    gc $SrcDir\$path.3\gtmglobalsettings1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\gtmglobalsettings.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm monitor data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm monitor" | Out-File -width 10000 $SrcDir\$path.3\gtmmonitor1.txt
    gc $SrcDir\$path.3\gtmmonitor1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\gtmmonitor.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm pool data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm pool" | Out-File -width 10000 $SrcDir\$path.3\gtmltmpool1.txt
    gc $SrcDir\$path.3\gtmltmpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\gtmltmpool.txt

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm wideip data"
    gc $SrcDir\$path.3\bigip_gtm_output.txt -ReadCount 10000 | % { $_ } | sls "gtm wideip" | Out-File -width 10000 $SrcDir\$path.3\gtmwideip1.txt
    gc $SrcDir\$path.3\gtmwideip1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\gtmwideip.txt

    } else {
    Write-Host -ForegroundColor Yellow "Config file not found at $bigipgtmPath - skipping this section"
    }

# Extracts config from /config/partitions/<partitionname>/bigip.conf

    $partitions = $null
    $partitions = gci -name -directory $SrcDir\$path.3\config\partitions
    $partitions

    Foreach ($partition in $partitions) {
        $bigippartitionPath = "$SrcDir\$path.3\config\partitions\$partition\bigip.conf"
        $bigippartitionPath
        $bigippartitionoutputPath = "$SrcDir\$path.3\config\partitions\$partition\bigip_output.txt"
        $bigippartitionoutputPath

        $lines = $null
        $line2 = $null
        $lines = gc $bigippartitionPath

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing multiline to single line config"

        Foreach ($line in $lines) {
            if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("}")) {
                if ( $line  -match ("^[a-z].*") -and $line.EndsWith("}")) {
                    $line2 += $line+"`r`n"
                    } elseif ( $line  -match ("^[a-z].*")) {
                    $line2 += $line -replace '\s+',' '
                    } elseif ( $line  -match ("^[A-Z]")) {
                    $line2 += $line -replace '\s+',' '
                    } elseif ( $line.StartsWith(" ")) {
                    $line2 += $line -replace '\s+',' '
                    } elseif ( $line.StartsWith("}")) {
                    $line2 += $line+"`r`n"
                }
            }
        }

        $line2 | Out-File -width 10000 $bigippartitionoutputPath

#        Read-Host -Prompt "Press Enter to continue"        

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $bigippartitionoutputPath"

        New-Item -Path "$DstDir\$path\config\partitions\$partition" -ItemType Directory

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving profile data"
        gc $SrcDir\$path.3\config\partitions\$partition\bigip.conf -ReadCount 10000 | % { $_ } | sls "ltm profile" | ? {$_ -match '\S'} | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmprofile1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmprofile1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmprofile.txt
        gc $DstDir\$path\config\partitions\$partition\ltmprofile.txt | ac $DstDir\$path\ltmprofile.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving route data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "net route" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\route1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\route1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\route.txt
        gc $DstDir\$path\config\partitions\$partition\route.txt | ac $DstDir\$path\route.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving node data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm node" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmnode1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmnode1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmnode.txt
        gc $DstDir\$path\config\partitions\$partition\ltmnode.txt | ac $DstDir\$path\ltmnode.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving pool data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm pool" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmpool1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmpool.txt
        gc $DstDir\$path\config\partitions\$partition\ltmpool.txt | ac $DstDir\$path\ltmpool.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving virtual server data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm virtual " | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmvirtualserver1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmvirtualserver1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmvirtualserver.txt
        gc $DstDir\$path\config\partitions\$partition\ltmvirtualserver.txt | ac $DstDir\$path\ltmvirtualserver.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving virtual address data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm virtual-address " | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmvirtualaddress1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmvirtualaddress1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmvirtualaddress.txt
        gc $DstDir\$path\config\partitions\$partition\ltmvirtualaddress.txt | ac $DstDir\$path\ltmvirtualaddress.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving client SSL data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm profile client-ssl" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmclientssl1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmclientssl1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmclientssl.txt
        gc $DstDir\$path\config\partitions\$partition\ltmclientssl.txt | ac $DstDir\$path\ltmclientssl.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving server SSL data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm profile server-ssl" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmserverssl1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmserverssl1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmserverssl.txt
        gc $DstDir\$path\config\partitions\$partition\ltmserverssl1.txt | ac $DstDir\$path\ltmserverssl.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving certificate data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "certificate_d" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\cert1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\cert1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\cert.txt
        gc $DstDir\$path\config\partitions\$partition\cert.txt | ac $DstDir\$path\cert.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving key data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "certificate_key_d" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\key1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\key1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\key.txt
        gc $DstDir\$path\config\partitions\$partition\key.txt | ac $DstDir\$path\key.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cipher group"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm cipher group " | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmciphergroup1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmciphergroup1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmciphergroup.txt
        gc $DstDir\$path\config\partitions\$partition\ltmciphergroup.txt | ac $DstDir\$path\ltmciphergroup.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving cipher rule"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm cipher rule " | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmcipherrule1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmcipherrule1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmcipherrule.txt
        gc $DstDir\$path\config\partitions\$partition\ltmcipherrule.txt | ac $DstDir\$path\ltmcipherrule.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving monitor data"
        gc $bigippartitionoutputPath -ReadCount 10000 | % { $_ } | sls "ltm monitor " | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\ltmmonitor1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\ltmmonitor1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\ltmmonitor.txt
        gc $DstDir\$path\config\partitions\$partition\ltmmonitor.txt | ac $DstDir\$path\ltmmonitor.txt

    }

# Extracts config from /config/partitions/<partitionname>/bigip_gtm.conf

    $partitions = $null
    $partitions = gci -name -directory $SrcDir\$path.3\config\partitions
    $partitions

    Foreach ($partition in $partitions) {
        $bigipgtmpartitionPath = "$SrcDir\$path.3\config\partitions\$partition\bigip_gtm.conf"
        $bigipgtmpartitionPath
        $bigipgtmpartitionoutputPath = "$SrcDir\$path.3\config\partitions\$partition\bigip_gtm_output.txt"
        $bigipgtmpartitionoutputPath

        $lines = $null
        $line2 = $null
        $lines = gc $bigipgtmpartitionPath

        if (Test-Path -Path $bigipgtmpartitionPath) {
            $lines = $null
            $line2 = $null
            $lines = Get-Content $bigipgtmpartitionPath

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Collapsing multiline to single line config"

            Foreach ($line in $lines) {
                if ( $line  -match ("^[a-z].*") -or $line  -match ("^[A-Z].*") -or $line.StartsWith(" ") -or $line.StartsWith("}")) {
                    if ( $line  -match ("^[a-z].*") -and $line.EndsWith("}")) {
                        $line2 += $line+"`r`n"
                        } elseif ( $line  -match ("^[a-z].*")) {
                        $line2 += $line -replace '\s+',' '
                        } elseif ( $line  -match ("^[A-Z]")) {
                        $line2 += $line -replace '\s+',' '
                        } elseif ( $line.StartsWith(" ")) {
                        $line2 += $line -replace '\s+',' '
                        } elseif ( $line.StartsWith("}")) {
                        $line2 += $line+"`r`n"
                    }
                }
            }

        $line2 | Out-File -width 10000 $bigipgtmpartitionoutputPath

#        Read-Host -Prompt "Press Enter to continue"        

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Single line config output saved to $bigippartitionoutputPath"

#        New-Item -Path "$DstDir\$path\config\partitions\$partition" -ItemType Directory

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm datacenter data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm datacenter" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmdatacenter1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmdatacenter1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmdatacenter.txt
        gc $DstDir\$path\config\partitions\$partition\gtmdatacenter.txt | ac $DstDir\$path\gtmdatacenter.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm prober-pool data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm prober-pool" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmproberltmpool1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmproberltmpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmproberltmpool.txt
        gc $DstDir\$path\config\partitions\$partition\gtmproberltmpool.txt | ac $DstDir\$path\gtmproberltmpool.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm region data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm region" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmregion1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmregion1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmregion.txt
        gc $DstDir\$path\config\partitions\$partition\gtmregion.txt | ac $DstDir\$path\gtmregion.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm server data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm server" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmserver1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmserver1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmserver.txt
        gc $DstDir\$path\config\partitions\$partition\gtmserver.txt | ac $DstDir\$path\gtmserver.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm topology data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm topology" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmtopology1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmtopology1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmtopology.txt
        gc $DstDir\$path\config\partitions\$partition\gtmtopology.txt | ac $DstDir\$path\gtmtopology.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm global-settings data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm global-settings" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmglobalsettings1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmglobalsettings1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmglobalsettings.txt
        gc $DstDir\$path\config\partitions\$partition\gtmglobalsettings.txt | ac $DstDir\$path\gtmglobalsettings.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm monitor data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm monitor" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmmonitor1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmmonitor1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmmonitor.txt
        gc $DstDir\$path\config\partitions\$partition\gtmmonitor.txt | ac $DstDir\$path\gtmmonitor.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm pool data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm pool" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmltmpool1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmltmpool1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmltmpool.txt
        gc $DstDir\$path\config\partitions\$partition\gtmltmpool.txt | ac $DstDir\$path\gtmltmpool.txt

        Write-Host -ForegroundColor black -BackgroundColor Yellow "Retrieving gtm wideip data"
        gc $bigipgtmpartitionoutputPath -ReadCount 10000 | % { $_ } | sls "gtm wideip" | Out-File -width 10000 $SrcDir\$path.3\config\partitions\$partition\gtmwideip1.txt
        gc $SrcDir\$path.3\config\partitions\$partition\gtmwideip1.txt | Sort | gu | ? {$_ -match '\S'}  | Out-File -width 10000 $DstDir\$path\config\partitions\$partition\gtmwideip.txt
        gc $DstDir\$path\config\partitions\$partition\gtmwideip.txt | ac $DstDir\$path\gtmwideip.txt

        } else {
            Write-Host -ForegroundColor Yellow "Config file not found at $bigipgtmpartitionPath - skipping this section"
        }
    }

# Part 2: Populate extracted config into excel

    # At script start - kill any existing Excel processes

    Stop-Process -Name "EXCEL" -ErrorAction SilentlyContinue

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Launching Excel application"
    $excel = New-Object -comobject Excel.Application
    $excel.WindowState= "xlMaximized"
    $excel.Visible = $true
    $FilePath = "$SrcDir\$ReportTemplate1"
    $Workbook = $excel.Workbooks.Open($FilePath)
    cd $DstDir\$path
    $filenames = gci -name -Filter *.txt

    Foreach ($file in $filenames) {
        # Skip empty files
        if (-not (Test-Path -Path "$DstDir\$path\$file" -PathType Leaf) -or (Get-Item "$DstDir\$path\$file").Length -eq 0) {
            Write-Host "Skipping empty file: $file"
            continue
        }

        $file2 = $file | % {$_ -replace '.txt',''}
        echo "Populating data in $file2 sheet…"

        $content = gc $DstDir\$path\$file

        $Sheet = $Workbook.sheets.item($file2)
        $Sheet.activate()
        $startCell = $Sheet.Cells.Item(2,1)
        $startCell.activate()
        Start-Sleep -m 100
        $rowIndex = 2
        foreach ($line in $content) {
            $sheet.Cells.Item($rowIndex, 1).Value2 = $line
            $rowIndex++
        }
    }

    $Report = "$path.xlsx"
    $Workbook.SaveAs("$DstDir\$Report")
    Write-Host -ForegroundColor black -BackgroundColor Yellow "Report saved to $DstDir\$Report"
    $Workbook.close($false)
    $Excel.Quit()
    Clear-Variable line2

# Part 3: Extract certificate information

# Install the ImportExcel module if not already installed
    if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
        Install-Module -Name ImportExcel -Scope CurrentUser
    }

# Extract certificate information from \var\tmp\filestore_temp\files_d\Common_d\certificate_d directory

# Define the folder containing certificates
$certFolder = "$SrcDir\$path.3\var\tmp\filestore_temp\files_d\Common_d\certificate_d"

# Define the output Excel file path
$outputExcel = "$DstDir\$path\certdetails.xlsx"

# Initialize an array to store certificate details
$certDetails = @()

# Get all certificate files in the folder
$certFiles = Get-ChildItem -Path $certFolder -File | Where-Object { $_.Length -ne 0 }

    foreach ($certFile in $certFiles) {
        Write-Host "Processing certificate: $($certFile.FullName)" -ForegroundColor Cyan
    
        try {
            # Load the certificate
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile.FullName)
    
            # Extract SAN (Subject Alternative Name)
            $sanExtension = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
            $sanList = if ($sanExtension) {
                ($sanExtension.Format($true) -split "`n" | Where-Object { $_ -match "DNS Name" }) -join ", "
            } else {
                "No SAN Found"
            }
    
            # Create an object for the certificate details
            $certDetails += [PSCustomObject]@{
                FileName      = $certFile.Name
                FilePath      = $certFile.FullName
                Subject       = $cert.Subject
                Issuer        = $cert.Issuer
                Thumbprint    = $cert.Thumbprint
                NotBefore     = $cert.NotBefore
                NotAfter      = $cert.NotAfter
                SAN           = $sanList
            }
        } catch {
            Write-Host "Failed to process $($certFile.Name): $_" -ForegroundColor Red
        }
    
        Write-Host "----------------------------------------"
    }

# Export the certificate details into Excel
    $certDetails | Export-Excel -Path $outputExcel -AutoSize

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Certificate details exported to $outputExcel"

# Define paths
    $sourceFile = "$DstDir\$path\certdetails.xlsx"
    $targetFile = "$DstDir\$Report"
    $sourceSheetName = "Sheet1"
    $targetSheetName = "certdetails"  # Existing sheet in target file

# Start Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

# Open workbooks
    $wbSource = $excel.Workbooks.Open($sourceFile)
    $wbTarget = $excel.Workbooks.Open($targetFile)

# Get sheets
    $wsSource = $wbSource.Sheets.Item($sourceSheetName)
    $wsTarget = $wbTarget.Sheets.Item($targetSheetName)

# Get used range from source
    $usedRange = $wsSource.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $sourceValue = $wsSource.Cells.Item($r, $c).Value2
            $wsTarget.Cells.Item($r, $c).Value2 = $sourceValue.ToString()
        }
    }

# Save and clean up
    $wbTarget.Save()
    $wbSource.Close($false)
    $wbTarget.Close($true)
    $excel.Quit()

# Extract device certificate information from \var\tmp\cert_temp\conf\ssl.crt file

# Define the folder containing certificates
    $certFolder = "$SrcDir\$path.3\var\tmp\cert_temp\conf\ssl.crt"

# Define the output Excel file path
    $outputExcel = "$DstDir\$path\devicecert.xlsx"

# Initialize an array to store certificate details
    $certDetails = @()

# Get all certificate files in the folder
    $certFiles = Get-ChildItem -Path $certFolder -File | Where-Object { $_.Length -ne 0 }

    foreach ($certFile in $certFiles) {
        Write-Host "Processing certificate: $($certFile.FullName)" -ForegroundColor Cyan
    
        try {
            # Load the certificate
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile.FullName)
    
            # Extract SAN (Subject Alternative Name)
            $sanExtension = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
            $sanList = if ($sanExtension) {
                ($sanExtension.Format($true) -split "`n" | Where-Object { $_ -match "DNS Name" }) -join ", "
            } else {
                "No SAN Found"
            }
    
            # Create an object for the certificate details
            $certDetails += [PSCustomObject]@{
                FileName      = $certFile.Name
                FilePath      = $certFile.FullName
                Subject       = $cert.Subject
                Issuer        = $cert.Issuer
                Thumbprint    = $cert.Thumbprint
                NotBefore     = $cert.NotBefore
                NotAfter      = $cert.NotAfter
                SAN           = $sanList
            }
        } catch {
            Write-Host "Failed to process $($certFile.Name): $_" -ForegroundColor Red
        }
    
        Write-Host "----------------------------------------"
    }

# Export the certificate details into Excel
    $certDetails | Export-Excel -Path $outputExcel -AutoSize

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Certificate details exported to $outputExcel"

# Define paths
    $sourceFile = "$DstDir\$path\devicecert.xlsx"
    $targetFile = "$DstDir\$Report"
    $sourceSheetName = "Sheet1"
    $targetSheetName = "devicecert"  # Existing sheet in target file

# Start Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false

# Open workbooks
    $wbSource = $excel.Workbooks.Open($sourceFile)
    $wbTarget = $excel.Workbooks.Open($targetFile)

# Get sheets
    $wsSource = $wbSource.Sheets.Item($sourceSheetName)
    $wsTarget = $wbTarget.Sheets.Item($targetSheetName)

# Get used range from source
    $usedRange = $wsSource.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count
     
    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $sourceValue = $wsSource.Cells.Item($r, $c).Value2
            $wsTarget.Cells.Item($r, $c).Value2 = $sourceValue.ToString()
        }
    }

# Save and clean up
    $wbTarget.Save()
    $wbSource.Close($false)
    $wbTarget.Close($true)
    $excel.Quit()

# Extract certificate information from \var\tmp\gtm_tmp\gtm directory

# Define the folder containing certificates
    $certFolder = "$SrcDir\$path.3\var\tmp\gtm_tmp\gtm"

# Define the output Excel file path
    $outputExcel = "$DstDir\$path\gtmcert.xlsx"

# Initialize an array to store certificate details
    $certDetails = @()

# Get all certificate files in the folder
    $certFiles = Get-ChildItem -Path $certFolder -File | Where-Object { $_.Length -ne 0 }

    foreach ($certFile in $certFiles) {
        Write-Host "Processing certificate: $($certFile.FullName)" -ForegroundColor Cyan
    
        try {
            # Load the certificate
            $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certFile.FullName)
    
            # Extract SAN (Subject Alternative Name)
            $sanExtension = $cert.Extensions | Where-Object { $_.Oid.FriendlyName -eq "Subject Alternative Name" }
            $sanList = if ($sanExtension) {
                ($sanExtension.Format($true) -split "`n" | Where-Object { $_ -match "DNS Name" }) -join ", "
            } else {
                "No SAN Found"
            }
    
            # Create an object for the certificate details
            $certDetails += [PSCustomObject]@{
                FileName      = $certFile.Name
                FilePath      = $certFile.FullName
                Subject       = $cert.Subject
                Issuer        = $cert.Issuer
                Thumbprint    = $cert.Thumbprint
                NotBefore     = $cert.NotBefore
                NotAfter      = $cert.NotAfter
                SAN           = $sanList
            }
        } catch {
            Write-Host "Failed to process $($certFile.Name): $_" -ForegroundColor Red
        }

        Write-Host "----------------------------------------"
    }

# Export the certificate details to an Excel file
    $certDetails | Export-Excel -Path $outputExcel -AutoSize

    Write-Host -ForegroundColor black -BackgroundColor Yellow "Certificate details exported to $outputExcel"

    # Define paths
    $sourceFile = "$DstDir\$path\gtmcert.xlsx"
    $targetFile = "$DstDir\$Report"
    $sourceSheetName = "Sheet1"
    $targetSheetName = "gtmcert"  # Existing sheet in target file
    
    # Start Excel
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    
    # Open workbooks
    $wbSource = $excel.Workbooks.Open($sourceFile)
    $wbTarget = $excel.Workbooks.Open($targetFile)
    
    # Get sheets
    $wsSource = $wbSource.Sheets.Item($sourceSheetName)
    $wsTarget = $wbTarget.Sheets.Item($targetSheetName)
    
    # Get used range from source
    $usedRange = $wsSource.UsedRange
    $rowCount = $usedRange.Rows.Count
    $colCount = $usedRange.Columns.Count

    for ($r = 1; $r -le $rowCount; $r++) {
        for ($c = 1; $c -le $colCount; $c++) {
            $sourceValue = $wsSource.Cells.Item($r, $c).Value2
            $wsTarget.Cells.Item($r, $c).Value2 = $sourceValue.ToString()
        }
    }

    # Save and clean up
    $wbTarget.Save()
    $wbSource.Close($false)
    $wbTarget.Close($true)
    $excel.Quit()
}

(Get-Date)-$StartTime
