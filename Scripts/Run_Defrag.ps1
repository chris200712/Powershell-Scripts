﻿# Run-Defrag
# Defragments the targeted hard drives.
#
# Args:
#   $server: A target Server 2003 or 2008 system
#   $drive: An optional drive letter.  If this is blank then all 
#           drives are defragmented
#   $force: If this switch is set then a defrag will be forced
#           even if the drive is low on space
#
# Returns:
#   The result description for each drive.

function Run-Defrag {
  param([string]$server, [string]$drive, [switch]$force)

  [string]$query = 'Select * from Win32_Volume where DriveType = 3'

  if ($drive) {
    $query += " And DriveLetter LIKE '$drive%'"
  }

  $volumes = Get-WmiObject -Query $query -ComputerName $server

  foreach ($volume in $volumes) {
    Write-Host "Defragmenting $($volume.DriveLetter)..." -noNewLine
    $result = $volume.Defrag($force)
    switch ($result) {
      0 {'Success'}
      1 {'Access Denied'}
      2 {'Not Supported'}
      3 {'Volume Dirty Bit Set'}
      4 {'Not Enough Free Space'}
      5 {'Corrupt MFT Detected'}
      6 {'Call Cancelled'}
      7 {'Cancellation Request Requested Too Late'}
      8 {'Defrag In Progress'}
      9 {'Defrag Engine Unavailable'}
      10 {'Defrag Engine Error'}
      11 {'Unknown Error'}
    }
  }
}

Run-Defrag localhost