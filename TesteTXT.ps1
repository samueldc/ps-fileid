﻿Get-Content -Path "C:\sepac-files\ditecw10.txt" | ForEach-Object { Add-ADGroupMember -Identity DitecW10 -Members $PSItem }