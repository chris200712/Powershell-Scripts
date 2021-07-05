# Like this.....

$StartDate=(Get-Date)
$EndDate=[datetime]”08/19/1966 00:00”
New-Timespan –Start $StartDate –End $EndDate

# Or.....

New-TimeSpan -End 1-Jan-2014

# Or.....

New-TimeSpan -End ”01/01/2018 00:00”

# Or.....

New-TimeSpan -End 2018-01-01

# A truth test

(get-date 2010-01-02) -lt (get-date 2010-01-01)
