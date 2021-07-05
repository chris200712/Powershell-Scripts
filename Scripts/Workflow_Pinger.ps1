Workflow Test-PingWF{
    param([string[]]$iprange)

    foreach -parallel($ip in $iprange)
    {
        "Pinging: $ip"
        Test-Connection -ipaddres $ip -Count 1 -ErrorAction SilentlyContinue
    }
}

Test-PingWF -iprange (1..20 | % {"192.168.1."+$_})