Write-Host "Loading RGS holidays...." -NoNewline
$HolidaysSetsArr = Get-CsRgsHolidaySet | ?{$_.ownerpool -eq "inh-sfb-eep-2.eu.boehringer.com"}
Write-Host " Done" -ForegroundColor DarkGreen -NoNewline;Write-host "`n"

foreach($HolidaySet in $HolidaysSetsArr){    
    if($HolidaySet.holidaylist | ?{$_.startdate -like "*1/01/2019*"}){
        write "[OK] >> $($HolidaySet.name)"
        $HolidaySet.holidaylist | ?{$_.startdate -like "*1/01/2019*"}
    }else{
        write "[NOK] >> $($HolidaySet.name)"
        Write-Host "Adding..." -NoNewline
        <#
        $x = New-CsRgsHoliday -StartDate "01/01/2019 00:00" -EndDate "01/01/2019 23:59" -Name "01012019"
        $HolidaySet.holidaylist.Add($x)
        Set-CsRgsHolidaySet -Instance $HolidaySet | OUT-NULL
        #>
        Write-Host " Done" -ForegroundColor DarkGreen -NoNewline; Write-host ""

        if ($PSVersionTable.PSVersion.Major -ge 3) {Pause}
        else{
            $HOST.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown") | OUT-NULL
            $HOST.UI.RawUI.Flushinputbuffer()
        }
    }
}
