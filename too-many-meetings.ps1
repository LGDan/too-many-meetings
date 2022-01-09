#!/usr/bin/pwsh

<#
MIT License

Copyright (c) 2022 LGDan

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.
#>

function Set-MSGraphAPIToken {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute(
        'PSUseDeclaredVarsMoreThanAssignments',
        '',
        Justification='MS Graph token is used outside function scope.'
    )]
    [CmdletBinding()]
    param (
        [Parameter()][String]$Token
    )
    
    $Global:MSGraphToken = $Token
}

function Invoke-MSGraph {
    [CmdletBinding()]
    param (
        [Parameter()][String]$URL,
        [Parameter()][String]$Method,
        [Parameter()]$Body
    )
    if ($null -eq $Global:MSGraphToken) {
        "Use Set-MSGraphAPIToken -Token <Token>"
    }else{
        Invoke-RestMethod -Uri $URL -Method $Method -Body $Body -Headers @{
            Authorization = "Bearer $Global:MSGraphToken"
        }
    }
}

function Build-URL {
    [CmdletBinding()]
    param (
        [Parameter()][String]$Protocol,
        [Parameter()][String]$Domain,
        [Parameter()][String]$Path,
        [Parameter()][hashtable]$GetParameters
    )
    $sb = [System.Text.StringBuilder]::new()
    $sb.Append($Protocol + "://") | Out-Null
    $sb.Append($Domain) | Out-Null
    $sb.Append($Path) | Out-Null
    if ($null -ne $GetParameters) {$sb.Append("?")|Out-Null}
    $GetParameters.GetEnumerator() | ForEach-Object {
        $sb.Append(($_.Key + "=" + $_.Value)) | Out-Null
        $sb.Append("&") | Out-Null
    }
    $sb.Remove(($sb.Length-1),1) | Out-Null
    $sb.ToString()
}

function Get-ThisWeek {
    $monday = (Get-Date).Subtract(
        # Always gets monday, but with current time.
        [Timespan]::FromDays(
            (Get-Date).DayOfWeek.value__-1
        )
    ).Subtract(
        # Negate current time to get 00:00:00
        (Get-Date).TimeOfDay.Subtract([timespan]::FromSeconds(1))
    )

    $friday = $monday.AddDays(5)
    $friday = $friday.Subtract([timespan]::FromSeconds(1))

    [ordered]@{
        StartDate = $monday
        EndDate = $friday
    }
}

function Get-CalendarDaysIntoTheFuture {
    [CmdletBinding()]
    param (
        [Parameter()][Int]$DaysIntoTheFuture
    )
    Get-Calendar `
        -StartDate (Get-Date) `
        -EndDate (Get-Date).AddDays($DaysIntoTheFuture)
}

function Get-Calendar {
    [CmdletBinding()]
    param (
        [Parameter()][datetime]$StartDate,
        [Parameter()][datetime]$EndDate
    )
    # https://graph.microsoft.com
    # /v1.0/me/calendarview
    # ?startdatetime=2022-01-08T15:56:22.840Z
    # &enddatetime=2022-01-15T15:56:22.840Z
    $url = Build-URL `
        -Protocol "https" `
        -Domain "graph.microsoft.com" `
        -Path "/v1.0/me/calendarview" `
        -GetParameters ([ordered]@{
        startdatetime = $StartDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
        enddatetime = $EndDate.ToString("yyyy-MM-ddTHH:mm:ss.fffZ")
    })
    $allData = [System.Collections.ArrayList]::new()
    $calendarData = Invoke-MSGraph -URL $url -Method "GET"
    $allData.Add($calendarData) | Out-Null
    if ($null -ne $calendarData."@odata.nextLink") {
        Do {
            $calendarData = Invoke-MSGraph `
                -URL $calendarData."@odata.nextLink" `
                -Method "GET"
            $allData.Add($calendarData) | Out-Null
        } While ($null -ne $calendarData."@odata.nextLink")
    }
    $allData | ForEach-Object {
        $_.value
    }
}


# Why are these two functions below the same operation, but named differently?
# Because the same operation is used for two different reasons, and it
# increases readability for those who aren't familiar with the code.
function Get-MinutesOfTheWorkingDay($StartTime, $EndTime) {
    ([datetime]::Parse($EndTime) - [datetime]::Parse($StartTime)).TotalMinutes
}

function Get-MinuteBasedOnStart($StartTime, $InputTime) {
    ([datetime]::Parse($InputTime) - [datetime]::Parse($StartTime)).TotalMinutes
}

function Get-MeetingEvents {
    [CmdletBinding()]
    param (
        [Parameter()]
        [Int]
        $DaysIntoTheFuture
    )
    <#
    This is just a filter query to avoid any calendar appointments that 
    aren't disruptive / don't count toward suffering (i.e. lunch).
    
    Add in any more filters here to exclude any other events.

    Full list of stuff to filter by:
https://docs.microsoft.com/en-us/graph/api/resources/event?view=graph-rest-1.0
    #>
    $week = Get-ThisWeek
    Get-Calendar -StartDate $week.StartDate -EndDate $week.EndDate
    | Where-Object categories -NotContains "No Conflicts" # Unique to my cal.
    | Select-Object start, end, subject, responseStatus 
    | ForEach-Object {
        [pscustomobject]@{
            date = ([System.DateTime]$_.start.dateTime).ToShortDateString()
            event = $_
        }
    }
}

function Get-HowMuchTimeCanIActuallyDoWork {
    [CmdletBinding()]
    param (
        [Parameter()][Hashtable]$WorkingHoursTable,
        [Parameter()][timespan]$BreaktimeDuration
    )

    # Get all the calendar events up and coming, and sort them by the day.
    Get-MeetingEvents
    | Group-Object date 
    | ForEach-Object {
        # For each day...
        $events = $_.Group
        $date = $_.Name

        # Get the facts about the day, and preload a list of consumable minutes.
        $dayText = [datetime]::Parse($date).ToString("dddd")
        $dayStart = $WorkingHoursTable[$dayText][0]
        $dayEnd = $WorkingHoursTable[$dayText][1]
        $breakTimeMinutes = $BreaktimeDuration.TotalMinutes
        $totalMinutesToday = Get-MinutesOfTheWorkingDay `
            -StartTime $dayStart `
            -EndTime $dayEnd
        $minutesList = [System.Collections.Generic.list[Int]]::new()
        1..$totalMinutesToday | ForEach-Object {$minutesList.Add($_)|Out-Null}

        # Remove 'used' minutes from the list of available minutes in the 
        # working day. This means that double-booked time does not matter.
        foreach ($event in $events) {
            $meetingStartMinutes = Get-MinuteBasedOnStart `
                -StartTime $dayStart `
                -InputTime $event.event.start.dateTime.ToString("HH:mm")
            $meetingDuration = (
                $event.event.end.dateTime - $event.event.start.dateTime
            )
            $meetingEndMinutes = (
                $meetingStartMinutes + $meetingDuration.TotalMinutes
            )
            $meetingStartMinutes..($meetingEndMinutes-1) | ForEach-Object {
                $minutesList.Remove($_) | Out-Null
            }
        }

        # Now all consumed minutes have been removed from the list of available
        # minutes, lets add up what's left over to find out how many 'free'
        # miuntes there are left to do work.
        $availableTime = [timespan]::FromMinutes(
            ($minutesList.Count - ($breakTimeMinutes))
        )
        $unavailableTime = [timespan]::FromMinutes(
            $totalMinutesToday - $minutesList.Count
        )

        # Finally spit out an object containing all the stats for this day.
        [PSCustomObject]@{
            Date = $date
            TotalTimeInMeetings = $unavailableTime
            TotalTimeFree = $availableTime
            EventCount = ($events | Measure-Object).Count
            PercentInMeetings = (
                $unavailableTime.TotalMinutes / $totalMinutesToday
            ).ToString("P") # Percentage formatting w/auto dec place?! Who knew!
            PercentFree = (
                $availableTime.TotalMinutes / $totalMinutesToday
            ).ToString("P")
        }
    }    
}

function Get-ThePercentageOfMyWeekSpentInMeetings {
    # Works out the percentage of your working week that you spend in meetings.
    [CmdletBinding()]
    param (
        [Parameter()][Int]$WorkingHoursPerWeek,
        [Parameter()][Hashtable]$WorkingHoursTable,
        [Parameter()][timespan]$BreaktimeDuration
    )

    $totalTimeInMeetings = [timespan]::Zero
    $totalTimeFree = [timespan]::Zero
    $workingHours = [timespan]::FromHours($WorkingHoursPerWeek)

    Get-HowMuchTimeCanIActuallyDoWork `
        -WorkingHoursTable $WorkingHoursTable `
        -BreaktimeDuration $BreaktimeDuration
    | ForEach-Object {
        $totalTimeInMeetings = $totalTimeInMeetings.Add($_.TotalTimeInMeetings)
        $totalTimeFree = $totalTimeFree.Add($_.TotalTimeFree)
    }

    [PSCustomObject]@{
        TimeInMeetings = $totalTimeInMeetings
        TimeFree = $totalTimeFree
        PercentageInMeetings = (
            $totalTimeInMeetings.TotalMinutes / $workingHours.TotalMinutes
        ).ToString("P")
        PercentageFree = (
            $totalTimeFree.TotalMinutes / $workingHours.TotalMinutes
        ).ToString("P")
    }
}

function Get-TotalWorkingHours {
    [CmdletBinding()]
    param (
        [Parameter()][Hashtable]$WorkingHoursTable,
        [Parameter()][timespan]$BreaktimeDuration
    )
    $totalHours = [timespan]::Zero

    $WorkingHoursTable.GetEnumerator() | ForEach-Object {
        $timeWorkedToday = (
            [datetime]::Parse($_.Value[1]) - [datetime]::Parse($_.Value[0])
        ).Subtract($BreaktimeDuration)
        $totalHours = $totalHours.Add($timeWorkedToday)
    }
    $totalHours
}

function Invoke-Pain {
    if ($null -ne $Global:MSGraphToken) {
        $myWorkingHours = @{
            Monday =    ("09:00","17:30")
            Tuesday =   ("09:00","17:30")
            Wednesday = ("09:00","17:30")
            Thursday =  ("09:00","17:00")
            Friday =    ("09:00","16:30")
        }

        $myBreakTime = [timespan]::FromHours(1)

        [timespan]$myWorkingHoursTotal = Get-TotalWorkingHours `
            -WorkingHoursTable $myWorkingHours `
            -BreaktimeDuration $myBreakTime

        Write-Output "------ Days of this week in meetings ----"

        Get-HowMuchTimeCanIActuallyDoWork `
            -WorkingHoursTable $myWorkingHours `
            -BreaktimeDuration $myBreakTime
        
        Write-Output "-----------------------------------------"
        Write-Output "------ Review of this week --------------"

        Get-ThePercentageOfMyWeekSpentInMeetings `
            -WorkingHoursPerWeek $myWorkingHoursTotal.TotalHours `
            -WorkingHoursTable $myWorkingHours `
            -BreaktimeDuration $myBreakTime

        Write-Output "-----------------------------------------"

    }else{
        Write-Host "Set your token with Set-MSGraphAPIToken -Token ""tokenfoo"""
    }
}

Invoke-Pain
