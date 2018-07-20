
Function Get-Office365BuildToVersionMap {

    <#

        The purpose of this function is to go to the Microsoft support sites and download the Version and Build information.  It will then
        map it into an array that can be referenced for other things or exported to a CSV.

        I use the support pages for each channel and year instead of the master table because it was easier to manipulate the standard support
        pages than it was to manipulate the single table with RegEx.  I'm not using the Parse HTML options of Invoke-WebReqeust to capture table
        cells or individual elements because Invoke-WebRequest hangs when downloading the support pages (seems to be a frequently reported issue).

        Author: Christopher Kibble
        Created: 2018-07-20

        Version 1.0 - Initial Build.

        To Do:

            * Bring in Insider Build Information.
            * Better Error Handling.

    #>

    [regex]$rxBuilds = '<p><em>Version (.*)<\/em><\/p>'                 # Regular Expression that Finds the Version/BUild Numbers from the Page
    $urlBase = "https://docs.microsoft.com/en-us/officeupdates"         # Start page for all Office Update pages by Year and Update Type
    $officeBuildList = @()                                              # Array to hold the list of Versions and Builds

    # Identify all the possible URLs from 2015 to the Current Year (future proofing?).  Not all of these channels existed in 2015, so a 404 is
    # expected on at least one of them.  There may also not be all pages for the current year if updates haven't been released yet.

    $urlList = @()
    ForEach($year in $yearList) {
        $urlList += "$urlBase/monthly-channel-$year"
        $urlList += "$urlBase/semi-annual-channel-$year"
        $urlList += "$urlBase/semi-annual-channel-targeted-$year"
    }

    ForEach($url in $urlList) {

        Try {
            $web = Invoke-WebRequest -Uri $url -UseBasicParsing
        } catch {
            <# #>
        }

        if($web.StatusCode -ne 200)  {
            Write-Information "$url Returned Error $($web.StatusCode)"
        } else {
            $content = $web.RawContent

            $rxMatches = $rxBuilds.matches($content)

            ForEach($entry in $rxMatches) {
                $buildLine = $entry.Groups[1].Value
                $buildNumber = $buildLine.substring(0,$buildLine.indexOf(' '))
                $versionNumber = $($buildLine.substring($buildLine.indexOf('Build') + 6)) -replace '\)',''
                [version]$versionNumber = "16.0.$versionNumber"

                $o = New-Object -TypeName PSObject

                Add-Member -InputObject $o -MemberType NoteProperty -Name BuildNumber -Value $buildNumber
                Add-Member -InputObject $o -MemberType NoteProperty -Name VersionNumber -Value $versionNumber
                Add-Member -InputObject $o -MemberType NoteProperty -Name Source -Value $url

                $officeBuildList += $o
            }
        }
    }

    Return $officeBuildList
}