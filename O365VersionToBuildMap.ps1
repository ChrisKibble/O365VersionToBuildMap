
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
        Version 2.0 - Removed Source - No value and hurts removing values not unique.
                    - Returned unique list of resources instead of including the duplicates.

        To Do:

            * Bring in Insider Build Information.
            * Better Error Handling.

    #>

    [regex]$rxBuilds = '<p><em>Version (.*)<\/em><\/p>'                 # Regular Expression that Finds the Version/BUild Numbers from the Page
    $urlBase = "https://docs.microsoft.com/en-us/officeupdates"         # Start page for all Office Update pages by Year and Update Type
    $officeBuildList = @()                                              # Array to hold the list of Versions and Builds
    $yearList = 2015..$(get-date).Year                                  # List of all years from 2015 to now.

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

                $officeBuildList += $o
            }
        }
    }

    $officeBuildList = $officeBuildList | Sort-Object -Property VersionNumber -Unique

    Return $officeBuildList
}

Function Get-OfficeBuildListToCase {

    <#

        The purpose of this function is to use the build list generated by Get-Office365BuildToVersionMap and to
        generate SQL Case Statements for Configuration Manager (SCCM) SQL reporting or adhoc queries.

        Author: Christopher Kibble
        Created: 2018-07-20

        Version 1.0 - Initial Build.

        To Do:

            * Bring in Insider Build Information.
            * Better Error Handling.

    #>

    Param(
        $officeBuildList
    )

    $sql = "case "
    ForEach($build in $officeBuildList) {
        $sql += "when v_GS_OFFICE365PROPLUSCONFIGURATIONS.VersionToReport0 = '$($build.versionNumber)' then '$($build.buildNumber)'`r`n"
    }
    $sql += "else 'Unknown' end as Office365Build"
    Return $sql
}

Function Get-OfficeBuildListToSQLTable {

    <#

        The purpose of this function is to use the build list generated by Get-Office365BuildToVersionMap and to
        generate drop, create, and insert statements to build a mapping table that you can link to Configuration
        Manager (SCCM) for reporting or adhoc queries.

        Author: Christopher Kibble
        Created: 2018-07-20

        Version 1.0 - Initial Build.

        To Do:

            * Validate table name?  Not sure this is necessary.
            * Use SQL Parameters instead of straight text?  Not sure this is necessary, either.
            * Document my parameters.

    #>

    Param(
        $officeBuildList,
        $tableName = "O365BuildToVersionMap",
        [switch]$dontDropTable,
        [switch]$deleteExistingRecords
    )

    $sql = ""
    $sql += "-- You will need to uncomment the code below to execute it.  Please ensure`r`n"
    $sql += "-- that you've reviewed it before executing.  Specifically, ensure that you're`r`n"
    $sql += "-- not doing something like dropping a table that exists for some other purpose.`r`n"
    $sql += "`r`n"
    $sql += "-- Always test scripts & SQL in a TEST environment before using in production.`r`n"
    $sql += "`r`n"
    $sql += "-- The author(s) of this script cannot take responsibility for the script being used`r`n"
    $sql += "-- improperly or not having the intended effect in your environment.`r`n"
    $sql += "`r`n"
    $sql += "/*`r`n"

    if(!$dontDropTable) {
        # There are better ways to do this in SQL2016, but we can't know the version this will run against.
        $sql += "IF OBJECT_ID('$tableName', 'U') IS NOT NULL DROP TABLE $tableName;`r`n"
    }

    # Create table if it doesn't already exist from a previous run.
    $sql += "IF OBJECT_ID('$tableName', 'U') IS NULL CREATE TABLE $tableName (VersionNumber VARCHAR(32) NOT NULL, BuildNumber VARCHAR(8) NOT NULL, CONSTRAINT PK_O365Version PRIMARY KEY (VersionNumber));`r`n"

    If($deleteExistingRecords) {
        $sql += "DELETE FROM $tableName;"
    }

    ForEach($build in $officeBuildList) {
        $sql += "INSERT INTO $tableName (VersionNumber, BuildNumber) VALUES ('$($build.VersionNumber)','$($build.BuildNumber)');`r`n"
    }

    $sql += "*/"

    Return $sql
}


