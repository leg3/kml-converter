$inputPath = Read-Host -Prompt "Enter input filepath"
$outputPath = Read-Host -Prompt "Enter output filepath without filetype"
[xml]$kml = Get-Content $inputPath # this is the command we will use to parse the kml file which is actually xml

# Convert a date string into an object

function Convert-DateString ([String]$Date, [String[]]$Format)
{
   $result = New-Object DateTime
 
   $convertible = [DateTime]::TryParseExact(
      $Date,
      $Format,
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::None,
      [ref]$result)
 
   if ($convertible) { $result }
}

$kml.kml.Folder.Folder | ForEach-Object { 

    # Select the inner text and cast it into a variable
    $nodeValue = $_.InnerText -replace "<br>" , "" -replace "</br>" , "" -replace "<hr>" , "" -replace "</hr>" , "" -replace "<b>" , "" -replace "</b>" , "" -replace "  " , "" -replace " " , "_" -replace " " , "" -replace "First-seen:" , " " -replace "Last-seen:" , " " -replace "BSSID:" , " " -replace "Manufacturer:" , " " -replace "Channel:" , " " -replace "Frequency:" , " " -replace "Mhz" , "" -replace "Encryption:" , " " -replace "Min-Signal:" , " " -replace "dBm" , "" -replace "Max-Signal:" , " " -replace "GPS" , " " -replace "Avg lat/lon: " , " " -replace "Captured Packets" , " " -replace "LLC" , " " -replace "data" , " " -replace "crypt:" , " " -replace "total" , " " -replace "fragments" , " fragments" -replace "retries:" , " " -replace "node" , " " -replace "clampedToGround00" , " " -replace "`n" , "" -replace "`r" , ""

    # Split this string into a multi line string array using the "space" as a delimiter
    $nodeValueArray = @($nodeValue.Split(" "))

    # Mangle each object in the array to remove any remaining spaces and convert previously inserted underscores back to spaces
    $nodeValueArray = $nodeValueArray -replace " " , "" -replace "_" , " "

    # Load the array object into variables
    $firstSeenText = $nodeValueArray[1] | Select-Object
    $lastSeenText = $nodeValueArray[2] | Select-Object
    $bssid = $nodeValueArray[3] | Select-Object
    $manufacturer = $nodeValueArray[4] | Select-Object
    $channel = $nodeValueArray[5] | Select-Object
    $frequency = $nodeValueArray[6] | Select-Object
    $encryption = $nodeValueArray[7] | Select-Object
    $minimumSignal = $nodeValueArray[8] | Select-Object
    $maximumSignal = $nodeValueArray[9] | Select-Object

    # Convert the date strings into a date time object
    $firstSeen = Convert-DateString -Date $firstSeenText -Format 'ddd MMM dd HH:mm:ss yyyy'
    $lastSeen = Convert-DateString -Date $lastSeenText -Format 'ddd MMM dd HH:mm:ss yyyy'

    # Create a new object for SSID information that holds all of our values for this row. 
    $compositeSSID = New-Object Object
 
    # Add properties and values to the newly created  SSID object
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name NAME -Value $_.name
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name LATITUDE -Value $_.lookat.latitude
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name LONGITUDE -Value $_.lookat.longitude
    
    # Create a new object for the details associated with the SSID values for this row.
    $compositeDetails = New-Object Object 

    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name BSSID -Value $bssid
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MANUFACTURER -Value $manufacturer 
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name CHANNEL -Value $channel
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name FREQUENCY -Value $frequency
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name ENCRYPTION -Value $encryption
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MINIMUM_SIGNAL -Value $minimumSignal
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MAXIMUM_SIGNAL -Value $maximumSignal

    # Create a new object for the first-seen date values
    $compositeFirstdate = New-Object Object 

    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name FIRST_SEEN -Value $firstSeen.date
    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name DAY -Value $firstSeen.Day
    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name DAY_OF_WEEK -Value $firstSeen.DayOfWeek
    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name DAY_OF_YEAR -Value $firstSeen.DayOfYear
    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name HOUR -Value $firstSeen.Hour
    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name MINUTE -Value $firstSeen.Minute
    Add-Member -InputObject $compositeFirstDate -MemberType NoteProperty -Name SECOND -Value $firstSeen.Second

    # Create a new object for the last-seen date vales
    $compositeLastDate = New-Object Object

    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name LAST_SEEN -Value $lastSeen.date
    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name DAY -Value $lastSeen.Day
    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name DAY_OF_WEEK -Value $lastSeen.DayOfWeek
    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name DAY_OF_YEAR -Value $lastSeen.DayOfYear
    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name HOUR -Value $lastSeen.Hour
    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name MINUTE -Value $lastSeen.Minute
    Add-Member -InputObject $compositeLastDate -MemberType NoteProperty -Name SECOND -Value $lastSeen.Second

    # Export the $compositeSSID object to csv file and append
    $compositeSSID | Export-Csv -Path ( $outputPath + "_SSID.csv" ) -Append

    # Exort the $conpositeDetails object to a csv file and append
    $compositeDetails | Export-Csv -Path ( $outputPath + "_DETAILS.csv" ) -Append

    # Exort the $conpositeDetails object to a csv file and append
    $compositeFirstDate | Export-Csv -Path ( $outputPath + "_FIRST-DATE.csv" ) -Append 

    # Exort the $conpositeDetails object to a csv file and append
    $compositeLastDate | Export-Csv -Path ( $outputPath + "_LAST-DATE.csv" ) -Append 

}