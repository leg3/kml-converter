# Convert a date string into an object
function convertDateToString ([String]$Date, [String[]]$Format)
{
   $result = New-Object DateTime
 
   $itIsConvertible = [DateTime]::TryParseExact(
      $Date,
      $Format,
      [System.Globalization.CultureInfo]::InvariantCulture,
      [System.Globalization.DateTimeStyles]::AdjustToUniversal,
      [ref]$result)
 
   if ($itIsConvertible) { $result } else { 
      
      [System.Media.SystemSounds]::Hand.Play() 
      #get-date -Format 'ddd MMM dd HH:mm:ss yyyy'
   }
}

# Parse kml files and convert to csv
function sqlPreParse ($file) {

$private:inputPath = $file
$private:outputPath = $file

[xml]$kml = Get-Content $inputPath # this is the command we will use to parse the kml file which is actually xml

$kml.kml.Folder.Folder | ForEach-Object { 

    # Select the inner text and cast it into a varitable
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
    $firstSeen = convertDateToString -Date $firstSeenText -Format 'ddd MMM dd HH:mm:ss yyyy'
    $lastSeen = convertDateToString -Date $lastSeenText -Format 'ddd MMM dd HH:mm:ss yyyy'

    # Value pre-processing
    $frequency = $frequency.trimEnd( "1","2","3","4","5","6","7","8","9","0" )

    # Create a new object for SSID information that holds all of our values for this row. 
    $compositeSSID = New-Object Object
 
    # Add properties and values to the newly created  SSID object
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name NAME -Value $_.name
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name LATITUDE -Value $_.lookat.latitude
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name LONGITUDE -Value $_.lookat.longitude
    
    # Create a new object for the details associated with the SSID values for this row.
    $compositeDetails = New-Object Object 

    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name BSSID -Value $bssid.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MANUFACTURER -Value $manufacturer.trim();
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name CHANNEL -Value $channel.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name FREQUENCY -Value $frequency.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name ENCRYPTION -Value $encryption.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MINIMUM_SIGNAL -Value $minimumSignal.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MAXIMUM_SIGNAL -Value $maximumSignal.trim()

    # Create a new object for the first-seen date values
    $compositeFirstSeen = New-Object Object 

   if ($NULL -eq $firstSeen ) {
      
      $firstSeen = Get-Date
      
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name FIRST_SEEN -Value $firstSeen.date.tostring("yyy-MM-dd")
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY -Value $firstSeen.Day
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY_OF_WEEK -Value $firstSeen.DayOfWeek
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY_OF_YEAR -Value $firstSeen.DayOfYear
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name HOUR -Value $firstSeen.Hour
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name MINUTE -Value $firstSeen.Minute
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name SECOND -Value $firstSeen.Second
   
   } 
   
   else 
   
   { 
      
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name FIRST_SEEN -Value $firstSeen.date.tostring("yyy-MM-dd")
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY -Value $firstSeen.Day
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY_OF_WEEK -Value $firstSeen.DayOfWeek
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY_OF_YEAR -Value $firstSeen.DayOfYear
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name HOUR -Value $firstSeen.Hour
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name MINUTE -Value $firstSeen.Minute
      Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name SECOND -Value $firstSeen.Second

   }

    # Create a new object for the last-seen date vales
    $compositeLastSeen = New-Object Object

    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name LAST_SEEN -Value $lastSeen.date.tostring("yyy-MM-dd")
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name DAY -Value $lastSeen.Day
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name DAY_OF_WEEK -Value $lastSeen.DayOfWeek
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name DAY_OF_YEAR -Value $lastSeen.DayOfYear
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name HOUR -Value $lastSeen.Hour
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name MINUTE -Value $lastSeen.Minute
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name SECOND -Value $lastSeen.Second

    # Export the $compositeSSID object to csv file and append
    $compositeSSID | Export-Csv -Path ( $outputPath.replace(".kml","") + "_SSID.csv" ) -Append -NoTypeInformation 

    # Export the $conpositeDetails object to a csv file and append
    $compositeDetails | Export-Csv -Path ( $outputPath.replace(".kml","") + "_DETAILS.csv" ) -Append -NoTypeInformation 

    # Export the $conpositeDetails object to a csv file and append
    $compositeFirstSeen | Export-Csv -Path ( $outputPath.replace(".kml","") + "_FIRST_SEEN.csv" ) -Append -NoTypeInformation 

    # Export the $conpositeDetails object to a csv file and append
    $compositeLastSeen | Export-Csv -Path ( $outputPath.replace(".kml","") + "_LAST_SEEN.csv" ) -Append -NoTypeInformation 

}

}


Get-ChildItem -Path ${psscriptroot} -Name -Filter "*.kml" | ForEach-Object { 
   
   $name = $_.replace(".kml", "")
   mkdir $name
   $destination = Get-ChildItem -Path ./ -Directory $name
   Write-Host ( "Processing " + $_ + "`n" )
   sqlPreParse ($_); 
   Write-Host ( $_ + " Processed...`n" )

   Get-ChildItem -Path ${psscriptroot} -Name -Filter "*.csv" | ForEach-Object {
   
      Move-Item -path .\*.csv -Destination  $destination
      
      }

}

Get-ChildItem -Path ${psscriptroot} -Recurse -Name -Filter "*.csv" | ForEach-Object {

$_ | Out-File log.txt -Append  
Import-Csv $_ | Measure-Object | Out-File log.txt -Append

}
[System.Media.SystemSounds]::Beep.Play()
Read-Host -Prompt "Enter to exit"