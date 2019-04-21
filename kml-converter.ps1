# Convert a date string into an object
function convertDateToString ([string]$Date,[String[]]$Format)
{
  $result = New-Object DateTime

  $itIsConvertible = [datetime]::TryParseExact(
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

    # Select the inner text and cast it into a variable
    $nodeValue = $_.InnerText -replace "<br>","" -replace "</br>","" -replace "<hr>","" -replace "</hr>","" -replace "<b>","" -replace "</b>","" -replace "  ","" -replace " ","__" -replace " ","" -replace "First-seen:"," " -replace "Last-seen:"," " -replace "BSSID:"," " -replace "Manufacturer:"," " -replace "Channel:"," " -replace "Frequency:"," " -replace "Mhz","" -replace "Encryption:"," " -replace "Min-Signal:"," " -replace "dBm","" -replace "Max-Signal:"," " -replace "GPS"," " -replace "Avg lat/lon: "," " -replace "Captured Packets"," " -replace "LLC:"," " -replace "data:"," " -replace "crypt:"," " -replace "total:"," " -replace "fragments:"," fragments:" -replace "retries:"," " -replace "node_wpa.png"," " -replace "clampedToGround00"," " -replace "`n","" -replace "`r",""

    # Split this string into a multi line string array using the "space" as a delimiter
    $nodeValueArray = @($nodeValue.Split(" "))

    # Mangle each object in the array to remove any remaining spaces and convert previously inserted underscores back to spaces
    $nodeValueArray = $nodeValueArray -replace " ","" -replace "__"," "

    # Load the array object into variables
    $firstSeenText = $nodeValueArray[1].trim()
    $lastSeenText = $nodeValueArray[2].trim()
    $bssid = $nodeValueArray[3].trim()
    $manufacturer = $nodeValueArray[4].trim()
    $channel = $nodeValueArray[5].trim()
    $frequency = $nodeValueArray[6].trim()
    $encryption = $nodeValueArray[7].trim()
    $minimumSignal = $nodeValueArray[8].trim()
    $maximumSignal = $nodeValueArray[9].trim()

    # Convert the date strings into a date time object
    $firstSeen = convertDateToString -Date $firstSeenText -Format 'ddd MMM dd HH:mm:ss yyyy'
    $lastSeen = convertDateToString -Date $lastSeenText -Format 'ddd MMM dd HH:mm:ss yyyy'

    # Value pre-processing
    $frequency = $frequency.trimEnd("1","2","3","4","5","6","7","8","9","0")

    # Create a new object for SSID information that holds all of our values for this row. 
    $compositeSSID = New-Object Object

    # Add properties and values to the newly created  SSID object
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name NAME -Value $_.Name
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name LATITUDE -Value $_.lookat.latitude
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name LONGITUDE -Value $_.lookat.longitude

    # Create a new object for the details associated with the SSID values for this row.
    $compositeDetails = New-Object Object

    #Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name DEBUG -Value $debug
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name BSSID -Value $bssid
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MANUFACTURER -Value $manufacturer
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name CHANNEL -Value $channel
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name FREQUENCY -Value $frequency.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name ENCRYPTION $encryption.trim()
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MINIMUM_SIGNAL -Value $minimumSignal
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name MAXIMUM_SIGNAL -Value $maximumSignal

    # Create a new object for the first-seen date values
    $compositeFirstSeen = New-Object Object

    if ($NULL -eq $firstSeen) { $firstSeen = Get-Date }

    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name FIRST_SEEN -Value $firstSeen.date.ToString("yyy-MM-dd")
    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY -Value $firstSeen.Day
    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY_OF_WEEK -Value $firstSeen.DayOfWeek
    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name DAY_OF_YEAR -Value $firstSeen.DayOfYear
    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name HOUR -Value $firstSeen.Hour
    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name MINUTE -Value $firstSeen.Minute
    Add-Member -InputObject $compositeFirstSeen -MemberType NoteProperty -Name SECOND -Value $firstSeen.Second

    # Create a new object for the last-seen date values
    $compositeLastSeen = New-Object Object

    if ($NULL -eq $lastSeen) { $lastSeen = Get-Date }

    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name LAST_SEEN -Value $lastSeen.date.ToString("yyy-MM-dd")
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name DAY -Value $lastSeen.Day
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name DAY_OF_WEEK -Value $lastSeen.DayOfWeek
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name DAY_OF_YEAR -Value $lastSeen.DayOfYear
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name HOUR -Value $lastSeen.Hour
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name MINUTE -Value $lastSeen.Minute
    Add-Member -InputObject $compositeLastSeen -MemberType NoteProperty -Name SECOND -Value $lastSeen.Second

    # Export the $compositeSSID object to csv file and append
    $compositeSSID | Export-Csv -Path ($outputPath.Replace(".kml","") + "_SSID.csv") -Append -NoTypeInformation

    # Export the $conpositeDetails object to a csv file and append
    $compositeDetails | Export-Csv -Path ($outputPath.Replace(".kml","") + "_DETAILS.csv") -Append -NoTypeInformation

    # Export the $conpositeDetails object to a csv file and append
    $compositeFirstSeen | Export-Csv -Path ($outputPath.Replace(".kml","") + "_FIRST_SEEN.csv") -Append -NoTypeInformation

    # Export the $conpositeDetails object to a csv file and append
    $compositeLastSeen | Export-Csv -Path ($outputPath.Replace(".kml","") + "_LAST_SEEN.csv") -Append -NoTypeInformation

  }

}

# Move the newly generated files into folders that are named after the individual .kml files
Get-ChildItem -Path ${psscriptroot} -Name -Filter "*.kml" | ForEach-Object {

  $name = $_.Replace(".kml","")
  mkdir $name
  $destination = Get-ChildItem -Path ./ -Directory $name
  Write-Host ("Processing " + $_ + "`n")
  sqlPreParse ($_);
  Write-Host ($_ + " Processed...`n")

  Get-ChildItem -Path ${psscriptroot} -Name -Filter "*.csv" | ForEach-Object {

    Move-Item -Path .\*.csv -Destination $destination

  }

}

# Log the count from our .csv files for reference
Get-ChildItem -Path ${psscriptroot} -Recurse -Name -Filter "*.csv" | ForEach-Object {

  $_ | Out-File log.txt -Append
  Import-Csv $_ | Measure-Object | Out-File log.txt -Append

}

# Concatenate the "DETAILS" .csv files 
$getFirstLine = $true
Get-ChildItem -Path ${psscriptroot} -Recurse -Name -Filter "*DETAILS.csv" | ForEach-Object {
  $filePath = $_
  $lines = Get-Content $filePath
  $linesToWrite = switch ($getFirstLine) {
    $true { $lines }
    $false { $lines | Select-Object -Skip 1 }
  }
  $getFirstLine = $false
  Add-Content "${psscriptroot}\COMPOSITE_DETAILS.csv" $linesToWrite
}

# Concatenate the "FIRST_SEEN" .csv files 
$getFirstLine = $true
Get-ChildItem -Path ${psscriptroot} -Recurse -Name -Filter "*FIRST_SEEN.csv" | ForEach-Object {
  $filePath = $_
  $lines = Get-Content $filePath
  $linesToWrite = switch ($getFirstLine) {
    $true { $lines }
    $false { $lines | Select-Object -Skip 1 }
  }
  $getFirstLine = $false
  Add-Content "${psscriptroot}\COMPOSITE_FIRST_SEEN.csv" $linesToWrite
}

# Concatenate the "LAST_SEEN" .csv files
$getFirstLine = $true
Get-ChildItem -Path ${psscriptroot} -Recurse -Name -Filter "*LAST_SEEN.csv" | ForEach-Object {
  $filePath = $_
  $lines = Get-Content $filePath
  $linesToWrite = switch ($getFirstLine) {
    $true { $lines }
    $false { $lines | Select-Object -Skip 1 }
  }
  $getFirstLine = $false
  Add-Content "${psscriptroot}\COMPOSITE_LAST_SEEN.csv" $linesToWrite
}

# Concatenate the "SSID" .csv files
$getFirstLine = $true
Get-ChildItem -Path ${psscriptroot} -Recurse -Name -Filter "*SSID.csv" | ForEach-Object {
  $filePath = $_
  $lines = Get-Content $filePath
  $linesToWrite = switch ($getFirstLine) {
    $true { $lines }
    $false { $lines | Select-Object -Skip 1 }
  }
  $getFirstLine = $false
  Add-Content "${psscriptroot}\COMPOSITE_SSID.csv" $linesToWrite
}
