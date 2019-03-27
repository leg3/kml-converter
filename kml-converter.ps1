$inputPath = Read-Host -Prompt "Enter input filepath"
$outputPath = Read-Host -Prompt "Enter output filepath without filetype"
[xml]$kml = Get-Content $inputPath # this is the command we will use to parse the kml file which is actually xml

$kml.kml.Folder.Folder | ForEach-Object { 

    # Select the inner text and cast it into a variable
    $nodeValue = $_.InnerText -replace "<br>" , "" -replace "</br>" , "" -replace "<hr>" , "" -replace "</hr>" , "" -replace "<b>" , "" -replace "</b>" , "" -replace "  " , "" -replace " " , "_" -replace " " , "" -replace "First-seen:" , " " -replace "Last-seen:" , " " -replace "BSSID:" , " " -replace "Manufacturer:" , " " -replace "Channel:" , " " -replace "Frequency:" , " " -replace "Mhz" , "" -replace "Encryption:" , " " -replace "Min-Signal:" , " " -replace "dBm" , "" -replace "Max-Signal:" , " " -replace "GPS" , " " -replace "Avg lat/lon: " , " " -replace "Captured Packets" , " " -replace "LLC" , " " -replace "data" , " " -replace "crypt:" , " " -replace "total" , " " -replace "fragments" , " fragments" -replace "retries:" , " " -replace "node" , " " -replace "clampedToGround00" , " " -replace "`n" , ""

    # Split this string into a multi line string array using the "space" as a delimiter
    $nodeValueArray = @($nodeValue.Split(" "))

    # Mangle each object in the array to remove any remaining spaces and convert previously inserted underscores back to spaces
    $nodeValueArray = $nodeValueArray -replace " " , "" -replace "_" , " "

    # Load the array object into variables
    $firstSeen = $nodeValueArray[1] | Select-Object 
    $lastSeen = $nodeValueArray[2] | Select-Object
    $bssid = $nodeValueArray[3] | Select-Object
    $manufacturer = $nodeValueArray[4] | Select-Object
    $channel = $nodeValueArray[5] | Select-Object
    $frequency = $nodeValueArray[6] | Select-Object
    $encryption = $nodeValueArray[7] | Select-Object
    $minimumSignal = $nodeValueArray[8] | Select-Object
    $maximumSignal = $nodeValueArray[9] | Select-Object

    # Create a new object for SSID information that holds all of our values for this row. 
    $compositeSSID = New-Object Object
 
    # Add properties and values to the newly created  SSID object
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name BSSID -Value $bssid
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name Name -Value $_.name
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name Latitude -Value $_.lookat.latitude
    Add-Member -InputObject $compositeSSID -MemberType NoteProperty -Name Longitude -Value $_.lookat.longitude
    
    # Create a new object for the details associated with the SSID values for this row.
    $compositeDetails = New-Object Object 

    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name BSSID -Value $bssid
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name First-seen -Value $firstSeen
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Last-seen -Value $lastSeen
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Manufacturer -Value $manufacturer 
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Channel -Value $channel
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Frequency -Value $frequency
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Encryption -Value $encryption
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Minimum-Signal -Value $minimumSignal
    Add-Member -InputObject $compositeDetails -MemberType NoteProperty -Name Maximum-signal -Value $maximumSignal
    
    # Export the $compositeSSID object to csv file and append
    $compositeSSID | Export-Csv -Path ( $outputPath + "_SSID.csv" ) -Append

    # Exort the $conpositeDetails object to a csv file and append
    $compositeDetails | Export-Csv -Path ( $outputPath + "_DETAILS.csv" ) -Append 

}