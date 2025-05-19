function ArrayConversion{
    $IntArray= (..)


    $StringArray = $IntArray | ForEach-Object {
        $_.ToString()  
    }

	#Optionally append a character to each entry in the array
    #$newYears = $convertedYears | ForEach-Object {
        $_ + "*" # $_ represents the current item in the pipeline
    }
}
