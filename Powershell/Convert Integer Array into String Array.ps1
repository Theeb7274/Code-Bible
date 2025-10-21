# A function to convert an array of integers into strings, useful when defining a large range of numbers using (10..1000) as this defaults to integers
# This then allows you to append ASCII characters to all entries in the array, such as a wildcard

function ArrayConversion{
    $IntArray= (..)


    $StringArray = $IntArray | ForEach-Object {
        $_.ToString()  
    }

	# Optionally append a character to each entry in the array
    $AmmendedArray = $IntArray | ForEach-Object {
        $_ + "*" # $_ represents the current item in the pipeline
    }
}



