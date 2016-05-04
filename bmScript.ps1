$filePath = "K:\MIS\Ola\cis_win7.csv"
<# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false
#>
#Import the data from csv file and store in a variable
$rawData = Import-Csv -path $filePath

#select all hive objects
$cisData = ($rawData | Where-Object{$_.hive, $_.key, $_.name , $_.type, $_.value})

#create objects of each column entry in $cisData
$hive = ($cisData | ForEach-Object{$_.hive})
$key = ($cisData | ForEach-Object{$_.key})
$name = ($cisData | ForEach-Object{$_.name})
$value = ($cisData | ForEach-Object{$_.value})
$fullPath = ($cisData | ForEach-Object{"Registry::" +  $_.hive + "\" + $_.key})
$count = $fullPath.count

cls #Clear the screen
#check if the registry path exists
$propertyFound = for($index = 0 ; $index -lt $count; $index++)
{
    if ((Test-path $fullPath[$index]) -and $key[$index] -ne "")
    {
		function getValue($value){
		
        Try
        {
            Get-ItemProperty -Path $fullPath[$index] | Select -ExpandProperty $name[$index] -EA stop
        }
        
        Catch [System.Exception $ex]
        {
            $ex | Out-File "C:\Users\olao\Desktop\exception.txt"
                      
        }
        
        Finally
        {
            $error.Clear()
        }             
    }
}

$propertyFound