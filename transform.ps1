# INput Fields
#  Site,Control System,Unit,Application Type,Application Instance,Points

# Output fields
# Store,Catagory,Sub Catagory,Sub Catagory 2,Sub Catagory 3,Product Name,Name,Serial,Model,Size,Date of Manufacture,Fuel Type,VFD,Economizer,Power Exhaust,Heating Stages,Cooling Stages,Reheat/Dehumidification,Notes,Heatpump

$infile = ".\Connect_Prod_02162024.txt"
$outfile = ".\dataout.csv"


$datain = Import-Csv -Delimiter "`t" -Header "Site","Control System","Unit","Application Type","Application Instance","Points" -Path $infile | where -Property "Application Type" -Match "lighting|ARTC|Circuits|AHU|RCB" #| where -Property "Site" -like "10649*"

$dataout = @()

$datain | ForEach-Object {
    $store = $_.Site.Substring(0,5)
    $Catagory = "Assets"
    $SubCatagory = ""
    $SubCatagory2 = "Generic"
    $SubCatagory3 = ""
    $ProductName = "Generic"
    $Name = $_.{Application Instance}
    $Serial = ""
    $Model = ""
    $Size = ""
    $DateOfManufacture = ""
    $FuelType = ""
    $VFD = ""
    $Economizer = ""
    $PowerExhaust = ""
    $HeatingStages = ""
    $CoolingStages = ""
    $ReheatDehumidification = ""
    $Notes = ""
    $Heatpump = ""

    $CoolingStages = ($_.Points  | select-string -AllMatches -pattern "COOL").Matches.length
    $HeatingStages = ($_.Points  | select-string  -AllMatches -pattern "HEAT").Matches.length
    
    if (($_.Points  | select-string  -AllMatches -pattern "(MOD ECON)|(ANALOG DAMPER)|(OA DAMPER OUT)").Matches.length -gt 0) {
        $Economizer = "Yes"
    } else {
        $Economizer = "No"
    }

    

    if  ($_.{Application Type} -like "lighting*") {
        $SubCatagory = "Lighting"
        $SubCatagory2 = "Generic"
        $SubCatagory3 = ""
        $ProductName = "LRP"
    }

    if  ($_.{Application Type} -like "ARTC*") {
        $SubCatagory = "HVAC"
        $SubCatagory2 = "Generic"
        $SubCatagory3 = "Package Unit"
    }

    if ($_.{Application Type} -like "Circuits") {
        $SubCatagory = "Refrigeration"
        $SubCatagory2 = "Generic"
        $SubCatagory3 = "Unknown"
    }
    
    if ($_.{Application Type} -like "AHU") {
        $SubCatagory = "HVAC"
        $SubCatagory2 = "Generic"
        $SubCatagory3 = "Air Handler"
    }
    
    if ($_.{Application Type} -like "*RCB*") {
        $SubCatagory = "HVAC"
        $SubCatagory2 = "Generic"
        $SubCatagory3 = "Package Unit"
    }
   
    if ($_.{Application Type} -like "") {
        $SubCatagory = ""
        $SubCatagory2 = ""
        $SubCatagory3 = ""
    }


   $dataout +=  [PSCustomObject]@{
        Store = $store
        Catagory = $Catagory
        {Sub Catagory} = $SubCatagory
        {Sub Catagory 2} = $SubCatagory2
        {Sub Catagory 3} = $SubCatagory3
        {Product Name} = $ProductName
        Name = $Name
        Serial = $Serial
        Model = $Model
        Size = $Size
        {Date of Manufacture} = $DateOfManufacture
        {Fuel Type} = $FuelType
        VFD = $VFD
        Economizer = $Economizer
        {Power Exhaust} = $PowerExhaust
        {Heating Stages} = $HeatingStages
        {Cooling Stages} = $CoolingStages
        {Reheat / Dehumidification} = $ReheatDehumidification
        Notes = $Notes
        Heatpump = $Heatpump
   }
   
}

$datain | select -Unique -Property Unit, Site  -ExpandProperty Unit| ForEach-Object {
    try {
        $name =  $_.Unit.Split(":")[0].Trim()
    } catch {
        $name =  $Unit.Trim()
    }
    
    $store = $_.Site.Substring(0,5)
      
    $dataout +=  [PSCustomObject]@{
        Store = $store
        Catagory = ""
        {Sub Catagory} = "EMS"
        {Sub Catagory 2} = "Generic"
        {Sub Catagory 3} = ""
        {Product Name} = ""
        Name = $name
        Serial = ""
        Model = ""
        Size = ""
        {Date of Manufacture} = ""
        {Fuel Type} = ""
        VFD = ""
        Economizer = ""
        {Power Exhaust} = ""
        {Heating Stages} = ""
        {Cooling Stages} = ""
        {Reheat / Dehumidification} = ""
        Notes = ""
        Heatpump = ""
    }
}


$dataout | Export-Csv -Path $outfile -NoTypeInformation

