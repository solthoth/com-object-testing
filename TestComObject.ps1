$cashObj = New-Object -ComObject TestComApplication.TestComApp # name as seen in "Registered Components"
$parameter = "VERSION                           "
# variable must be wrapped in a Variant when input parameter type is a Variant
$wrappedParameter = New-Object System.Runtime.InteropServices.VariantWrapper($parameter)
# Passing variable by reference since input is also the output
$cashObj.getCashWebData([ref] $wrappedParameter)
$wrappedParameter