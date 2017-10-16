'calculates the arithmetic mean excluding outliers
function mean_outlier(nums as range) as double

    'Outliers in this case will be determined by the [outer fence](https://www.wikihow.com/Calculate-Outliers)

    'Declare the variables
    Dim elem as variant
    dim counter as integer
    dim sum as double
    Dim first_quartile As Double
    Dim third_quartile As Double
    Dim iqr As Double
    dim lb as Double
    dim up as double
    dim outerfence as double

    'assign values to the variables

    'calculate the first quartile using the worksheetfunction quartile
    first_quartile = WorksheetFunction.Quartile(nums, 1)
    'calculate the third quartile using the worksheetfunction quartile
    third_quartile = WorksheetFunction.Quartile(nums, 3)
    'calculate the inter-quartile range
    iqr = third_quartile - first_quartile
    'calculate the outerfence which is 3x the iqr
    outerfence = iqr * 3.0

    'calculate boundaries
    lb = first_quartile - outerfence
    ub = third_quartile + outerfence

    'conditionally calculate the mean(average)

    'initialize the sum and counter to zero
    sum = 0
    counter = 0
    ' run the loop for each element in the range
    for each elem in nums
        if elem.value > lb and elem.value < ub then
            sum = sum + elem.value
            counter = counter + 1 'keep track of the number of elements
        end if

    next elem

    mean_outlier = sum / counter
    
End function
