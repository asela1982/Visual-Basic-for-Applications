function mean(nums as range) as double

    dim elem as variant
    dim sum as double
    dim counter as integer
    
    sum = 0
    counter = 0

    for each elem in nums
        sum = sum + elem
        counter = counter + 1
    next elem

    mean = sum / counter

end function