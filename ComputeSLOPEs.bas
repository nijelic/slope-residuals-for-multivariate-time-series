Attribute VB_Name = "Module1"
' MIT License
'
' Copyright (c) 2023 Nikola Jelic
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.




'       *******************************************
'
'                 CORE FUNCTIONs & SUBs
'
'       *******************************************


'
' Line explicit equation as : y = c0 * x + c1
' Return: Coefficients of the line as array
'
Private Function CalculateLine(x1 As Double, x2 As Double, y1 As Double, y2 As Double) As Double()
    Dim lineCoeffs(1) As Double
    lineCoeffs(0) = (y2 - y1) / (x2 - x1)
    lineCoeffs(1) = y2 - lineData(0) * x2
    CalculateLine = lineCoeffs
End Function

'
' Calculates the distance between the point and the line.
' Return: The distance between the point and the line.
'
Private Function CalculateDistance(x1 As Double, y1 As Double, lineCoeffs() As Double) As Double
    CalculateDistance = Abs(lineCoeffs(0) * x1 + lineCoeffs(1) - y1) / Sqr(1 + lineCoeffs(0) ^ 2)
End Function

'
' Computes the distances between multiple points and the line.
' Return: The distances, for each point.
'
Private Function ComputeDistances(xArray() As Double, yArray() As Double, lineCoeffs() As Double) As Double()
    Dim distances() As Double
    Dim N As Long
    N = UBound(xArray) - LBound(xArray) + 1
    ReDim distances(N - 1)
    
    For i = LBound(xArray) To UBound(xArray)
        distances(i) = CalculateDistance(xArray(i), yArray(i), lineCoeffs)
    Next i
    
    ComputeDistances = distances
End Function

'
' Returns array elements as array from [lowerIndex, upperIndex].
' Return: Array of doubles, segmented.
'
Private Function GetSegment(inputArray() As Double, lowerIndex As Long, upperIndex As Long) As Double()
    Dim segmentedArray() As Double
    ReDim segmentedArray(upperIndex - lowerIndex + 1)
    
    Dim i As Long
    i = 0
    For j = lowerIndex To upperIndex
        segmentedArray(i) = inputArray(j)
        i = i + 1
    Next j
    
    GetSegment = segmentedArray
End Function

'
' Converts Range to Array of doubles.
' Return: Array of doubles.
'
Private Function ConvertRangeToArray(someRange As Range, commaUsed As Boolean) As Double()
    Dim N As Long
    N = someRange.Cells.Count
    
    Dim data() As Double
    ReDim data(N - 1)
    
    Dim i As Long
    i = 0
    
    ' Comment this line if needed
    commaUsed = False
    
    For Each c In someRange.Cells
        If commaUsed Then
            data(i) = Replace(Replace(c.Value, ".", ""), ",", ".")
        Else
            data(i) = c.Value
        End If
        i = i + 1
    Next
    ConvertRangeToArray = data
End Function

'
' Calculates the linear regression from data: x and y.
' Return: c0 and c1 as: y = c0 * x + c1.
'
Private Function LinearRegression(xArray() As Double, yArray() As Double) As Double()
    Dim N As Long
    Dim sumX As Double
    Dim sumY As Double
    Dim sumXY As Double
    Dim sumX2 As Double
    
    N = UBound(xArray) - LBound(xArray) + 1
    sumX = 0
    sumY = 0
    sumXY = 0
    sumX2 = 0
    For i = 0 To (N - 1)
        sumX = sumX + xArray(i)
        sumY = sumY + yArray(i)
        sumXY = sumXY + xArray(i) * yArray(i)
        sumX2 = sumX2 + xArray(i) * xArray(i)
    Next i
    
    Dim lineCoeffs(1) As Double
    lineCoeffs(0) = (N * sumXY - sumX * sumY) / (N * sumX2 - (sumX) ^ 2)
    lineCoeffs(1) = sumY / N - lineCoeffs(0) * (sumX / N)
    
    LinearRegression = lineCoeffs
End Function

'
' Computes linear regeression on segment [lowerBound, upperBound].
' Return: Array with coefficients c0 and c1, as: y = c0 * x + c1.
'
Private Function LinearRegressionBounded(xArray() As Double, yArray() As Double, ByVal lowerBound As Long, ByVal upperBound As Long) As Double()
    Dim N As Long
    N = UBound(xData) - LBound(xData) + 1
    
    Dim sumX As Double
    Dim sumY As Double
    Dim sumXY As Double
    Dim sumX2 As Double
    
    sumX = 0
    sumY = 0
    sumXY = 0
    sumX2 = 0
    
    If lowerBound < LBound(xArray) Then
        lowerBound = LBound(xArray)
    End If
        If upperBound > UBound(xArray) Then
        upperBound = UBound(xArray)
    End If
    For i = lowerBound To upperBound
        sumX = sumX + xArray(i)
        sumY = sumY + yArray(i)
        sumXY = sumXY + xArray(i) * yArray(i)
        sumX2 = sumX2 + xArray(i) * xArray(i)
    Next i
    
    Dim bigN As Long
    bigN = upperBound - lowerBound + 1
    
    Dim lineCoeffs(1) As Double
    lineCoeffs(0) = (bigN * sumXY - sumX * sumY) / (bigN * sumX2 - (sumX) ^ 2)
    
    lineCoeffs(1) = sumY / bigN - lineCoeffs(0) * (sumX / bigN)
    
    LinearRegressionBounded = lineCoeffs
End Function

'
' Computes linear regeression respect to selected points (as booleans array).
' Return: Array with coefficients c0 and c1, as: y = c0 * x + c1.
'
Private Function LinearRegressionBooleans(xArray() As Double, yArray() As Double, booleans() As Boolean) As Double()
    Dim N As Long
    N = UBound(xArray) - LBound(xArray) + 1
    
    Dim sumX As Double
    Dim sumY As Double
    Dim sumXY As Double
    Dim sumX2 As Double
    
    sumX = 0
    sumY = 0
    sumXY = 0
    sumX2 = 0
    
    Dim Count As Long
    Count = 0
    For i = 0 To N - 1
        If booleans(i) Then
            sumX = sumX + xArray(i)
            sumY = sumY + yArray(i)
            sumXY = sumXY + xArray(i) * yArray(i)
            sumX2 = sumX2 + xArray(i) * xArray(i)
            Count = Count + 1
        End If
    Next i
    
    Dim lineCoeffs(1) As Double
    lineCoeffs(0) = (Count * sumXY - sumX * sumY) / (Count * sumX2 - (sumX) ^ 2)
    
    lineCoeffs(1) = sumY / Count - lineCoeffs(0) * (sumX / Count)
    
    LinearRegressionBooleans = lineCoeffs
End Function

'
' Finds the threshold distance that contains defined percentage of inliers.
' Return: Distance that contains percentageOfInliers.
'
Private Function FindDistanceContainingPercentageOfInliers(distances() As Double, percentageOfInliers As Double) As Double
    Dim distancesCopy() As Double
    ReDim distancesCopy(UBound(distances) - LBound(distances) + 1)
    
    Dim i As Long
    i = 0
    
    For j = LBound(distances) To UBound(distances)
        distancesCopy(i) = distances(j)
        i = i + 1
    Next j
    
    QuickSort distancesCopy, LBound(distancesCopy), UBound(distancesCopy)
    Dim index As Integer
    index = CInt((UBound(distancesCopy) + 1) * percentageOfInliers) - 1
    
    FindDistanceContainingPercentageOfInliers = distancesCopy(index)
End Function

'
' Computes linear regeression iteratively.
' Return: Array with coefficients c0 and c1, as: y = c0 * x + c1.
'
Private Function LinearRegressionRANSACLike(xArray() As Double, yArray() As Double, ByVal percentageOfInliers As Double, ByVal iterations As Long) As Double()
    Dim N As Long
    N = UBound(xArray) - LBound(xArray) + 1
    Dim numberOfInliers As Long
    numberOfInliers = percentageOfInliers * N
    
    Dim booleans() As Boolean
    ReDim booleans(N - 1)
    
    ' First iteration will look at all elements
    For i = 0 To N - 1
        booleans(i) = True
    Next i
    
    Dim lineCoeffs() As Double
    Dim distances() As Double
    Dim thresholdDistance As Double
    
    For i = 0 To iterations - 1
        lineCoeffs = LinearRegressionBooleans(xArray, yArray, booleans)
        distances = ComputeDistances(xArray, yArray, lineCoeffs)
        thresholdDistance = FindDistanceContainingPercentageOfInliers(distances, percentageOfInliers)
        
        ' Pick elements that are good fit
        For j = 0 To N - 1
            If distances(j) < thresholdDistance Then
                booleans(j) = True
            Else
                booleans(j) = False
            End If
        Next j
    Next i
    
    LinearRegressionRANSACLike = lineCoeffs
End Function


'
' Subtracts one array from another.
' Return: difference as array.
'
Private Function SubtractArrays(minuendArray() As Double, subtrahendArray() As Double) As Double()
    Dim difference() As Double
    ReDim difference(UBound(minuendArray) - LBound(minuendArray))
    
    For i = LBound(minuendArray) To UBound(minuendArray)
        difference(i) = minuendArray(i) - subtrahendArray(i)
    Next i
    
    SubtractArrays = difference
End Function

'
' Calculates average distance of selected points from line.
' Return: average distance as Double.
'
Private Function CalculateAverageDistance(xArray() As Double, yArray() As Double, lineCoeffs() As Double, ByVal lowerBound As Long, ByVal upperBound As Long) As Double
    Dim averageDistance As Double
    averageDistance = 0
    
    For i = lowerBound To upperBound
        averageDistance = averageDistance + CalculateDistance(xArray(i), yArray(i), lineCoeffs)
    Next
    
    averageDistance = averageDistance / (upperBound - lowerBound + 1)
    CalculateAverageDistance = averageDistance
End Function

'
' Finds the point where the curvature turns the most.
' Return: deflection point as index.
'
Private Function FindDeflectionPointIndex(xArray() As Double, yArray() As Double) As Long
    Dim ignoreFirstSize As Long
    Dim minSetSize As Long
    Dim stepSize As Long
    Dim begin As Long
    Dim middle As Long
       
    ignoreFirstSize = 7
    minSetSize = 17
    stepSize = 3
    
    begin = ignoreFirstSize
    middle = minSetSize + ignoreFirstSize
    
    ' initialize minAverageDistance
    Dim minAverageDistance As Double
    Dim deflectionPointIndex As Long
    deflectionPointIndex = middle
    
    Dim N As Long
    N = UBound(xArray) - LBound(xArray) + 1

    minAverageDistance = CalculateAverageDistance(xArray, yArray, LinearRegressionBounded(xArray, yArray, begin, middle), begin, middle)
    minAverageDistance = minAverageDistance + CalculateAverageDistance(xArray, yArray, LinearRegressionBounded(xArray, yArray, middle, N - 1), middle, N - 1)
    
    Dim average_Distance As Double
    Do While middle < (N - 1) - minSetSize
        average_Distance = CalculateAverageDistance(xArray, yArray, LinearRegressionBounded(xArray, yArray, begin, middle), begin, middle)
        average_Distance = average_Distance + CalculateAverageDistance(xArray, yArray, LinearRegressionBounded(xArray, yArray, middle, N - 1), middle, N - 1)
        If average_Distance < minAverageDistance Then
            minAverageDistance = average_Distance
            deflectionPointIndex = middle
            
        End If
        middle = middle + stepSize
    Loop
    
    FindDeflectionPointIndex = deflectionPointIndex
End Function

'
' Computes the residuals.
' Return: Array of residuals as double.
'
Private Function ComputeResiduals(xArray() As Double, yArray() As Double, lineCoeffs() As Double) As Double()
    Dim residualsArray() As Double
    ReDim residualsArray(UBound(xArray) - LBound(xArray) + 1)
    
    Dim i As Long
    i = 0
    
    For j = LBound(xArray) To UBound(xArray)
        residualsArray(i) = yArray(j) - (lineCoeffs(0) * xArray(j) + lineCoeffs(1))
        i = i + 1
    Next j
    
    ComputeResiduals = residualsArray
End Function


'
' Quick Sort used for sorting.
'
Private Sub QuickSort(vArray() As Double, inLow As Long, inHi As Long)
  Dim pivot   As Double
  Dim tmpSwap As Double
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
End Sub


'       *******************************************
'
'                       Main MACROs
'
'       *******************************************


Public Sub ComputeSLOPEs()
    Dim convertCommaToDot As Boolean
    convertCommaToDot = False

    colsCount = Selection.Columns.Count
    rowsCount = Selection.Rows.Count
    
    Dim timeRange As Range, backgroundRange As Range
    Set timeRange = Selection.Columns(1).Rows
    Set backgroundRange = Selection.Columns(colsCount).Rows

    Debug.Print "rows = " & rowsCount
    Dim timeArray() As Double, backgroundArray() As Double
    ReDim timeArray(rowsCount - 1)
    ReDim backgroundArray(rowsCount - 1)
    
    timeArray = ConvertRangeToArray(timeRange, convertCommaToDot)
    backgroundArray = ConvertRangeToArray(backgroundRange, convertCommaToDot)
    
    Dim valuesRange As Range
    Dim valuesArray() As Double
    ReDim valuesArray(rowsCount - 1)
    
    Dim deflectionPointIndex As Long
    
    Dim beginIndex As Long
    Dim middleIndex As Long
    Dim endIndex As Long
    Dim deviateFromMiddle As Long
    
    '  ****************************************************************************************************
    '
    '                                             MAIN PARAMETERS
    '
    '  beginIndex        The first index after which your data should be consistent.
    '
    '  middleIndex       The index to separate the data so that the first part is used for the first slope,
    '      and the data after it is used for the second slope.
    '
    '  endIndex          The last index of consistent data.
    '
    '  deviateFromMiddle The size of deviation from middleIndex.
    '      So first slope will be calculated on: [beginIndex, middleIndex - deviateFromMiddle].
    '      Second slope will be calculate on the [middleIndex + deviateFromMiddle, endIndex].
    '
    '  ****************************************************************************************************
    beginIndex = 5
    middleIndex = 19
    endIndex = 53
    deviateFromMiddle = 7

    ' Used for RANSAC Like Linear Regression
    Dim percentageOfInliers As Double, iterations As Long
    percentageOfInliers = 0.9
    iterations = 20
    
    Dim lineCoeffs() As Double
    Dim selectedValues() As Double
    Dim selectedTimes() As Double
    
    ' i = 2 because in the first column are time stamps
    Dim i As Long
    i = 2
    Do While i < colsCount

        Set valuesRange = Selection.Columns(i).Rows
              
        valuesArray = SubtractArrays(ConvertRangeToArray(valuesRange, convertCommaToDot), backgroundArray)
                
        ' Compute first line
        selectedValues = GetSegment(valuesArray, beginIndex - 1, middleIndex - deviateFromMiddle)
        selectedTimes = GetSegment(timeArray, beginIndex - 1, middleIndex - deviateFromMiddle)
        
        lineCoeffs = LinearRegressionRANSACLike(selectedTimes, selectedValues, percentageOfInliers, iterations)
        
        Cells(rowsCount + 2, valuesRange.Column - 1).Value = "Slope"
        Cells(rowsCount + 2, valuesRange.Column).Value = lineCoeffs(0)
        Cells(rowsCount + 3, valuesRange.Column).Value = lineCoeffs(1)
        
        ' Compute second slope
        selectedValues = GetSegment(valuesArray, middleIndex + deviateFromMiddle, endIndex - 1)
        selectedTimes = GetSegment(timeArray, middleIndex + deviateFromMiddle, endIndex - 1)
        
        selectedValues = ComputeResiduals(selectedTimes, selectedValues, lineCoeffs)
        lineCoeffs = LinearRegressionRANSACLike(selectedTimes, selectedValues, percentageOfInliers, iterations)
        
        Cells(rowsCount + 5, valuesRange.Column).Value = lineCoeffs(0)
        Cells(rowsCount + 6, valuesRange.Column).Value = lineCoeffs(1)
        Cells(rowsCount + 8, valuesRange.Column).Value = Abs(lineCoeffs(0))
        Cells(rowsCount + 10 + i / 2, 25).Value = Abs(lineCoeffs(0))
        
        i = i + 2
    Loop
End Sub






