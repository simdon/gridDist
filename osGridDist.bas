Function gridDist(east1 As Long, north1 As Long, east2 As Long, north2 As Long)
' SD 2016
' Uses Pythagoras' Theorem to calculate the straight line distance between two
' pairs of Eastings and Northings on the OS National Grid.
' Assumes 12 figure reference (6 figure Easting and 6 figure Northing) i.e. accuracy
' of 1m. Assumes first digit is national grid line reference, not two-letter reference.

'Dimension required variables
Dim dEast As Long, dNorth As Long, dLine As Double

'Subtract the eastings/northings of each location to get east/north components
' of vector.
dEast = east2 - east1
dNorth = north2 - north1

'Use Pythagoras' Theorem to calculate distance (a^2)=(b^2)+(c^2) therefore
' a=sqrt((b^2)+(c^2)) 
' Divide the result by 1000 to convert from metres to kilometres.
dLine = Sqrt((dEast ^ 2) + (dNorth ^ 2)) / 1000

'Pass the calculated distance to the output.
gridDist = dLine

End Function
