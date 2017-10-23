# gridDist 

VBA function for use in Excel which uses Pythagoras' Theorem to calculate the straight line distance between two pairs of Eastings and Northings on the OS National Grid.

Assumes 12 figure reference (6 figure Easting and 6 figure Northing) i.e. accuracy of 1m. Assumes first digit is national grid line reference, not two-letter reference.

## Usage
`gridDist(east1, north1, east2, north2)`
