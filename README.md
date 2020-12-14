# timecode
Excel macro pour computing Timecode
' Module TimeCode
' DLEP - 202012
' Version 1.0 du 13/12/2020

Function FromTimeCode(nbimg As Variant, tc As Variant)
nbimg = nbimg by sec
tc = timecode as hh:mm:ss:ii

Function ToTimeCode(nbimg As Variant, img As Variant) As Variant
nbimg = nbimg by sec
img = number of images

Function TCDuree(nbimg As Variant, deb As Variant, fin As Variant)
nbimg = nbimg by sec
deb = timecode begin as hh:mm:ss:ii
fin = timecode end as hh:mm:ss:ii
