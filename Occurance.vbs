str = "asdsgfsd"
count = 0
For i = 1 To len(str) Step 1
	For j = 1 To len(str) Step 1
		a = mid (str,i,1)
		b = mid (str,j,len(str))
		c = left (b,1)

		If strcomp(a,c) = 0 Then
		count = count+1
		End If
	Next
	If a <> " " Or c <> " " Then
	msgbox a & " occurred "& count &" time."
	End If
	
	str = replace (str,a," ")
	'msgbox str
	count = 0
Next

