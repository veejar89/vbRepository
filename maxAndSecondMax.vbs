arrLength = InputBox("Enter the length of Array")
dim arrInt()
redim arrInt(arrLength)
For a = 0 to arrLength-1
  arrInt(a) = InputBox("Enter the number at Index at " & a)
Next
Msgbox "Max and Second Max Numbers are " & maxAndSecondMaxInt(arrLength, arrInt)

Function maxAndSecondMaxInt(l, arr)
If (l >=2) Then
  If arr(0)> arr(1) Then
    big1 = arr(0)
    big2 = arr(1)
  Else
    big1 = arr(1)
    big2 = arr(0)
  End If
    for a = 2 to ubound(arr)-1
      If (arr(a) > big2) Then
        big2 = arr(a)
      End If
      If (arr(a)> big1) Then
        big2 = big1
        big1 = arr(a)
      End If
    next
  maxAndSecondMaxInt = big1 & " and " & big2
Else
  maxAndSecondMaxInt = "Array doesnt have second element"
End If
End Function
