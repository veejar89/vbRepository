'You are given an integer N. You need to convert all zeroes of N to 5.
'Example 1:
'Input:
'N = 1004
'Output: 1554
'Explanation: There are two zeroes in 1004
'on replacing all zeroes with "5", the new
'number will be "1554".

numb = InputBox("Enter total numbers")
for a = 1 to numb
  n = Inputbox("Enter the number")
  msgbox convertFive(n)
next

Function convertFive(n)
  givenNumber = n &""
  givenNumber = Replace(givenNumber, "0","5")
  convertFive = CInt(givenNumber)
End Function
