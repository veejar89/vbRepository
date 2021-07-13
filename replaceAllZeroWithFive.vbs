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
