totNum = InputBox("Enter Total Numbers to be added", "Get Total Numbers from user")
sum =0
for a = 1 to totNum
  b = InputBox("Enter Number" & a, "Get Numbers to be added from User")
  sum = sum + b
next
Msgbox sum
