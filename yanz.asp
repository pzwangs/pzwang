<%
Function gen_key(digits)

'Create and define array
dim char_array(50)
char_array(0) = "0"
char_array(1) = "1"
char_array(2) = "2"
char_array(3) = "3"
char_array(4) = "4"
char_array(5) = "5"
char_array(6) = "6"
char_array(7) = "7"
char_array(8) = "8"
char_array(9) = "9"
char_array(10) = "A"
char_array(11) = "B"
char_array(12) = "C"
char_array(13) = "D"
char_array(14) = "E"
char_array(15) = "F"
char_array(16) = "G"
char_array(17) = "H"
char_array(18) = "I"
char_array(19) = "J"
char_array(20) = "K"
char_array(21) = "L"
char_array(22) = "M"
char_array(23) = "N"
char_array(24) = "O"
char_array(25) = "P"
char_array(26) = "Q"
char_array(27) = "R"
char_array(28) = "S"
char_array(29) = "T"
char_array(30) = "U"
char_array(31) = "V"
char_array(32) = "W"
char_array(33) = "X"
char_array(34) = "Y"
char_array(35) = "Z"

'Initiate randomize method for default seeding
randomize

'Loop through and create the output based on the the variable passed to
'the function for the length of the key.
do while len(output) < digits
num = char_array(Int((35 - 0 + 1) * Rnd + 0))
num = char_array(Int((9 - 0 + 1) * Rnd + 0))
output = output + num
loop

'Set return
gen_key = output
Session("L_YanZhengM")=gen_key
End Function

'Write the results to the browser, currently setting a 13 digit key
'response.write "<pre>" & gen_key(13) & "</pre>" & vbcrlf
'response.write "<pre>" & gen_key(5) & "</pre>" & vbcrlf
'response.write ""&Session("L_YZM")&""
%>¡¡
