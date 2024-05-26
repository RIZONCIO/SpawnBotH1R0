'Convert UTF-8 file to ANSI 
set shell = createobject ("wscript.shell")

strtext = inputbox ("Adicione a frase que vocé gostaria de fazer spam")
strtimes = inputbox ("Qual e a quantidade de vezes que voce gostaria de spam ")
strspeed = inputbox ("Quao rapido voce gosta de spam? (1000 = um por segundo, 100 = 10 por segundo, etc.)")
strtimeneed = inputbox ("Quantos segundos voce precisa para chegar a caixa de entrada de suas vitimas?")

If not isnumeric (strtimes & strspeed & strtimeneed) then
msgbox "Voce inseriu algo que nao e numerico. Tente novamente"
wscript.quit
End If
strtimeneed2 = strtimeneed * 1000
do
msgbox "Voce tem " & strtimeneed & " segundos para chegar a sua area de entrada onde voce vai enviar spam."
wscript.sleep strtimeneed2
for i=0 to strtimes
shell.sendkeys (strtext & "{enter}")
wscript.sleep strspeed
Next
wscript.sleep strspeed * strtimes / 10
returnvalue=MsgBox ("Deseja enviar spam novamente com a mesma informaçao?",36)
If returnvalue=6 Then
Msgbox "Ok SpamBotH1R0 vai começar de novo"
End If
If returnvalue=7 Then
msgbox "SpamBotH1R0 indo dormir"
wscript.quit
End IF
loop