dim i, j

for i=3 to 50
Prime = false
'isPrime = false
	for j = 2 to i-1
	'msgbox i Mod j <> 0
	   modval=i Mod j
	   'msgbox modval
		if modval = 0 then
			Prime = False
			exit for
		else
            Prime = True		
		end if
	next
		
  	if(Prime) then
		msgbox i
'msgbox Prime
	end if
next
