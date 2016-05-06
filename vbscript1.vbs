option explicit

Dim abc,i,j

abc = "Hai how are you"

msgbox abc

Dim myArr(5)

myArr(0) = "Hai"
myArr(1) = "Hello"
myArr(2) = "My"
myArr(3) = "Name"
myArr(4) = "Don't know"

for i=0 to 5-1
	msgbox myArr(i)
next


Dim marr()
redim preserve marr(5)
marr(0) = "i"
marr(1) = "disco"
marr(2) = "dancer"



marr(3) = "hello again"
marr(4) = "trying"

for j=0 to 5-1
MsgBox marr(j)
next