#!/usr/bin/python3
#encoding: utf-8

nums=range(1,10)
print(" ",end="")

###print a line 1-9
for i in nums:
	print("%+5s"%(i),end="")
print("\n")

###print cow 1-9,line by line
for j in nums:
	###print cow num,no return.
	print("%-5s"%(j),end="")
	for i in nums:
		if i<=j:
			###print result
			print("%-5s"%(i*j),end="")
	###return a new line
	print("\n")
print("\n")
