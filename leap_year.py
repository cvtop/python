#!/usr/bin/python3
#encoding: UTF-8

year=input("pls input a year: ")
year=int(year)
#print(year)

if year % 4 == 0 and year % 100 != 0:
	print("%-6s is a leap year"%(year))
elif year % 400 == 0 :
	print("%-6s is a leap year"%(year))
else:
	print("%-6s is NOT a leap year"%(year))
