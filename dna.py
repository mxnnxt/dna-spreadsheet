import xlwt
import string
import random
from random import getrandbits

print ("")
print (" --------------------------------------------")
print ("          DNA SPREADSHEET CREATOR            ")
print ("           Developed by @mxnnxt              ")
print (" --------------------------------------------\n")

#user input
times = int(input("Enter the number of times you would like to fill: "))
profilename = str(input("Enter Profile Name: "))
sku = str(input("Enter the SKU: "))
addy1 = str(input("Enter Address 1: "))
city = str(input("Enter City: "))
state = str(input("Enter State (ex. NY) : "))
area = str(input("Enter Zip Code: "))
country = str(input("Enter Country (ex. US) : "))
phone = input("Phone Number Prefix: ")

book = xlwt.Workbook(encoding="utf-8")
#create sheet
sheet1 = book.add_sheet("Sheet 1")
#create columns
sheet1.write(0, 0, "Profile")
sheet1.write(0, 1, "Site")
sheet1.write(0, 2, "Email")
sheet1.write(0, 3, "Password")
sheet1.write(0, 4, "Product SKU")
sheet1.write(0, 5, "Size")
sheet1.write(0, 6, "First Name")
sheet1.write(0, 7, "Last Name")
sheet1.write(0, 8, "Address Line 1")
sheet1.write(0, 9, "Address Line 2")
sheet1.write(0, 10, "City")
sheet1.write(0, 11, "State")
sheet1.write(0, 12, "Zip")
sheet1.write(0, 13, "Country")
sheet1.write(0, 14, "Phone")
sheet1.write(0, 15, "Card Type")
sheet1.write(0, 16, "Card Number")
sheet1.write(0, 17, "Expiry Month (MM)")
sheet1.write(0, 18, "Expiry Year (YYYY)")
sheet1.write(0, 19, "CVV")


#start under column titles
i=2
#write data
for i in range(times):
	i = i+1
	#profile
	paste = profilename+str(i)
	sheet1.write(i, 0, paste)
	#SITE
	site = "SNKRS"
	sheet1.write(i, 1, site)
	#SKU
	sheet1.write(i, 4, sku)
	#size
	size_run = ['7','7.5','8','8.5','9','9.5','10','10.5','11','11.5','12','12.5','13','14']
	size = random.choice(size_run)
	sheet1.write(i, 5, size)
	#firstname
	names = ["Beck","Glenn","Becker","Carl","Beckett","Samuel","Beddoes","Mick","Beecher","HenryWard","Beethoven","Ludwigvan","Begin","Menachem","Bell","Alexander","Graham","Belloc","Hilaire","Bellow","Saul","Benchley","Robert","Benenson","Peter","BenGurion","David","Benjamin","Walter","Benn","Tony","Bennington","Chester","Benson","Leana","Bent","Silas","Bentsen","Lloyd","Berger","Ric","Bergman","Ingmar","Berio","Luciano","Berle","Milton","Berlin","Irving","Berne","Eric","Bernhard","Sandra","Berra","Yogi","Berry","Halle","Berry","Wendell","Bethea","Erin","Bevan","Aneurin","Bevel","Ken","Biden","Joseph","Bierce","Am","Brose","Biko","Steve","Billings","Josh","Biondo","Frank","Birrell","Augustine","Black","Elk","Blair","Ro","Bert","Blair","Tony","Blake","William","Blakey","Art","Blalock","Jolene","Blanc","Mel","Blanc","Raymond","Blanchet","Cate","Blix","Hans","Blood","Rebecca"]
	firstName = names[random.randint(0, 99)]
	sheet1.write(i, 6, firstName)
	#lastname
	lastName = names[random.randint(0, 99)]
	sheet1.write(i, 7, lastName)
	#address line 1
	size = 4
	chars1 = string.ascii_uppercase + string.digits
	chars2 = ''.join(random.choice(chars1) for _ in range(size))
	addy2 = chars2+" "+addy1
	sheet1.write(i, 8, addy2)
	#City
	sheet1.write(i, 10, city)
	#state
	sheet1.write(i, 11, state)
	#zip
	sheet1.write(i, 12, area)
	#country
	sheet1.write(i, 13, country)
	#phone
	number5 = random.sample(range(10), 7)
	num2 = str((''.join(map(str, number5))))
	phone_num = phone+num2
	sheet1.write(i, 14, phone_num)


print("SUCCESSFULLY SAVED TO SPREADSHEET")
book.save("dnaNike.xls")
