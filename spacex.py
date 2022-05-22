from bs4 import BeautifulSoup 
import requests
import xlsxwriter
import schedule

# start new wb and title xlsx
wb = xlsxwriter.Workbook('spacex.xlsx')
#add ws from wb
ws = wb.add_worksheet()

# define source
source = requests.get('https://spacecoastlaunches.com/launch-list/').text

# num var = 0
num =0


center1 = wb.add_format({'align': 'center' , 'valign' : 'vcenter'})
center1.set_font_size(13)
center1.set_bg_color("3F387C")
center1.set_font_color("white")
center1.set_bold()
center1.set_font_name("Times New Roman")


#center = wb.add_format({'valign' : 'vcenter'})

# write titles for ws in first row
# (row ind, col ind, 'text')
ws.write(0, 0, 'Date', center1) 
ws.write(0, 1, 'Vehicle', center1) 
ws.write(0, 2, 'Mission', center1) 
ws.write(0, 3, 'Launch Site', center1) 
ws.write(0, 4, 'Launch Window',center1) 



# add text wrap formatting
wrap = wb.add_format({'text_wrap': True})
wrap.set_font_size(13)
wrap.set_valign('vcenter')
wrap.set_font_name("Times New Roman")




center = wb.add_format({'align': 'center' , 'valign' : 'vcenter'})
center.set_font_size(13)
center.set_font_name("Times New Roman")

#Date
ws.set_column(0,1,25, center)
#vehicle
ws.set_column(1,2,42, center)
#mission
ws.set_column(2,3,65,wrap)
#ls
ws.set_column(3,4,45, center)
#lw
ws.set_column(4,5,25, center)

#link
#ws.set_column(5,6,35, center)


# Set every row height to 100
ws.set_default_row(115)

# set TOP row height to 25 
ws.set_row(0,30)


# Define soup
soup = BeautifulSoup(source)



yellow = wb.add_format({'bg_color': 'yellow'})

ws.conditional_format('A1:E7', {'type': 'text',
                                'criteria': 'containing',
                                 'value': "SpaceX",
                                 'format': yellow})

# For every div with class "three-fourth et-column-last"
for div in soup.findAll("div", {"class": "three_fourth et_column_last"}):
	print(div)

	# DATE
	# from each div select the p with index num 0 and get text
	# split it a the : and get the 2nd part (index 1 aka 2nd)
	d = div.select("p")[0].get_text(strip = True).split(":")[1]
	#print(d)



	# VEHICLE
	v = div.select("p")[1].get_text(strip = True).split(":")[1]
	#print(v)




	# MISSION
	m = div.select("p")[2].get_text(strip = True).split(":",1)[1]
	#print(m)







	# LAUNCH SITE
	ls = div.select("p")[3].get_text(strip = True).split(":")[1]
	#print(ls)





	# LAUNCH WINDOW
	lw = div.select("p")[4].get_text(strip = True).split(":",1)[1]
	#print(lw)


	# add one to num every loop
	num +=1


	# writing into ws
	ws.write(num, 0, d)
	ws.write(num, 1, v)
	ws.write(num, 2, m)
	ws.write(num, 3, ls)
	ws.write(num, 4, lw)


	"""
	# Find link to pictures
	numpic = 0
	for div in soup.findAll("div", {"class": "one_fourth"}):
		pic = div.img['src']
		print(pic)
		numpic +=1
		ws.write(numpic, 5, pic)

	"""





# close wb 
wb.close()



