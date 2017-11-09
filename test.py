import requests
from bs4 import BeautifulSoup
import re
from xlwt import Workbook
import csv
web_link= "https://cottonon.com/AU/"
headers={'User-Agent': 'Mozilla/5.0 (Windows NT 6.3; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/54.0.2840.71 Safari/537.36'}
r= requests.get(web_link,headers)
soup=BeautifulSoup(r.content,"lxml")
#----------------------------------------------------------------------------
#extract Category
sp2=soup.find_all('a',{"class":"topLevelCatButton "})
#----------------------------------------------------------------------------
wb = Workbook()
for i in range(len(sp2)):
	product_cat=sp2[i].text.split()
	print(product_cat)
	#------------------------------------------------------------------------
	#create sheet for each product
	sheet = wb.add_sheet(product_cat[0])
	#------------------------------------------------------------------------
	row_count=0
	#------------------------------------------------------------------------
	#extract category link
	extracted_link=sp2[i].get('data-href')
	#------------------------------------------------------------------------
	#print(extracted_link)
	#------------------------------------------------------------------------
	#Enters in category link
	inner_link=requests.get(extracted_link,headers)
	soup_inner=BeautifulSoup(inner_link.content,'lxml')
	#------------------------------------------------------------------------
	#extract sub_category
	sp1=soup_inner.find_all('a',{"class":"refinement-link "})
	#------------------------------------------------------------------------
	#loop for sub_cat
	for j in range(int (len(sp1)/2)):
		col_count=0
		row_count+=2
		#--------------------------------------------------------------------
		#get sub_category text
		pro_sub_cat=sp1[j].text.strip()
		print(pro_sub_cat)
		#--------------------------------------------------------------------
		#get sub_category link
		inner_extracted_link=sp1[j].get('href')
		#--------------------------------------------------------------------
		#print(inner_extracted_link)
		#--------------------------------------------------------------------
		#enters in sub_category link
		extracted_link_sub=requests.get(inner_extracted_link,headers)
		#--------------------------------------------------------------------
		super_soup=BeautifulSoup(extracted_link_sub.content,'lxml')
		#--------------------------------------------------------------------
		#extract super_sub_cat
		sp3=super_soup.find_all('a',{"class":"refinement-link "})
		#--------------------------------------------------------------------
		try:
			#check for sub_cat toggle
			sp5=super_soup.find('h3',{"class":"toggle dropdown-title"}).text.strip()
		except Exception as e:
			sp5=" "
		try:
			#count sub_category products
			sub_count=super_soup.find('span',{"class":"paging-information-items"}).text.strip()
		except Exception as e:
			sub_count=" "
		print(sub_count)
		#---------------------------------------------------------------------
		#write sub_cat in product sheet
		if col_count==0:
			sheet.write(row_count,col_count,pro_sub_cat)
			sheet.write(row_count+1,col_count,sub_count)
			#-----------------------------------------------------------------
		col_count+=1
		#check if sub_cat toggle is same as product, terminate
		if sp5==product_cat[0]:
			continue
			#-----------------------------------------------------------------
		#loop for super_sub_cat
		for k in range(int (len(sp3)/2)):
			super_sub_cat=sp3[k].text.strip()
			print(super_sub_cat)
			#-----------------------------------------------------------------
			#get link of super_sub_cat
			super_inner_link=sp3[k].get('href')
			#-----------------------------------------------------------------
			#print(super_inner_link)
			super_inner_link_=requests.get(super_inner_link,headers)
			super_sub_soup=BeautifulSoup(super_inner_link_.content,'lxml')
			try:
				#get no of super_sub_cat product
				super_sub_count=super_sub_soup.find('span',{"class":"paging-information-items"}).text.strip()
			except Exception as e:
				super_sub_count=" "
			print(super_sub_count)
			#write super_sub_cat in product sheet
			if col_count>0:
				sheet.write(row_count,col_count,super_sub_cat)
				sheet.write(row_count+1,col_count,super_sub_count)
			col_count+=1
wb.save('test.xls')


