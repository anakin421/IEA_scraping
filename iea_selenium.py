from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium. common. exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.options import Options
import csv
import pandas as pd


options = Options()
options.add_experimental_option("prefs",{'profile.managed_default_content_settings.images': 2})
chromedriver = '/home/abhishek/chromedriver'
driver = webdriver.Chrome(executable_path= chromedriver,options=options)
driver.maximize_window()

result = dict()


def get_header():

	category_list = ['Coal','Electricity',f'Natural%20gas','Oil']

	header_list = ['Year','Country']

	for category in category_list:

		driver.get(f'https://www.iea.org/data-and-statistics/data-tables?energy={category}')

		try:
			WebDriverWait(driver,30).until(EC.presence_of_element_located((By.TAG_NAME, "table")))
		except TimeoutException:
			print("element not found")
			driver.close()
			exit()

		headers = driver.find_elements_by_css_selector('.m-data-table__th')

		for header in headers:
			header_list.append(header.text)

		for header in header_list:   #it will create all key for dict via header_list 
			result[header] = []


def scraper():

	global result 

	category_list = ['Coal','Electricity',f'Natural%20gas','Oil']
	year_list = [year for year in range(1990,2018)]

	with open('iea_country_final.csv', 'r') as file:
	    reader = csv.reader(file, skipinitialspace=True)
	    country_tuple_list = list(map(tuple, reader))

	urls = [f"https://www.iea.org/data-and-statistics/data-tables?country={country[1]}&year={year}" for country in country_tuple_list[2:] for year in year_list]

	for url in urls:  #url without category

		year = url.split('=')[-1]
		country_name = [name[0] for name in country_tuple_list[2:] if name[1]==url.split('=')[1].split('&')[0]][0] #get country name from country_tuple_list generated from csv file 
	
		result['Year'].append(year)
		result['Country'].append(country_name)

		for category in category_list:

			scraping_url = f"{url}&energy={category}"   # main scraping url

			driver.get(scraping_url)

			try:
				WebDriverWait(driver,50).until(EC.presence_of_element_located((By.TAG_NAME, "table")))

			except TimeoutException:  #in case if exception occurs
				print(f"\n element not found at url: {driver.current_url}\n")


			if scraping_url.split('=')[-1] == 'Coal':

				try:
					final_consumption = driver.find_element_by_xpath("//td[contains(text(),'Total final consumption')]")
				except NoSuchElementException:
					result['Anthracite'].append("")
					result['Coking coal'].append("")
					result['Other bituminous coal'].append("")
					result['Sub-bituminous coal'].append("")
					result['Lignite'].append("")
					result['Patent fuel'].append("")
					result['Coke oven coke'].append("")
					result['Gas coke'].append("")
					result['Coal tar'].append("")
					result['BKB'].append("")
					result['Gas works gas'].append("")
					result['Coke oven gas'].append("")
					result['Blast furnace gas'].append("")
					result['Other recovered gases'].append("")
				
				else:

					parent_tr = driver.execute_script("return arguments[0].parentNode;", final_consumption)

					total_raw_data = parent_tr.find_elements_by_css_selector('.m-data-table__data')

					sub_list = [i.text for i in total_raw_data]

					result['Anthracite'].append(sub_list[0].replace('\u202f',''))
					result['Coking coal'].append(sub_list[1].replace('\u202f',''))
					result['Other bituminous coal'].append(sub_list[2].replace('\u202f',''))
					result['Sub-bituminous coal'].append(sub_list[3].replace('\u202f',''))
					result['Lignite'].append(sub_list[4].replace('\u202f',''))
					result['Patent fuel'].append(sub_list[5].replace('\u202f',''))
					result['Coke oven coke'].append(sub_list[6].replace('\u202f',''))
					result['Gas coke'].append(sub_list[7].replace('\u202f',''))
					result['Coal tar'].append(sub_list[8].replace('\u202f',''))
					result['BKB'].append(sub_list[9].replace('\u202f',''))
					result['Gas works gas'].append(sub_list[10].replace('\u202f',''))
					result['Coke oven gas'].append(sub_list[11].replace('\u202f',''))
					result['Blast furnace gas'].append(sub_list[12].replace('\u202f',''))
					result['Other recovered gases'].append(sub_list[13].replace('\u202f',''))


			elif scraping_url.split('=')[-1] == 'Electricity':

				try:
					final_consumption = driver.find_element_by_xpath("//td[contains(text(),'Final consumption')]")
				except NoSuchElementException:
					result['Electricity'].append("")
					result['Heat'].append("")
				
				else:
					parent_tr = driver.execute_script("return arguments[0].parentNode;", final_consumption)

					total_raw_data = parent_tr.find_elements_by_css_selector('.m-data-table__data')

					sub_list = [i.text for i in total_raw_data]

					result['Electricity'].append(sub_list[0].replace('\u202f',''))
					result['Heat'].append(sub_list[1].replace('\u202f',''))



			elif scraping_url.split('=')[-1] == f'Natural%20gas':

				try:
					final_consumption = driver.find_element_by_xpath("//td[contains(text(),'Final consumption')]")
				except NoSuchElementException:
					result['Natural gas'].append("")

				else:
					parent_tr = driver.execute_script("return arguments[0].parentNode;", final_consumption)

					raw_data = parent_tr.find_element_by_css_selector('.m-data-table__data').text # 1 column

					result['Natural gas'].append(raw_data.replace('\u202f',''))


			elif scraping_url.split('=')[-1] == 'Oil':

				try:
					final_consumption = driver.find_element_by_xpath("//td[contains(text(),'Final consumption')]")

				except NoSuchElementException:
					result['Crude oil'].append("")
					result['Natural gas liquids'].append("")
					result['Refinery feedstocks'].append("")
					result['Other primary oil'].append("")
					result['LPG/Ethane'].append("")
					result['Naphtha'].append("")
					result['Motor gasoline'].append("")
					result['Jet kerosene'].append("")
					result['Other kerosene'].append("")
					result['Gas/Diesel'].append("")
					result['Fuel oil'].append("")
					result['Other oil products'].append("")

				else:

					parent_tr = driver.execute_script("return arguments[0].parentNode;", final_consumption)

					total_raw_data = parent_tr.find_elements_by_css_selector('.m-data-table__data')

					sub_list = [i.text for i in total_raw_data]

					result['Crude oil'].append(sub_list[0].replace('\u202f',''))
					result['Natural gas liquids'].append(sub_list[1].replace('\u202f',''))
					result['Refinery feedstocks'].append(sub_list[2].replace('\u202f',''))
					result['Other primary oil'].append(sub_list[3].replace('\u202f',''))
					result['LPG/Ethane'].append(sub_list[4].replace('\u202f',''))
					result['Naphtha'].append(sub_list[5].replace('\u202f',''))
					result['Motor gasoline'].append(sub_list[6].replace('\u202f',''))
					result['Jet kerosene'].append(sub_list[7].replace('\u202f',''))
					result['Other kerosene'].append(sub_list[8].replace('\u202f',''))
					result['Gas/Diesel'].append(sub_list[9].replace('\u202f',''))
					result['Fuel oil'].append(sub_list[10].replace('\u202f',''))
					result['Other oil products'].append(sub_list[11].replace('\u202f',''))


	driver.close()

	result = pd.DataFrame(result)

	writer = pd.ExcelWriter('iea_data_final.xlsx', engine='xlsxwriter')
	result.to_excel(writer, sheet_name='Sheet1',index=False)
	writer.save()


if __name__=='__main__':
	get_header()
	scraper()