from selenium import webdriver
from time import sleep
import os
import xlsxwriter
import warnings
warnings.filterwarnings("ignore", category=DeprecationWarning)

options = webdriver.ChromeOptions()
options.add_argument('--headless')
options.add_argument('--log-level=3')
options.page_load_strategy = 'normal'
driver = webdriver.Chrome(executable_path=r'C:\Program Files (x86)\Google\Driver\chromedriver.exe', chrome_options=options)
driver.get("https://axie.zone/leaderboard")
dir = os.path.dirname(__file__)
file_path = os.path.join(dir, 'lb_data.xlsx')
workbook = xlsxwriter.Workbook(file_path)
worksheet = workbook.add_worksheet('lb_data')
worksheet.write(0, 0, 'Rank')
worksheet.write(0, 1, 'Name')
worksheet.write(0, 2, 'Points')
col = 3
for x in range(3):
	worksheet.write(0, col + x * 6, "Axie " + str(x + 1) + " Eyes")
	worksheet.write(0, col + x * 6 + 1, "Axie " + str(x + 1) + " Ears")
	worksheet.write(0, col + x * 6 + 2, "Axie " + str(x + 1) + " Back")
	worksheet.write(0, col + x * 6 + 3, "Axie " + str(x + 1) + " Mouth")
	worksheet.write(0, col + x * 6 + 4, "Axie " + str(x + 1) + " Horn")
	worksheet.write(0, col + x * 6 + 5, "Axie " + str(x + 1) + " Tail")
for x in range(100):
	for classes in driver.find_elements_by_class_name('lb_rank' + str(x + 1)):
		classes.find_element_by_xpath('.//a[starts-with(@href, "/profile")]').click()
	sleep(5)
	if driver.find_elements_by_xpath('//*[contains(text(), "No profile data found!")]'):
		print('Skipping ' + str(x + 1) + '(no data)')
		driver.back()
		driver.execute_script('window.scrollTo(0,' + str((x + 1) * 45) + ')')
		sleep(1)
		continue
	print('Progress: ' + str(x + 1) + '/100')
	worksheet.write(x + 1, 0, str(x + 1))
	player_name = driver.find_elements_by_class_name("ib")[0].text
	worksheet.write(x + 1, 1, player_name)
	player_points = driver.find_element_by_xpath('.//td[text()="Points: "]').text
	worksheet.write(x + 1, 2, player_points.removeprefix('Points: '))
	for y in range(3):
		eyes = driver.find_element_by_xpath('(.//div[starts-with(@title, "Eyes")])[' + str(y + 1) + ']').get_attribute("title")
		worksheet.write(x + 1, col + y * 6, eyes.removeprefix('Eyes: '))
		ears = driver.find_element_by_xpath('(.//div[starts-with(@title, "Ears")])[' + str(y + 1) + ']').get_attribute("title")
		worksheet.write(x + 1, col + y * 6 + 1, ears.removeprefix('Ears: '))
		back = driver.find_element_by_xpath('(.//div[starts-with(@title, "Back")])[' + str(y + 1) + ']').get_attribute("title")
		worksheet.write(x + 1, col + y * 6 + 2, back.removeprefix('Back: '))
		mouth = driver.find_element_by_xpath('(.//div[starts-with(@title, "Mouth")])[' + str(y + 1) + ']').get_attribute("title")
		worksheet.write(x + 1, col + y * 6 + 3, mouth.removeprefix('Mouth: '))
		horn = driver.find_element_by_xpath('(.//div[starts-with(@title, "Horn")])[' + str(y + 1) + ']').get_attribute("title")
		worksheet.write(x + 1, col + y * 6 + 4, horn.removeprefix('Horn: '))
		tail = driver.find_element_by_xpath('(.//div[starts-with(@title, "Tail")])[' + str(y + 1) + ']').get_attribute("title")
		worksheet.write(x + 1, col + y * 6 + 5, tail.removeprefix('Tail: '))
	driver.back()
	driver.execute_script('window.scrollTo(0,' + str((x + 1) * 45) + ')')
	sleep(1)
workbook.close()
input('Completed, press enter to exit')
driver.quit()
exit()