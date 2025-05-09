import ssl
import urllib.request
from bs4 import BeautifulSoup as bs
import os

def download_pdf(_id, download_path):

	context = ssl._create_unverified_context()
	url = f"https://www.kap.org.tr/tr/Bildirim/{_id}"

	response = urllib.request.urlopen(url.replace('\n',''), context = context)
	content = response.read()

	soup = bs(content, 'html.parser')
	links = soup.find_all('a')

	for link in links:
		href = str(link['href'])

		file_count = 0
		if 'https://www.kap.org.tr/tr/api/file/download/' in href:
			print(href)
			response = urllib.request.urlopen(href, context = context)
			content = response.read()

			file_name = f"{str(href).split('/')[-1]}.pdf"
			file_path = os.path.join(download_path, file_name)

			with open(file_path, 'wb') as file:
				file.write(content)
