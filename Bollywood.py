import requests, openpyxl
from bs4 import BeautifulSoup

excel = openpyxl.Workbook()
sheet = excel.active
sheet.title = 'Top Rated Hindi Movies'
sheet.append(['Movie Rank', 'Movie Name', 'Year', "IMDB Rating"])

try:
    source = requests.get('https://www.imdb.com/india/top-rated-indian-movies')
    source.raise_for_status()
    
    soup = BeautifulSoup(source.text, 'html.parser')
    movies = soup.find('tbody', class_="lister-list").find_all('tr')
    year = soup.find('tbody', class_="lister-list").find_all('tr')
    
    for movie in movies:
        rank = movie.find('td', class_="titleColumn").get_text(strip=True).split('.')[0]
        name = movie.find('td', class_="titleColumn").a.text
        year = movie.find('td', class_="titleColumn").span.text.strip('()')
        rating = movie.find('td', class_="ratingColumn").strong.text
        print (rank, name, year, rating)
        sheet.append([rank, name, year, rating])
    
except Exception as e:
    print(e)
    
file = 'IMDBRatings.xlsx'
excel.save(file)
