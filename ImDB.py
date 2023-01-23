import requests , openpyxl
from bs4 import BeautifulSoup

file_excel = openpyxl.Workbook()

sheet = file_excel.active
sheet.title = 'Top rated movies of all time'
print(file_excel.sheetnames)
sheet.append(['Rank','Name','Year','IMDB Rating'])


try:
    source = requests.get('https://www.imdb.com/chart/top/')
    source.raise_for_status()

    soup = BeautifulSoup(source.text , 'html.parser')

    movies = soup.find('tbody' , class_="lister-list").find_all('tr')

    for movie in movies:
        rank = movie.find('td', class_='titleColumn').getText(strip = True).split('.')[0]
        name = movie.find('td' , class_='titleColumn').a.getText()
        year = movie.find('td' , class_='titleColumn').span.getText().strip('()')
        rating = movie.find('td', class_="ratingColumn imdbRating").getText()
        #print(rank, name , year , rating)
        sheet.append([rank, name , year , rating])

except Exception as e:
    print(e)


file_excel.save('Top Rated Movies.xlsx')