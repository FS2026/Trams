## Trams
# Tram Traffic Analysis in Warsaw
This project analyzes tram traffic intensity in Warsaw based on schedule data from ZTM. 

1. Input: Excel file with URLs to tram timetable pages (based on ZTM data).
2. Scraping: Python script using `requests` and `BeautifulSoup` collects the number of daily trams arrivals from each tram stop.
3. Visualization: Output imported into QGIS to generate a map showing tram line activity.

# Technologies Used
- Python (`openpyxl`, `requests`, `bs4`)
- QGIS (for spatial visualization)
