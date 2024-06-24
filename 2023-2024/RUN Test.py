try:
    # run Test Scraper
    with open(
            r"C:\Users\Harry Higgins\PycharmProjects\pythonProject\2023-2024\Individual Scrapers\Test.py"
    ) as file:
        exec(file.read())
    print('Test successful')

except IOError:
    with open(
            r"C:\Users\Harry\OneDrive\University Work\S4\Project\GitHub\KNHB_Scraper\2023-2024\Individual Scrapers\Test.py"
    ) as file:
        exec(file.read())
    print('Test successful')
