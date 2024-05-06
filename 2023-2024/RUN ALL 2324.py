try:
    # run D1 1K Scraper
    with open(r"C:\Users\Harry Higgins\PycharmProjects\pythonProject\2023-2024\D1 1K Scraper 2324.py") as file:
        exec(file.read())
    print('D1 1K Scraper Complete (1/3)')

    # run H1 1K Scraper
    with open(r"/2023-2024/H1 1K Scraper 2324.py") as file:
        exec(file.read())
    print('H1 1K Scraper Complete (2/3)')

    # run H1 OK Scraper
    with open(r"C:\Users\Harry Higgins\PycharmProjects\pythonProject\2023-2024\H1 OK Scraper 2324.py") as file:
        exec(file.read())
    print('H1 OK Scraper Complete (3/3)')

except IOError:
    # run D1 1K Scraper
    with open(
            r"C:\Users\Harry\OneDrive\University Work\S4\Project\GitHub\KNHB_Scraper\2023-2024\D1 1K Scraper 2324.py"
    ) as file:
        exec(file.read())

    # run H1 1K Scraper
    with open(
            r"C:\Users\Harry\OneDrive\University Work\S4\Project\GitHub\KNHB_Scraper\2023-2024\H1 1K Scraper 2324.py"
    ) as file:
        exec(file.read())
    print('H1 1K Scraper Complete (2/3)')

    # run H1 OK Scraper
    with open(
            r"C:\Users\Harry\OneDrive\University Work\S4\Project\GitHub\KNHB_Scraper\2023-2024\H1 OK Scraper 2324.py"
    ) as file:
        exec(file.read())
    print('H1 OK Scraper Complete (3/3)')

print('All Programs Complete')
