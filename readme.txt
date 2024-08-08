The Artemis Scraper code (or artemis Scraper python file) retrieves key data, such as spread, expected loss, maturity, attachment point, etc. from Artemis (https://www.artemis.bm/dashboard/) - the main catastrophe bond directory, by opening each deal article and analyzing it with regex patterns and other language processing methods. The data is then returned in an excel sheet (Transactions_Chart). If the Transactions_Chart file already exists, the code will only open new deals and update the excel sheet by adding new rows below it with transactions that have not been scraped yet (if there are any). If it does not exist, the code will create a sheet with that name and scrape the last 1000 transactions on the Artmeis directory.

The Pricing_Chart file shows regressions of spread on expected loss based on a set number of parameters.

IMPORTANT: 
Before running, install the relevant modules (bs4, selenium, openpyxl) by running "pip install ___ " "in the console.
To run the code, you need to have on your local computer (ex in desktop) the "Pricing_Chart" file with a sheet Called "All_transaction"
