The Artemis Scraper code (or artemis Scraper python file) retrieves key data, such as spread, expected loss, maturity, attachment point, etc. from Artemis - the main catastrophe bond directory, by opening each deal and analyzing it with regex patterns and other language processing methods.
The data is then returned in an excel sheet (Transactions_Chart). 
If the Transactions_Chart file already exists, running the code will update it by creating new rows with new transactions that have not been scraped yet (if there are any).
If it does not exist, it will be created and the first 1000 transactions on the Artmeis directory will be scarped. 
The code is not 100% accurate in all the datapoints. The transactions_chart file contains many deals that have been manually checked. 
