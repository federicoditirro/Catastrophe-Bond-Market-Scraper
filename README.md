# Artemis Scraper

## Overview

The Artemis Scraper retrieves key data such as spread, expected loss, maturity, attachment point, etc., from [Artemis](https://www.artemis.bm/dashboard/), the main catastrophe bond directory. It works by opening each deal article on the Artemis platform and analyzing the content using regex patterns and other language processing methods. The extracted data is then saved in an Excel sheet named `Transactions_Chart.xlsx`.

### Key Features:
- **Incremental Data Retrieval:** If `Transactions_Chart.xlsx` already exists, the scraper will only add new transactions to it by fetching data from deals that haven't been scraped yet. If the file doesn't exist, the scraper will create it and populate it with the latest 1000 transactions from the Artemis directory.
- **Pricing Analysis:** The scraper can also generate a `Pricing_Chart.xlsx` file that shows regressions of spread on expected loss based on a set number of parameters.

## Getting Started

### Prerequisites

Before running the scraper, ensure that the necessary Python modules are installed. The required modules are listed in the `requirements.txt` file.

### Installation

1. **Clone the Repository:**

   ```bash
   git clone https://github.com/yourusername/artemis-scraper.git
   cd artemis-scraper
   ```

2. **Install Dependencies**

   Next, install the necessary Python packages using pip:

   ```bash
   pip install -r requirements.txt 
   ```
   This will install all the required libraries listed in the `requirements.txt` file.

## Usage
### Running the Scraper
To run the scraper, simply execute the `artemis_scraper.py` script:

```bash
python artemis_scraper.py
```

The script will begin scraping the Artemis Deal Directory, extracting the specified data points from each deal article.

### Output
The scraped data will be automatically saved to an Excel file named `Transactions_Chart.xlsx`. The scraper behaves as follows:

If the `Transactions_Chart.xlsx` file already exists, the code will only open new deals and update the Excel sheet by adding new rows with transactions that have not been scraped yet (if any).
If the `Transactions_Chart.xlsx` file does not exist, the code will create the file and scrape the last 1000 transactions from the Artemis directory.
Additionally, the `Pricing_Chart.xlsx` file shows regressions of spread on expected loss based on a set number of parameters.

## Customisation

### Specifying Data Points
If you want to customize which data points are extracted, you can modify the `artemis_scraper.py` script. Locate the section where data is parsed and add or remove fields according to your needs.

## License
This project is licensed under the MIT License. See the **LICENSE** file for more details.