# DMC Ecommerce Webscraper

DMC Ecommerce Webscraper is a software designed for the Digital Multumedia Center at MSU Libraries to automate the process of finding deals for future catalogue items on ecommerce websites. It facilitates the extraction of product information and screenshots, organizing them neatly into an Excel spreadsheet for easy reference and analysis. Previously this was an intensely laborious task requiring manually surfing websites often several dozens of times copy pasting titles and noting down prices and links. Automating this has saved hours every time the catalogue needs to be updated.

## Installation

To install DMC Ecommerce Webscraper, follow these steps:

1. Download the Microsoft Edge WebDriver for Selenium:
    - [Microsoft WebDriver](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/?form=MA13LH)
    - Download the latest stable release.
    - Ensure that it is saved to program files (C:\Program Files (x86)\msedgedriver.exe).

2. Download the Executable and the Format:
    - [All Releases](https://github.com/Swefton/DMC-ecommerce-webscraper/releases).


3. Ensure that Local Microsoft Edge and WebDriver are the same version:
    - The local edge browser and the webdriver needs to be the same version. Whenever the webdriver is installed, also update edge and then try not to update it again.
    - This should be easy because nobody uses edge anyway.

## Usage

To use DMC Ecommerce Webscraper, follow these steps:

1. Fill up the games list format:
    - Rows A3 to F3 Contain headings.
    - Fill Up Video game titles in the first row and "y" indicating yes for the appropriate platforms.

2. Run the Main Executable.

3. This will generate a xlsm file with the filled up information.

## Challenges

- Certain websites had bot checks which meant that my script had to beat these checks. To overcome this I used certain flags and techniques to avoid certain actions that would trigger the bot detection.
- Websites update their UI often which mean that my script would break, to overcome this I avoided using XPATHs to make the code more dynamic.

## Future Plans

- [ ] Currently the script only checks the first item when searched. Moving forward I want for it to check every element on the page.
- [ ] Sometimes the first result on the webpage isn't the correct game or it is for the wrong platform etc. I want to use a string similarity index check to error check for these.

```python
import difflib

def string_similarity(string1, string2):
    return difflib.SequenceMatcher(None, string1, string2).ratio()

# Example usage
user_product = "Nike Air Max 90"
product_listing_title = "Nike Air Max 90 Sneakers"

similarity_score = string_similarity(user_product.lower(), product_listing_title.lower())
print("Similarity score:", similarity_score)
