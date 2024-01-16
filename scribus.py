"""
Script Name: Scribus Sale Tag Importer

Description: This script automates the process of extracting product sale pricing
and UPC from any csv file. 

Requirements: 
Data CSV that has 3 columns: qty, sale price, and upc. Since the default Scribus
document only has room for 56 tags, the CSV should be limited to 56 total items
at a time. If you want to do more, duplicate as required. Be sure if making more
pages, that you continue the naming structure that is present on page one (label
each price and upc text field based on row_column notation)

Scribus 1.6x or above (could work on previous versions as well), some
sort of software to create/manipulate csv files.

Set an environment variable for the location of your data file called 'SCRIBUS_DATA_PATH'.
Or, conversely, manually enter a file path in place of the environment variable.

Usage: Load up the document.sla file in Scribus. Then use Script > Execute Script... 
to select the scribus.py script. It will replace

License: MIT License

Contact: [cj@spite.ca]
"""
import os # to retrieve the environment variable for the data file location
import csv
import scribus # library will only be recognized within scribus at runtime

# Function to create a tag
def create_tag(row, col, price, upc):
    # Frame names based on row and column
    price_frame_name = f"price_{row}_{col}"
    upc_frame_name = f"upc_{row}_{col}"

    # Prepend a dollar sign to the price if it's not already there
    price_with_dollar = f"${price}" if not price.startswith('$') else price

    # Check if the frame exists and update the text
    if scribus.objectExists(price_frame_name):
        # Set text for price
        scribus.setText(price_with_dollar, price_frame_name)

        # Set font size for price
        scribus.setFontSize(17, price_frame_name)

        # Set font for price
        scribus.setFont('Verdana Bold', price_frame_name)

        # Set alignment to center for price
        scribus.setTextAlignment(1, price_frame_name)  # Using 1 for center alignment

    if scribus.objectExists(upc_frame_name):
        # Set text for UPC
        scribus.setText(upc, upc_frame_name)

        # Set font size for UPC
        scribus.setFontSize(8, upc_frame_name)

        # Set font for UPC
        scribus.setFont('Tahoma Regular', upc_frame_name)

        # Set alignment to center for UPC
        scribus.setTextAlignment(1, upc_frame_name)  # Using 1 for center alignment

# Function to clear a tag
def clear_tag(row, col):
    price_frame_name = f"price_{row}_{col}"
    upc_frame_name = f"upc_{row}_{col}"

    if scribus.objectExists(price_frame_name):
        scribus.deleteText(price_frame_name)
    if scribus.objectExists(upc_frame_name):
        scribus.deleteText(upc_frame_name)

# Read the CSV file using a file path with linux or windows
# with open('linux/filelocation/goes/here/data.csv', 'r') as csvfile:
# with open('C:\\windows\\filelocation\\goes\\here\\data.csv', 'r') as csvfile:

# Read the CSV file using an environment variable
data_file_path = os.path.join(os.getenv('SCRIBUS_DATA_PATH', 'default_path_if_not_set'), 'data.csv')

with open(data_file_path, 'r') as csvfile:
    datareader = csv.reader(csvfile)
    next(datareader)  # Skip the header row
    row = 1
    col = 1
    tags_filled = 0
    total_tags = 56  # Total number of tags on a page (8 cols x 7 rows)

    for data in datareader:
        qty, price, upc = data[:3]  # Reading qty, price, and UPC

        for _ in range(int(qty)):
            create_tag(row, col, price, upc)
            tags_filled += 1

            # Update row and column counters
            col += 1
            if col > 8:  # Reset column and move to the next row
                col = 1
                row += 1
            if tags_filled >= total_tags:
                break

        if tags_filled >= total_tags:
            break
    
    # Clear remaining tags if not all are filled
    while tags_filled < total_tags:
        clear_tag(row, col)
        tags_filled += 1
        col += 1
        if col > 8:
            col = 1
            row += 1