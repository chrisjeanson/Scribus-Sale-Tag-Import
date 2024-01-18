"""
Script Name: Scribus Sale Tag Importer

Description: This script retrieves product sale price and upc for quick
tagging of said product. This script will create tags for any listed on 
the CSV and will create new pages as needed.

Requirements: 
Data CSV that has 3 columns: qty, sale price, and upc. The included Scribus 
document is already configured to work with the script. 

Scribus 1.6x or above (could work on previous versions as well), some
sort of software to create/manipulate csv files.

Set an environment variable for the location of your data file called 'SCRIBUS_DATA_PATH'.
Or, conversely, manually enter a file path in place of the environment variable.

Usage: Load up the document.sla file in Scribus. Then use Script > Execute Script... 
to select the scribus.py script. It will replace

License: MIT License

Contact: [cj@spite.ca]
"""
import os
import csv
import scribus # Will be used interally by Scribus when run within the program

# List of element names from the template tag in Scribus
template_elements = ['border', 'price', 'upc', 'sale_bg', 'sale', 'line']

# Function to duplicate the base tag
def duplicate_template(row, col, data):
    # Calculate the position for the new tag
    offset_x, offset_y = calculate_position(row, col)
    new_elements = {}

    for element_name in template_elements:
        if not scribus.objectExists(element_name):
            print(f"Error: Object named '{element_name}' does not exist in the document.")
            continue

        scribus.copyObject(element_name)
        pasted_element = scribus.pasteObject()  # Get the name of the pasted object

        # Immediately move the pasted object
        scribus.moveObject(offset_x, offset_y, pasted_element)

        # Store the name of the pasted object with a reference to the original name
        new_elements[element_name] = pasted_element

    # Update text and other properties for the new tag
    update_tag_data(row, col, data, new_elements)

    return new_elements

# Function to update the newly created tag fields
def update_tag_data(row, col, data, new_elements):
    # Unpack data to get qty, price, and upc
    qty, price, upc = data
    qty = int(qty)
    price = str(price)
    upc = str(upc)

    # Correctly reference the new element names
    price_frame_name = new_elements['price']
    upc_frame_name = new_elements['upc']

    # Now check if the frame exists and update the text
    if scribus.objectExists(price_frame_name):
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

# Function to calculate the position of newly created tags
def calculate_position(row, col):
    page_height = 23 # 8.5" (21.59cm) converted to cm plus a lil love
    # Define the starting position (top-left corner of the first tag)
    start_x = 4.1  # Starting x position (in cm) relative to the off-canvas tag
    # start_y = 3.2  # Starting y position (in cm) relative to the off-canvas tag
    start_y = 3.2 + ((current_page - 1) * page_height)  # Adjust for the current page

    # Define the spacing between tags
    horizontal_spacing = 3.1  # Spacing to the right (in cm) 3.1
    vertical_spacing = 2.6    # Spacing down (in cm) 2.6

    # Calculate the position for the new tag
    x = start_x + (col - 1) * horizontal_spacing
    y = start_y + (row - 1) * vertical_spacing

    return x, y

# Function to clear a tag
def clear_tag(row, col):
    price_frame_name = f"price_{row}_{col}"
    upc_frame_name = f"upc_{row}_{col}"

    if scribus.objectExists(price_frame_name):
        scribus.deleteText(price_frame_name)
    if scribus.objectExists(upc_frame_name):
        scribus.deleteText(upc_frame_name)

# Read the CSV file using an environment variable
data_file_path = os.path.join(os.getenv('SCRIBUS_DATA_PATH', 'default_path_if_not_set'), 'data.csv')

with open(data_file_path, 'r') as csvfile:
    datareader = csv.reader(csvfile)
    next(datareader)  # Skip the header row
    current_page = 1
    row = 1
    col = 1
    tags_filled = 0
    total_tags_per_page = 56  # Total number of tags on a page (8 cols x 7 rows)

    for data in datareader:
        qty, price, upc = data[:3]

        for _ in range(int(qty)):
            if tags_filled >= total_tags_per_page:
                # Add a new page
                scribus.newPage(-1)  # -1 appends the page at the end
                current_page += 1
                row = 1
                col = 1
                tags_filled = 0  # Reset the counter for the new page

            new_element_names = duplicate_template(row, col, data)
            update_tag_data(row, col, data, new_element_names)
            tags_filled += 1

            # Update row and column counters
            col += 1
            if col > 8:  # Assuming 8 columns per page
                col = 1
                row += 1
                if row > 7:  # Assuming 7 rows per page, reset for the new page
                    row = 1
                    col = 1
                    scribus.newPage(-1)
                    current_page += 1
                    tags_filled = 0
    
    # Clear remaining tags if not all are filled
    while tags_filled < total_tags_per_page:
        clear_tag(row, col)
        tags_filled += 1
        col += 1
        if col > 8:
            col = 1
            row += 1