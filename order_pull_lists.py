# This script creates the pull list and order list using the cleaned bookstore data.
# Last modified: 01-19-24 by GI

import pandas as pd     # Pandas library to handle dataframes and spreadsheet manipulation.
import requests         # HTTP requests library used to scrape the catalog.
import re               # Regular Expressions library used to parse strings and retrieve relevant data from large blocks of text on a webpage.
import time             # Allows us to track how long the script takes to run for assessment purposes.
import os               # Grants access to our machine's operating system so that we can access and export files on the G:/ drive.

# Set the start time before the script begins running.
start = time.time()
pd.options.mode.chained_assignment = None

# ***********************
# User Input & File Setup
# ***********************

# Retrieve semester from user. This is then used to isolate the semester so we can create a directory path to the relevant files. We also save a list of all the files
# beginning with the semester name so that we can display them to the user. This prevents having to open the file explorer to get the exact file name.
sem_folder = input("Enter the semester in [semester] [year] format (ex: Fall 2023). Note: this is case-sensitive!\n")
sem = ''.join(re.findall('^\W*([\w-]+)', sem_folder))
sem_file_path = os.listdir("G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\semesters\\" + sem_folder + "\\")
file_options = []
for file in sem_file_path: 
    if file.startswith(sem + "Book"): file_options.append(file)

# Display the possible bookstore list options to the user. Then, retrieve exact bookstore list file name from user.
print("\nBookstore lists in this directory:\n")
print("\n".join(file_options))
bkstr_file_name = input("\nEnter the full name of the bookstore list spreadsheet (ex: FallBookstoreList 9-8-2023). Note: this is case-sensitive!\n")

# Ask if there is a previous bookstore list to compare to from user.
prev_ord_lists = []
prev_pull_lists = []
print("\nThe following order and pull lists will be used to exclude titles that have already been ordered or previously pulled:\n")
for file in sem_file_path:
    if file.startswith("order_list"): 
        print(file)
        prev_ord_lists.append(file)
    elif file.startswith("pull_list"):
        print(file)
        prev_pull_lists.append(file)
print("\nNote: If this list is empty, no previous order lists for the " + sem_folder + " semester were found.")
print("If this is an error, make sure the previous order list exists, and verify it is named correctly (order_list [date]).\n")

# Ask if user would like to be notified when script is done running. By default, the notification will assume no.
# If the user selects yes, the script will automatically open the file folder of the exported pull/order lists.
# notify_user = input("Would you like to be notified when the script is done running? The script will automatically open to the location of the order and pull lists upon completion. \nType yes or no below.\n").lower()
# if notify_user != ("yes" or "no"):  notify_user == "no"

# Extract the date from the filename entered above - this will be used to later name the order and pull list files, so that they match the bookstore list date
bkstr_file_date = re.sub(" ", "", ((re.search('\s(.*)', bkstr_file_name)).group()))
bkstr_file_with_path = "G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\semesters\\" + sem_folder + "\\" + bkstr_file_name + ".xlsx"

# **********************************
# Handling Previously Ordered Titles
# **********************************

# Retrieve ISBNs from previous order lists, if they exist. These will be excluded from the bookstore list.
prev_ord_isbns = []

for file in prev_ord_lists:
    temp_ord_df = pd.read_excel("G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\semesters\\" + sem_folder + "\\" + file, sheet_name='Order List').fillna('')
    prev_ord_isbns.extend(temp_ord_df["ISBN-13"])

# Dedupe the ISBNs gathered from previous order lists.
prev_ord_isbns = list(set(prev_ord_isbns))
prev_ord_df = pd.DataFrame({'Previously Ordered ISBNs': prev_ord_isbns}).astype(str)

# Create a dataframe from the "For database process" tab of the bookstore list.
tb_df = pd.read_excel(bkstr_file_with_path, sheet_name='formatted for DB processing').fillna('')

# Remove ISBN matches of previous order lists from the current bookstore list.
tb_df = tb_df[~tb_df["ISBN-13"].isin(prev_ord_isbns)]

# *********************************************************
# Handling Special Titles (Excluded and Replacement Titles)
# *********************************************************

# Create three dataframes from the Special Titles spreadsheet: replaced titles, excluded titles, and titles with replacements that have no ISBN matches in the catalog (these get matched via catkey).
sptitles_replace_df = pd.read_excel('G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\SpecialTitles.xlsx', sheet_name='replace', dtype=str).fillna('')
sptitles_exclude_df = pd.read_excel('G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\SpecialTitles.xlsx', sheet_name='exclude', dtype=str).fillna('')

# Create a list of deduped, unique ISBN-13s from the bookstore data.
unique_isbns = list(set(tb_df["ISBN-13"]))

# Add leading zeros back to entries in replaced/excluded lists
sptitles_replace_df["Bookstore ISBN"] = [str(isbn).zfill(10) if len(isbn) < 10 else isbn for isbn in sptitles_replace_df["Bookstore ISBN"]]
sptitles_replace_df["Catalog ISBN"] = [str(isbn).zfill(10) if len(isbn) < 10 else isbn for isbn in sptitles_replace_df["Catalog ISBN"]]

# Create a list of key-value pairs for the Special Titles "replace" spreadsheet. Bookstore ISBNs are paired with Catalog ISBNs for replacements.
replacement_isbns = dict(zip(sptitles_replace_df["Bookstore ISBN"], sptitles_replace_df["Catalog ISBN"]))
# Create a list of ISBNs to exclude from the bookstore list.
excluded_isbns = list(sptitles_exclude_df["Bookstore ISBN"])
# Create a list of key-value pairs for the replace without ISBN matches spreadsheet.

# Handling replacement titles.
bkstr_isbn_to_replace = []
final_replacement_isbn = []
count = 0

# For every key in the list of replacement ISBNs, see if there's a match in the list of deduped ISBNs from the bookstore list. If there's a match, save these matches to separate lists.
# Then, remove those matched ISBNs from the list of deduped bookstore ISBNs. This prevents the script from attempting to add ISBNs we have replacements for to the order list.
for key in replacement_isbns.keys():
    for isbn in unique_isbns:
        if isbn != '':
            if int(isbn) == int(key):
                bkstr_isbn_to_replace.append(key)
                final_replacement_isbn.append(int(replacement_isbns[key]))
                unique_isbns.remove(isbn)

# Create a dataframe of the current bookstore list's ISBNs to replace, and the replaced ISBNs. This will later become the "Replaced ISBNs" tab on the order list spreadsheet.
# We'll also use the final_replacement_isbn list and add those ISBNs to the pull list, and exclude them from the order list. 
final_replacements_df = pd.DataFrame({'Bookstore ISBNs to Replace': bkstr_isbn_to_replace,
                                      'Replacement ISBNs': final_replacement_isbn})
final_replacements_df['Bookstore ISBNs to Replace'] = final_replacements_df['Bookstore ISBNs to Replace'].astype(str)
final_replacements_df['Replacement ISBNs'] = final_replacements_df['Replacement ISBNs'].astype(str)
final_replacements_df["Replacement ISBNs"] = [str(isbn).zfill(10) if len(isbn) < 10 else isbn for isbn in final_replacements_df["Replacement ISBNs"]]

# *********************
# Searching the Catalog
# *********************

# Handling excluded titles. This creates a list of all matches between the deduped ISBNs and a list of ISBNs to exclude. 
excl_matches = []
for isbn in excluded_isbns: 
    if float(isbn) in unique_isbns: 
        excl_matches.append(isbn)
        unique_isbns.remove(float(isbn))

# Create a dataframe of the bookstore items we can exclude.
final_exclusions_df = pd.DataFrame({'Bookstore ISBNs to Exclude': excl_matches})

catalog_url = "https://catalog.lib.ncsu.edu/"    # This URL string prepends the information needed to search an item by ISBN.
catkeys = []                                    # Contains a list of catkeys scraped from searching by ISBN.
isbns_not_found = []                            # Contains a list of ISBNs not found in the search results.
bookstore_isbns_in_catalog = []                            # Contains a list of ISBNs that were found, and are captured in the bookstore data. 
isbn_errors = 0                                 # Bookstore ISBNs that = 0 or are blank spaces. We'll count these up to save for later.
                                                # Specificed as bookstore ISBNs here, to differentiate from the list of all possible ISBNs found in the item's catalog entry.

# Add the replaced ISBNs to the list of deduped ISBNs from the bookstore. These will eventually make it to the pull list.
unique_isbns.extend(final_replacements_df['Replacement ISBNs'])

# WIP for the next two: turn these for loops into a function, since the same process is repeated twice.
for key in replacement_isbns.keys():
    for isbn in unique_isbns:
        if isbn != '':
            if int(isbn) == int(key):
                bkstr_isbn_to_replace.append(key)
                final_replacement_isbn.append(int(replacement_isbns[key]))
                unique_isbns.remove(isbn)

# Status update messages. This script currently takes around 40 minutes to run on 1,200 items, so having status updates is helpful to note where the script is at in processing.
print("\nStarting up...")
print("\n********************************************************************************************\nProcessing " + str(len(unique_isbns)) + " ISBNs.")
print(str(len(excl_matches)) + " ISBNs will be excluded. Check 'Excluded Titles' tab on the order list for details.")
print(str(len(prev_ord_isbns)) + " ISBNs were previously ordered. Check 'Previously Ordered' tab on the order list for details.")
print("**********************************************************************************************\n")

# Detailed overview:
# For every ISBN in the list of unique ISBNs, create a URL from the ISBN to search the catalog.
# Search the HTML of the search result page. If the text "<a data-context-href="/catalog/" exists, there's a matching result.
# Retrieve the catkey of the top result from the page. This gets saved to the catkeys list.
# If the HTML text doesn't exist, there's no results found and the ISBN is added to the "ISBNs Not Found" list.

count = 0       # Counter variable for indicating where the script is at in the list.

# For every ISBN in the list of deduped ISBNs, print out it's index in the list. This serves as a status update message.
for isbn in unique_isbns:
    print("Processing item #" + str(unique_isbns.index(isbn)+1) + "/" + str(len(unique_isbns)))
    # If the ISBN from the bookstore is blank or 0, then it's an error. We save these to a separate list to display to the user later on.
    if(isbn == "" or isbn == 0 or isbn == "0"):
        isbn_errors += 1
    # If an ISBN is valid (not blank or 0), retrieve the HTML of the catalog search page associated with that ISBN.
    else: 
        isbn_search_url = catalog_url + "?search_field=all_fields&q=" + str(isbn)
        r = requests.get(isbn_search_url)
        page_text = r.text
        # If an ISBN is found, retrieve the catkey from the page results. These will be used for the pull list.
        if('<a data-context-href="/catalog/') in page_text: 
            catkey = re.findall('catalog/(.*)counter=1', page_text)
            catkey = re.sub('/track[?]', '', str(catkey))
            catkey = re.sub('NCSU', '', catkey)
            catkey = re.sub('[^A-Za-z0-9]', '', str(catkey))
            catkeys.append(catkey)
            bookstore_isbns_in_catalog.append(int(isbn))
            count = count+1
        # If no ISBN is found, we don't own it. Add that ISBN to the list of ISBNs not found. These will be used for the order list.
        else:
            isbns_not_found.append(int(isbn))

# More status updates. Produces the number of items found in the catalog and the number of items not found in the catalog.
print("\n******************************\n" + str(count) + " items found in catalog.")
print(str(len(isbns_not_found)) + " items not found in catalog.\n******************************\n")

# ***********************
# Creating the Order List
# ***********************

print("Creating order list...")

# Creates the order list. This part is really straightforward: just match the ISBNs Not Found back to the ISBNs in the bookstore list, and export it as a spreadsheet.
# Create a new dataframe that matches on ISBN-13s. If the ISBN-13 in the bookstore list is found in the list of ISBNs not found, save that information to this new dataframe.
order_df = tb_df[tb_df['ISBN-13'].isin(isbns_not_found)]
# Since the order list dataframe is a direct copy of the bookstore dataframe, we need to drop the irrelevant columns.if {'Dept', 'Crs', 'Sect', 'Est Enr', 'Stat', 'ISBN-10', 'Last Used', 'List_New', 'Net_New', 'List_Used', 'Net_Used', 'Copyright', 'Date', 'Instructor'}.issubset(tb_df.columns)
order_df = order_df.drop(['Dept', 'Crs', 'Sect', 'Est Enr', 'Stat', 'ISBN-10', 'Last Used', 'List_New', 'Net_New', 'List_Used', 'Net_Used', 'Copyright', 'Date', 'Instructor', 'Book Status'], axis=1, errors='ignore')
order_df = order_df.drop_duplicates(keep='first', subset='ISBN-13')
# Not sure why, but for whatever reason, ISBNs first need to be converted to int64, and then saved as a string. It's a little clunky, but it works.
order_df["ISBN-13"] = order_df["ISBN-13"].astype('int64')
order_df["ISBN-13"] = order_df["ISBN-13"].astype(str)

# This merges the department, course, and section information into one column, pulling from dept/crs/sect of the original bookstore list. 
# The bookstore list has each section of a course listed separately, with repeating ISBNs. If multiple sections use the same book, it's confusing to display.
# This below uses ISBN as the key, and flattens all results that match that ISBN. That way, multiple sections of the same course are matched to only one ISBN. It displays per ISBN rather than per course.
dept_crs_sect = [] 
for isbn in order_df["ISBN-13"]:
    filtered_isbns_tb_df = tb_df[tb_df['ISBN-13'] == int(isbn)]
    filtered_isbns_tb_df['Course Info'] = filtered_isbns_tb_df['Dept'] + " " + filtered_isbns_tb_df['Crs'].astype(str) + "-" + filtered_isbns_tb_df['Sect'].astype(str)
    course_info = list(filtered_isbns_tb_df['Course Info'])
    dept_crs_sect.append('\n'.join(course_info))
order_df.insert(loc=1, column='Dept Crs-Sect', value=dept_crs_sect)

# Filepath for output
print("\nExporting order list to " + sem_folder + " folder...")

# Reorder the dataframe if things got out of place.
order_df = order_df.reindex(columns=['Term', 'Dept Crs-Sect', 'Author', 'Binding', 'Title', 'ISBN-13', 'Edition'])

# Export the dataframe to an Excel spreadsheet. The sheet is named Order List, and the cells are formatted to have text wrapping and specific column widths for readability purposes.
with pd.ExcelWriter("G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\semesters\\" + sem_folder + "\\" + 'order_list ' + bkstr_file_date + '.xlsx', engine='xlsxwriter') as writer:
    # Current Order List
    order_df.to_excel(writer, sheet_name='Order List', index=False, na_rep='Nan')
    workbook = writer.book
    worksheet = writer.sheets['Order List']
    # This cell_format is applied to all of the cells in the spreadsheet(s): allow for text wrapping and vertically align the text to the top of each cell.
    cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    # Order List
    worksheet.set_column('A:A', 5, cell_format=cell_format)     # term
    worksheet.set_column('B:B', 15, cell_format=cell_format)    # dept crs-sect
    worksheet.set_column('C:C', 30, cell_format=cell_format)    # author
    worksheet.set_column('D:D', 15, cell_format=cell_format)    # cover
    worksheet.set_column('E:E', 60, cell_format=cell_format)    # title
    worksheet.set_column('F:F', 15, cell_format=cell_format)    # isbn-13
    worksheet.set_column('G:G', 15, cell_format=cell_format)    # edition
    # Replaced Titles
    final_replacements_df.to_excel(writer, sheet_name='Replaced Titles', index=False, na_rep='Nan')
    worksheet = writer.sheets['Replaced Titles']
    worksheet.set_column('A:A', 25, cell_format=cell_format)    # bookstore isbn
    worksheet.set_column('B:B', 25, cell_format=cell_format)    # catalog isbn to replace bookstore isbn with
    # Excluded Titles
    final_exclusions_df.to_excel(writer, sheet_name='Excluded Titles', index=False, na_rep='Nan')
    worksheet = writer.sheets['Excluded Titles']                
    worksheet.set_column('A:A', 25, cell_format=cell_format)    # bookstore isbn excluded from list
    # Previously Ordered Titles
    prev_ord_df.to_excel(writer, sheet_name='Previously Ordered Titles', index=False, na_rep='Nan')
    worksheet = writer.sheets['Previously Ordered Titles']
    worksheet.set_column('A:A', 30, cell_format=cell_format)    # previously ordered ISBN
    worksheet.set_column('B:B', 50, cell_format=cell_format)    # previously ordered title
    worksheet.set_column('C:C', 15, cell_format=cell_format)    # previously ordered author

# **********************
# Creating the Pull List
# **********************

# Now for the pull list. Create empty lists for each column of the pull list. These columns will be populated using the catalog API.
all_isbns, titles, authors, editions, pub_years, item_locations, types, call_numbers, restriction_notes, barcode = [], [], [], [], [], [], [], [], [], []

# Status update that the catalog API is being scraped.
print("\nScanning the catalog for bibliographic information for pull list...\n")
# These lists are a little different from the ones above, since they'll all be merged into dept_crs_sect. Separating them here is just for readability to understand the code 
# a little better.
dept_crs_sect, dept, crs, sect = [], [], [], []         

# This loop merges department/course/section information into one cell per catalog item. This essentially does a "one to many" match - one textbook, multiple course matches.
# For every ISBN in the list of ISBNs identified in the catalog, match these ISBNs back to what's in the bookstore list to retrieve the course/dept/sect for that ISBN.
# For ISBNs in the bookstore list that had adequate replacements in the catalog (identified via Special Titles spreadsheet), these ISBNs will need to be temporarily replaced with the bookstore ISBN.
# Otherwise, there's no way to match the bookstore's dept/crs/sect information to the new ISBN. So, we go back to our key-value pairs to retrieve the original ISBN for every replacement ISBN.
# These original ISBNs are what's used to get the bookstore data for dept/crs/sect info. All of this later gets sent to the pull list with the original and replaced ISBN displayed.
bookstore_isbns_in_catalog = [str(isbn).zfill(10) if len(str(isbn)) < 10 else str(isbn) for isbn in bookstore_isbns_in_catalog]
for isbn in bookstore_isbns_in_catalog:
    # Note: Write a function for the next two if-elif statements.
    if isbn in replacement_isbns.values():
        key_list = list(replacement_isbns.keys())
        val_list = list(replacement_isbns.values())
        old_isbn = key_list[val_list.index(isbn)]
        bookstore_isbns_in_catalog[bookstore_isbns_in_catalog.index(isbn)] = old_isbn
        isbn = old_isbn
    filtered_isbns_tb_df = tb_df[tb_df['ISBN-13'] == int(isbn)]
    filtered_isbns_tb_df['Course Info'] = filtered_isbns_tb_df['Dept'] + " " + filtered_isbns_tb_df['Crs'].astype(str) + "-" + filtered_isbns_tb_df['Sect'].astype(str)
    course_info = list(filtered_isbns_tb_df['Course Info'])
    dept_crs_sect.append('\n'.join(course_info))

# Create lists for item locations and barcodes.
locations_list = []
bcodes = []
# Counter for status updates.
count = 0

# This function handles cases where the JSON for a certain item is missing.
def handle_missing_json(element, json_var):
    if element in json_var:
        return(json_var[element])
    else:
        return ""

# Compiled list of previously pulled catkeys.    
prev_pull_catkeys = []

for file in prev_pull_lists:
    temp_pull_df = pd.read_excel("G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\semesters\\" + sem_folder + "\\" + file, sheet_name='Pull List').fillna('')
    prev_pull_catkeys.extend(temp_pull_df["Catkey"])

catkeys = [key for key in catkeys if key not in prev_pull_catkeys]

# Detailed overview:
# For every catkey in the list of catkeys we scraped from the catalog earlier, scrape the JSON for that catkey.
# Capture all of the needed information for the pull list, and save each field into its own list.
# Give a status update at the end that shows what number in the list we're currently processing.
for catkey in catkeys:
    item_url = catalog_url + "/catalog/NCSU" + str(catkey) + ".json"
    r = requests.get(item_url)
    if r.status_code != 204:
        item_json = r.json()
        isbns = '\n'.join(handle_missing_json("isbn", item_json))
        all_isbns.append(isbns)
        titles.append(handle_missing_json("title", item_json))
        authors.append(handle_missing_json("statement_of_responsibility", item_json))
        editions.append(handle_missing_json("edition", item_json))
        pub_years.append(handle_missing_json("publication_year", item_json))
        locations = handle_missing_json("locations", item_json)
        # Handles multiple locations of a single catalog item.
        for loc in locations:
            locations_list.append(loc['library']['display'])
            locations_list.append("-")
            locations_list.append(loc['location']['display'])
            locations_list.append("\n")
        item_locations.append(' '.join(locations_list))
        locations_list = []
        call_num = handle_missing_json("call_number", item_json)
        call_numbers.append(call_num)
        # Handles ebook vs physical titles.
        if "ebook" in call_num:
            types.append("eBook")
        else:
            type = re.sub('[^A-Za-z0-9\s]', '', str(handle_missing_json("type", item_json)))
            types.append(type)
        item_info = handle_missing_json("items", item_json)
        # Handles multiple barcodes for a single catalog item.
        if item_info != "":
            for elem in item_info:
                bcode_list = re.findall("'item_id': '(.*)', 'loc", str(elem))
                for code in bcode_list:
                    bcode = re.sub(", '(.*)", '', str(code))
                    bcode = re.sub('[^A-Za-z0-9\s]', '', str(bcode))
                    bcodes.append(bcode)
            barcode_list = '\n'.join([str(item) for item in bcodes])
            bcodes=[]
            barcode.append(barcode_list)
        else:
            barcode.append("")
        restriction_notes.append(handle_missing_json("access_restrictions", item_json))
        count = count+1
    print("Formatting item " + str(count) + "/" + str(len(catkeys)))

# Create a spreadsheet from the data and export that data to the user's desktop. 
print("\nExporting pull list to " + sem_folder + " folder...")
catkey_isbn_df = pd.DataFrame(list(zip(dept_crs_sect, catkeys, titles, authors, item_locations, call_numbers, types, bookstore_isbns_in_catalog, all_isbns, editions, pub_years, barcode, restriction_notes)),
                              columns=['Dept Crs-Sect', 'Catkey', 'Title', 'Author', 'Item Location', 'Call Number', 'Item Type', 'Bookstore ISBN', 'All ISBNs', 'Edition', 'Year', 'Barcodes', 'Access Restrictions'])
catkey_isbn_df["Bookstore ISBN"] = catkey_isbn_df["Bookstore ISBN"].astype('int64')
catkey_isbn_df["Bookstore ISBN"] = catkey_isbn_df["Bookstore ISBN"].astype(str)

# Just like with the order list, format the pull list into an Excel spreadsheet.
with pd.ExcelWriter("G:\Acquisitions & Discovery\Data Projects & Partnerships Unit\Textbooks\semesters\\" + sem_folder + "\\" + 'pull_list ' + bkstr_file_date + '.xlsx', engine='xlsxwriter') as writer:
    catkey_isbn_df.to_excel(writer, sheet_name='Pull List', index=False, na_rep='Nan')
    workbook = writer.book
    worksheet = writer.sheets['Pull List']
    # Formatting columns to have top vertical alignment, text wrapping, and appropriate width.
    cell_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
    worksheet.set_column('A:A', 12, cell_format=cell_format)        # dept course-sect
    worksheet.set_column('B:B', 10, cell_format=cell_format)        # catkeys
    worksheet.set_column('C:D', 45, cell_format=cell_format)        # title and author
    worksheet.set_column('E:E', 35, cell_format=cell_format)        # item location
    worksheet.set_column('F:F', 20, cell_format=cell_format)        # call number
    worksheet.set_column('G:G', 10, cell_format=cell_format)        # item type
    worksheet.set_column('H:H', 15, cell_format=cell_format)        # bookstore ISBN
    worksheet.set_column('I:I', 15, cell_format=cell_format)        # all ISBNs
    worksheet.set_column('J:J', 10, cell_format=cell_format)        # edition
    worksheet.set_column('K:K', 5, cell_format=cell_format)         # year published
    worksheet.set_column('L:L', 15, cell_format=cell_format)        # barcodes
    worksheet.set_column('M:M', 20, cell_format=cell_format)        # access restrictions

# *******************************************
# Closing: Time Elapsed and File Notification
# *******************************************

# Gives time elapsed since the script first started running.
end = time.time()
time_elapsed = round((end-start)/60)
print("\nExecution time: " + str(time_elapsed) + " minutes.")
print(str(isbn_errors) + " ISBN(s) error from the original bookstore list.")
print("\nIf there are any ISBNs caught in error, check the bookstore list for any empty or invalid ISBNs.")

# Opens the folder location of the exported order/pull lists if user requested to be notified when the script is done.
# if notify_user == "yes":
#     path = sem_file_path
#     path = os.path.realpath(path)
#     os.startfile(path)