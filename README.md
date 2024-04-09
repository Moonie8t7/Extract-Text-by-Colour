Answer to https://www.reddit.com/r/GoogleAppsScript/comments/1bzukqu/google_sheets_extract_text_by_color_script/

Extract text based on the hexidecimal colour in a Google sheet.
Give the hex value of a colour you want to match, then provide it with a range (A1:D5, for example), then you give a starting column and row (you can use COLUMN(A1) and ROW(A1) for example).

It will then check through all the cells in that range and return a list of values that match the colour.
