# Assuming the two scripts are named script1.py and script2.py
# and are located in the same directory as this file.

# Run the first script
with open('json_to_excel.py') as file:
    exec(file.read())

# Run the second script
with open('excel_to_json.py') as file:
    exec(file.read())