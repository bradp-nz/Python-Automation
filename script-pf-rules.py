import pandas as pd

# Define the path to the Excel file and the sheet name
excel_file_path = '<excel_sheet_path>'
sheet_name = 'Export'
ip_column_name = 'Destination IP'
server_name_column_name = 'Destination server name'

# Load the Excel data
df = pd.read_excel(excel_file_path, sheet_name=sheet_name)

# Extract IP addresses and server names (assuming they are in columns named 'IP' and 'ServerName')
ip_addresses = df[ip_column_name].tolist()
server_names = df[server_name_column_name].tolist()

# Print extracted IP addresses and server names (for verification)
print("Extracted IP addresses:")
for ip in ip_addresses:
    print(ip)

print("\nExtracted Server Names:")
for server in server_names:
    print(server)

# Define the criteria to remove rows
# Example criteria, you can customize these lists as needed
ips_to_remove = ['0.0.0.0', '0.0.0.0' ]
server_names_to_remove = ['servername1', 'servername2']

# Remove rows with matching IP addresses or server names
df = df[~df[ip_column_name].isin(ips_to_remove) & ~df[server_name_column_name].isin(server_names_to_remove)]

# Write the modified DataFrame back to the Excel file
with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name=sheet_name, index=False)

print("Rows containing specified IP addresses or server names have been removed.")
