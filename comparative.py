import pandas as pd
import matplotlib.pyplot as plt
from io import BytesIO

# Read data from the Excel files
df_1pax_1veh = pd.read_excel('aller1paxv/aller1v_data.xlsx')
df_1pax = pd.read_excel('aller/aller1_data.xlsx')

# Merge the two dataframes based on common columns
merged_df = pd.merge(df_1pax_1veh, df_1pax, on=['Date', 'Trajet', 'Company'], suffixes=('_1PAX_1VEH', '_1PAX'))

# Create an Excel writer object
output_path = 'merged_data_with_graph.xlsx'
writer = pd.ExcelWriter(output_path, engine='xlsxwriter')

# Write the merged dataframe to the Excel file
merged_df.to_excel(writer, sheet_name='Merged_Data', index=False)

# Create a comparison graph
plt.figure(figsize=(10, 6))
plt.plot(merged_df['Date'], merged_df['Price 1PAX+1VEH'], label='1 PAX + 1 VEH')
plt.plot(merged_df['Date'], merged_df['Price1 PAX'], label='1 PAX')
plt.title('Comparison of Prices: 1 PAX + 1 VEH vs 1 PAX')
plt.xlabel('Date')
plt.ylabel('Price')
plt.legend()
plt.xticks(rotation=45)
plt.tight_layout()

# Save the graph to a BytesIO buffer
graph_buffer = BytesIO()
plt.savefig(graph_buffer, format='png')
plt.close()

# Create a new worksheet in the Excel file for the graph
graph_sheet = writer.book.add_worksheet('Graph')
writer.book.add_worksheet('Data')  # Add a new worksheet for data (if needed)

# Insert the graph image from the buffer
graph_sheet.insert_image('B2', 'graph.png', {'image_data': graph_buffer})

# Close the Excel writer
writer._save()

print(f'Excel file "{output_path}" with embedded graph created successfully.')
