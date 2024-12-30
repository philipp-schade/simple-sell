# SimpleSell Formatter

SimpleSell Formatter is a Streamlit web application that allows users to upload an Excel file, format it according to specific rules, and download the newly formatted file. The application includes authentication and a custom footer.

## Features

- Upload an Excel file (`.xlsx` format)
- Format the Excel file by:
  - Deleting specific columns (B, C, D, F)
  - Center-aligning the "order_item_quantity" column
  - Formatting "order_item_sku" entries with "4S" in bold
  - Formatting "order_item_sku" entries with "-2-" in bold
  - Formatting "order_item_sku" entries with "6S" in italics and underlined
  - Highlighting rows with "order_item_quantity" > 1 in light blue
  - Drawing a bold box around rows with the same reference code
- Download the newly formatted Excel file
- Authentication mechanism
- Custom footer

## Installation

1. Clone the repository:

   ```sh
   git clone https://github.com/philipp-schade/simple-sell.git
   cd simple-sell
   
2. Create and activate a virtual environment

3. Run the app with:
    ```sh
   streamlit run app.py
    