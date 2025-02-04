import re
import pandas as pd
import math

class IDGenerator:
    
    def __init__(self, fID_filename, fInput_filename):
        
        # Read excel files with ID data and file to format and generate IDs
        
        self.fID = pd.read_excel(fID_filename, sheet_name=None)
        self.fInput = pd.read_excel(fInput_filename, sheet_name=None)
        self.fID_keys = list(self.fID.keys())
        self.fInput_keys = list(self.fInput.keys())
        
        # Sheets with relevant data
        self.df_ID = self.fID[self.fID_keys[0]]
        self.df_Input = self.fInput[self.fInput_keys[0]]
        
        self.df_ID = self.clean_ID(self.df_ID)
        
        # Save the cleaned data to a new CSV file (optional)
        self.df_ID.to_csv('outputID.csv', index=False)
    
    
        
    def clean_ID(self, df):
        # Number of rows of df before any formatting
        unformatted_len = len(df)
        # Remove rows containing the keyword "dokująca"
        df = df[~df['Nazwa'].str.contains('dokująca', case=False, na=False)]

        # Apply the function to the "Specs" column to extract values from it
        df[['Manufacturer', 'Model', 'Processor', 'RAM', 'HDD', 'Graphics', 'Resolution', 'Touchscreen', 'Windows', 'Class']] = df['Nazwa'].apply(self.extract_specs)

        # Drop the original "Specs" column and other non-important columns
        df = df.drop(columns=['Producent', 'Nazwa'])

        # Addidtionaly format and fix errors that couldn't get fixed by extract_specs function
        df = self.format_specs(df)

        # Drop rows that didn't pass through extract_specs and format_specs because of all sort of formatting and data errors
        df = df[~df['Processor'].str.fullmatch("-")]

        # Drop rows with blank important data for random reasons
        df = df[~df['Processor'].str.fullmatch("")]

        formatted_len = len(df)

        print("Wszystkie znalezione ID: ", unformatted_len)
        print("Użyteczne ID: ", formatted_len, " Usunieto: ", unformatted_len - formatted_len)

        # Save the cleaned data to a new CSV file (optional)
        #df.to_csv('cleaned_laptops2.csv', index=False)
        
        return df    
        
    # Function to extract values from the "Specs" column
    def extract_specs(self, specs):
        parts = specs.split(' / ')
        
        # Extract Manufacturer and Model
        manufacturer = parts[0].split(' ')[1]  # Assumes format "Laptop Manufacturer Model"
        model = ' '.join(parts[0].split(' ')[2:])  # Capture the rest as the model
        
        # Extract Processor, RAM, Disk, Graphics, Resolution
        try:
            processor = parts[1].strip()  # Handle cases like "i5 - 2 generacji"   
            ram = parts[2]
            disk = parts[3]
        
            # Following columns are optional and may not always be present
            graphics = '-'  # Default value
            resolution = '-'  # Default value
            touchscreen = 'No'  # Default value
            windows_version = '-'  # Default value
            laptop_class = '-'  # Default value
            
            # Check for touchscreen, Windows version, and laptop class
            for part in parts:
                part = part.strip()
                
                # Check for touchscreen (keyword "dotyk")
                if 'dotyk' in part.lower():
                    touchscreen = 'Yes'
                
                # Check for Windows version (keywords like "W11P", "W11H", "Win11Pro", etc.)
                if any(win_keyword in part for win_keyword in ['W11P', 'W11H', 'Win11Pro', 'Win11Home']):
                    windows_version = part
                
                # Check for laptop class (keyword "Klasa")
                if 'Klasa' in part:
                    #temp_parts = part.split(' ')
                    laptop_class = part
            
            # Check for resolution (look for a part containing ")
            for part in parts:
                if '"' in part or "HD" in part or "XGA" in part:
                    resolution = part.strip()
                if any(gpu_keyword in part for gpu_keyword in ['GeForce', 'T2000', 'MX', 'RX', 'GTX', 'RTX', 'P3200', 'T1200']):
                    graphics = part.strip()
            
            return pd.Series([manufacturer, model, processor, ram, disk, graphics, resolution, touchscreen, windows_version, laptop_class])
        
        except:
            return pd.Series(["", "", "", "", "", "", "", "", "", "" ])
        
    # Additionalyy formats and fixes errors generated by extract_specs function
    # There are errors because of how poorly made is the source data excel file
    def format_specs(self, df):
        # Iterate over rows
        for index, row in df.iterrows():
            print(index)
            # Fix 'Model' being in 'Manufacturer' column because of wrong usage of ' / ' split character in source file
            if 'thinkpad' in row['Manufacturer'].lower():
                df.at[index, 'Model'] = "ThinkPad " + row['Model']
                df.at[index, 'Manufacturer'] = "Lenovo"
                
            elif 'thinkbook' in row['Manufacturer'].lower():
                df.at[index, 'Model'] = "ThinkBook " + row['Model']
                df.at[index, 'Manufacturer'] = "Lenovo"
                
            elif 'yoga' in row['Manufacturer'].lower():
                df.at[index, 'Model'] = "Yoga " + row['Model']
                df.at[index, 'Manufacturer'] = "Lenovo"
                
            elif 'probook' in row['Manufacturer'].lower():
                df.at[index, 'Model'] = "ProBook " + row['Model']
                df.at[index, 'Manufacturer'] = "HP"
                
            elif 'elitebook' in row['Manufacturer'].lower():
                df.at[index, 'Model'] = "ProBook " + row['Model']
                df.at[index, 'Manufacturer'] = "HP"
                
            # Fix column shift caused by wrong formatting of processor value in source file
            if 'generacji' in row['Model'].lower():
                parts = row['Model'].split(' ')
                model = ' '.join(parts[:-4])  # Every word excluding last 4
                processor = ' '.join(parts[-4:])  # Last 4 words
                df.at[index, 'HDD'] = df.at[index, 'RAM']
                df.at[index, 'RAM'] = df.at[index, 'Processor']
                df.at[index, 'Model'] = model
                df.at[index, 'Processor'] = processor
            
            # Every other shift error blanked    
            elif 'gb' in row['Processor'].lower():
                print(row['ID'])
                df.at[index, 'Processor'] = "-"
            
        return df
            
    def search(self, pattern, text):
        return 0
        
    def printMatches(self, all_matches):
        
        printList = ['Model', 'Procesor', 'RAM', 'Dysk', 'Grafika', 'Wyświetlacz', 'Dotyk', 'Klasa']
        
        for x in range(len(all_matches)):
            if all_matches[x] == 1:
                print(printList[x])
        
        return 0               
  
#ID_name = input("Podaj nazwe pliku z ID: ")
#Input_name = input("Podaj nazwe pliku wejsciowego: ")

ID_name = "plikID.xlsx"
Input_name = "D993.xls"

generator = IDGenerator(ID_name, Input_name)

temp = input("Press any key to continue")