import re
import pandas as pd

import logging

import gspread
from gspread_dataframe import get_as_dataframe
from google.oauth2.service_account import Credentials


# Redirect print statements to logging
def log_print(*args, **kwargs):
    logging.info(" ".join(map(str, args)))   
    

# Set up logging to log all print statements into a text file
logging.basicConfig(filename="output.log", level=logging.INFO, format="%(asctime)s - %(message)s") 

# Replace the default print function
print = log_print

class IDGenerator:
    
    def __init__(self):
        
        # All of the below values are explained in self.read_files
        self.fInput = None
        self.fID_keys = None
        self.fInput_keys = None
        self.dfID = None
        self.dfInput = None
        
        self.fID_URL = None
        self.fInput_filename = None
        self.fOutput_filename = None
        
        # If searching is run multiple times this variable prevents cleaning ID file multiple times
        self.fID_changed = True
        
        self.scopes = [
            "https://www.googleapis.com/auth/spreadsheets"
        ]
        self.credentials = Credentials.from_service_account_file("credentials.json", scopes=self.scopes)
        self.client = gspread.authorize(self.credentials)
    
    def toggle_fID_changed(self):
        self.fID_changed = True
      
    # Open and read files, returns values to be handled by GUI
    def read_files(self, fID_URL, fInput_filename, fOutput_filename):
        # Initialize the values to give acces to them for entire class
        if self.fID_URL != fID_URL:
            self.fID_URL = fID_URL
            self.fID_changed = True
        else:
            self.fID_changed = False
        self.fInput_filename = fInput_filename
        self.fOutput_filename = fOutput_filename
        
        # Read excel files with ID data and file to format and generate IDs
        if self.fID_changed is True:
            try:
                spreadsheet = self.client.open_by_url(fID_URL)  # Open google spreadsheet
                worksheet = spreadsheet.worksheet("Raw Date")   # Select Raw date sheet
                self.dfID = get_as_dataframe(worksheet, evaluate_formulas=True)  # Convert to dataframe
            except Exception as e:
                print(f"Cannot open file: {e}")
                return e
            
        try:
            self.fInput = pd.read_excel(fInput_filename, sheet_name=None)
        except Exception as e:
            print(f"Cannot open file: {e}")
            return e
        
        # Sheets with relevant data
        self.fInput_keys = list(self.fInput.keys())
        self.dfInput = self.fInput[self.fInput_keys[0]]
        
        return None
        
        
    # Makes the job done, returns values to be handled by GUI
    # This funtion runs as a thread
    def process_files(self, result_queue):
        try:
            if self.fID_changed is True:
                self.dfID = self.clean_ID(self.dfID)
        except Exception as e:
            result_queue.put(e)
            return 
        
        # Save the cleaned data to a new CSV file (optional)
        try:
            if self.fID_changed is True:
                self.dfID.to_csv('outputID-fromURL.csv', index=False)
        except Exception as e:
            print("Cannot save to file")
            result_queue.put(e)
            return 
        
        self.dfInput, matches_found = self.match_ID(self.dfInput, self.dfID)
        
        # Save data with added found IDs
        try:
            # Regex pattern to extract text before the extension .xls, xlsx
            pattern = r"^(.*?)(?=\.\w{3,4}$)"
            new_name = re.match(pattern, self.fInput_filename).group(1)
            new_name = new_name.rsplit('/', 1)[-1]  # extract everything after the last '/'
            output_name = self.fOutput_filename + "/" + new_name + "_znalezione.xlsx" # name with correct extension in specified output folder
            self.dfInput.to_excel(output_name, index=False)
        except Exception as e:
            print(f"Cannot save to file: {e}")
            result_queue.put(e)
            return
        
        result_queue.put(matches_found)
        return
            
     
    # Finds matches for one laptop in dfID
    # Returns df with with all the matches
    def match_one(self, input_row, dfID):
        # Structure od output df witch matched ID
        df_index = ['ID', 'count', 'manufacturer', 'model', 'processor', 'ram', 'hdd', 'gpu', 'resolution', 'touchscreen', 'windows', 'lap_class']
        df_matched = pd.DataFrame(columns=df_index)
        
        # Iterate dfID file by rows
        for index, row in dfID.iterrows():
            # temporary match variables. 0 if not matched, 1 if matched
            m_id = 0
            count = 0
            m_manufacturer = False
            m_model = False
            m_processor = False
            m_ram = False
            m_hdd = False
            m_gpu = False
            m_resolution = False
            m_touchscreen = False
            m_windows = False
            m_lap_class = False
            
            # For debugging purposes
            tID = int(row['ID'])
            tIndex = int(input_row['Lp.'])
            
            # Following if's search for matches for given laptop
            # Count matches and at the end return df row with boolean values of matched values
            if input_row['Producent'].lower() in row['Manufacturer'].lower():
                m_manufacturer = True
                count += 1
                    
            if input_row['Model'].lower() in row['Model'].lower():  # This is not ideal because some models wrongly match ex. T480 and T480s and silver versions of laptops
                m_model = True
                count += 1
            
            if input_row['Procesor'].lower() in row['Processor'].lower():
                m_processor = True
                count += 1
            
            # If Manufacturer, Model and Processor are not matched there is no point of checking other things    
            if count < 3:
                continue
            
            # Ignore all ID's that have specific HDD value. All ID's should only use 'BRAK DYSKU'  for reasons specific to company policy
            # These ID's could have been removed during cleaning process in self.clean_ID but it was left for legacy
            if 'BRAK DYSKU' not in row['HDD']:
                continue
            else: 
                m_hdd = True    # if there is NO HDD in ID then there is a match no matter what
                count += 1
                
            # Extract RAM and HDD values from input data
            matches = re.findall(r'\d+', input_row['Docelowa'])
            ram1 = matches[0]
            hdd1 = matches[1]
            
            # Extract RAM values from string containing all sort of additional info
            ram2 = re.search(r'\d+', row['RAM'])
            if ram2:
                ram2 = ram2.group(0)
            if ram1 == ram2:
                m_ram = True
                count += 1
            
            # Extract HDD values from string containing all sort of additional info
            # Matching HDD values is legacy code because all actively used ID's have 'BRAK DYSKU' value for reasons specific to company policy
            
            # hdd2 = re.search(r'\d+', row['HDD'])
            # if hdd2:
            #     hdd2 = hdd2.group(0)
            # if hdd1 == hdd2:
            #     m_hdd = True
            #     count += 1
            # elif 'BRAK DYSKU' in row['HDD']: # if there is NO HDD in ID then there is a match no matter what
            #     m_hdd = True
            #     count += 1
                
            if input_row['Grafika'].lower() in row['Graphics'].lower():
                m_gpu = True
                count += 1
                
            if input_row['Wyświetlacz'].lower() in row['Resolution'].lower():
                m_resolution = True
                count += 1
            elif 'fhd' in input_row['Wyświetlacz'].lower() and any(keyword in row['Resolution'].lower() for keyword in ['fullhd', 'full hd']):
                m_resolution = True
                count += 1
                
            if 'dotyk' in input_row['Wyświetlacz'].lower() and 'Yes' in row['Touchscreen']:
                m_touchscreen = True
                count += 1
            elif 'dotyk' not in input_row['Wyświetlacz'].lower() and 'No' in row['Touchscreen']:
                m_touchscreen = True
                count += 1
                
            if any(keyword in input_row['Windows'].lower() for keyword in ['win11pro', 'win11p', 'w11p']) and 'w11p' in row['Windows'].lower():
                m_windows = True
                count += 1
                
            elif any(keyword in input_row['Windows'].lower() for keyword in ['win11home', 'win11h', 'w11h']) and 'w11h' in row['Windows'].lower():
                m_windows = True
                count += 1
                
            if input_row['Klasa'].lower() == row['Class'].lower():
                m_lap_class = True
                count += 1
            
            # If less than that is matched then laptop is to different to even show it    
            if(count > 6):
                # Add new row to df_matched
                new_row = pd.Series([row['ID'], count, m_manufacturer, m_model, m_processor, m_ram, m_hdd, m_gpu, m_resolution, m_touchscreen, m_windows, m_lap_class], index=df_index)
                df_matched = pd.concat([df_matched, new_row.to_frame().T], ignore_index=True)
                
        if not df_matched.empty:
            return df_matched
        else:
            return None
                
    
    # Finds best matches for all laptops             
    def match_ID(self, dfInput, dfID):
        all_matches = []
        all_matches_count = 0
        dfCleaned = dfInput.dropna(subset=['Model']).copy()
        for index, row in dfCleaned.iterrows():
            print(index, "\t", row['S/N'], end="\t")
            if(index >= 0):
                matched = self.match_one(row, dfID)
                if matched is not None:
                    max_matched = matched['count'].max()    # Best match with most matched values
                    print("matches = ", max_matched, end=",\t")
                    max_rows = matched[matched['count'] == max_matched] # All rows with the most amout of matches
                    print("count = ", len(max_rows))
                    # String composed of all ID from max_rows separated by ', '
                    ID_matched = ', '.join(map(str, max_rows['ID'].tolist()))
                    
                    # Only 9 or 10 matches are perfect matches
                    # Add info about non-perfect matches otherwise
                    if max_matched < 9:
                        ID_matched = "Najbliższe: " + ID_matched
                    
                    all_matches.append(ID_matched)
                    all_matches_count += 1
                else:
                    all_matches.append('brak')
                    print() # Print newline because ther is no newline at the end of -- print(index, row['S/N'], end=" ") --
                

        print("Found ID: ", all_matches_count)
        print("----------------------------------------------------------------------------------------\n")
        dfCleaned['Znalezione ID'] = all_matches
        return dfCleaned, all_matches_count # return final dataframe and how many matches were found
              
              
    def clean_ID(self, df):
        # Number of rows of df before any formatting
        unformatted_len = len(df)
        
        # Remove rows containing the keyword "dokująca"
        df = df[~df['Pełna nazwa'].str.contains('dokująca', case=False, na=False)]

        # Apply the function to the "Specs" column to extract values from it
        df[['Manufacturer', 'Model', 'Processor', 'RAM', 'HDD', 'Graphics', 'Resolution', 'Touchscreen', 'Windows', 'Class']] = df['Pełna nazwa'].apply(self.extract_specs)

        # Drop the original "Specs" column and other non-important columns
        df = df.drop(columns=['Producent', 'Pełna nazwa'])

        # Addidtionaly format and fix errors that couldn't get fixed by extract_specs function
        df = self.format_specs(df)

        # Drop rows that didn't pass through extract_specs and format_specs because of all sort of formatting and data errors
        df = df[~df['Processor'].str.fullmatch("-")]

        # Drop rows with blank important data for random reasons
        df = df[~df['Processor'].str.fullmatch("")]

        # Convert 'ID' column to string and cut off '.0' from it
        df["ID"] = df["ID"].astype(str)
        df["ID"] = df["ID"].str.split('.').str[0]

        formatted_len = len(df)

        print("All matched IDs: ", unformatted_len)
        print("Useful IDs: ", formatted_len, " Deleted: ", unformatted_len - formatted_len)
        print("========================================================================================\n\n\n")

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
                    laptop_class = part.replace("Klasa ", "") # Remove 'Klasa '
            
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
        print("Iterating ID: ")
        # Iterate over rows
        for index, row in df.iterrows():
            #print(index)
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
                #print(row['ID'])
                df.at[index, 'Processor'] = "-"
            
        return df 
    