# Waterfall Data Processing Script

## 📌 Project Overview
This project automates **temperature and channel-specific data processing** using a Python script.  
It extracts data from given text files and **updates specific sheets in an Excel file**.  
This can be used for **wireless communication testing and sensor data analysis**.

---

## 🚀 Key Features
- **Automatic Extraction of Temperature & Channel Data**  
  - Identifies **temperature (TT_x_xx) and channel (CH_x)** information from filenames  
  - Groups and organizes data based on temperature and channel  

- **Loading and Preprocessing Data from Text Files**  
  - Reads data from `[data]` section and converts it into a `pandas.DataFrame`  
  - Handles `NaN` values and filters necessary data  

- **Updating Specific Sheets in an Excel File**  
  - Loads the existing `p5.xlsx` file  
  - Updates relevant sheets like **Waterfall_ch1_25**, **Waterfall_ch7_25**, etc., based on temperature and channel values  

- **Sorting Data by Temperature (25°C, -40°C, 85°C) and Channel (1, 7)**  
  - Writes only available data to the Excel sheets  
  - Automatically records data corresponding to the specific temperature and channel  

---

