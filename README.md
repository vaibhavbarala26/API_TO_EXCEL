# Cryptocurrency Tracker

This script fetches cryptocurrency data from the CoinMarketCap API, processes it, saves the data to Excel files, and uploads/updates the files on Google Drive. The script also calculates and displays useful statistics, such as the maximum and minimum percentage changes and the average price of the top 50 cryptocurrencies.

---

## **Features**
- Fetches live cryptocurrency data from the CoinMarketCap API.
- Saves cryptocurrency data to Excel files with detailed sorting and calculations.
- Automatically uploads or updates the files on Google Drive.
- Calculates key statistics, including maximum and minimum percentage changes and average prices.

---

## **Prerequisites**
1. Python 3.7 or higher.
2. Google Cloud Service Account credentials (`Credentials.json`).
3. `.env` file containing API keys and Google Drive folder ID.

---

## **Setup Instructions**

### **1. Clone the Repository**
Clone the project to your local system:
```bash
git clone <repository-url>
cd <project-folder>
```
##**2. Install Dependencies**
```bash
pip Install -r requirements
```

##**3. Create a .env file**
```bash
api_key="12044ab5-bd04-4c17-88ec-bdeffe7a61ae"
parent_folder=YOUR_GOOGLE_DRIVE_FOLDER_ID
```

##**4.Add Google Service Account Credential**
Place your Google Cloud Credentials.json file in the project root directory.


## How to Run

**1.Navigate to the folder**
```bash
cd <project-folder>
```
**2.Run the Script**

```bash
python <script-name>.python
```

## Dependencies

requests==2.31.0
pandas==1.5.3
openpyxl==3.1.2
google-auth==2.23.0
google-api-python-client==2.94.0
python-dotenv==1.0.0
