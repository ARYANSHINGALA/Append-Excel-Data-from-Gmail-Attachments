# 📊 Append Excel Data from Gmail Attachments using Power Query

## 🔍 Project Overview
This project automates the process of fetching Excel file attachments from Gmail and appending their data into a single table using **Power Query**.

Since Power Query does not directly support Gmail, **Google Apps Script** is used as a bridge to expose Gmail attachment data via a Web API.

---

## 🚀 Key Features
- Automatically reads Excel attachments from Gmail
- Converts Base64 encoded attachments into binary
- Loads Excel files dynamically into Power Query
- Appends data from multiple workbooks
- No manual download required
- Fully automated ETL workflow

---

## 🛠 Tools & Technologies Used
- Microsoft Excel
- Power Query (M Language)
- Google Apps Script
- Gmail API
- REST Web API

---

## 🧩 Architecture Flow

---

## ⚙️ Setup Instructions

### Step 1: Enable IMAP in Gmail
1. Open Gmail
2. Go to **Settings → See all settings**
3. Navigate to **Forwarding and POP/IMAP**
4. Enable **IMAP Access**
5. Save changes

---

### Step 2: Google Apps Script (Gmail Bridge)
1. Visit: https://script.google.com
2. Create a new project
3. Paste the Gmail attachment fetch script
4. Deploy as **Web App**
5. Set access to **Anyone**
6. Copy the Web App URL

---

### Step 3: Connect API in Power Query
1. Open Excel
2. Go to **Data → Get Data → From Web**
3. Paste the Web App URL
4. Load data into Power Query

---

### Step 4: Convert Base64 to Binary
Use this Power Query formula:

Binary.FromText([Binary], BinaryEncoding.Base64)
---

📂 Sample Output

Combined Excel table with appended data

Fully refreshable Power Query solution

🧠 Use Cases

Automated reporting

Finance & accounting data collection

Sales data consolidation

Email-based ETL pipelines

📄 Documentation

Detailed project explanation available in 

👤 Author

Aryan Shingala
🔗 GitHub: https://github.com/ARYANSHINGALA

🔗 LinkedIn: https://www.linkedin.com/in/aryanshingala
