# 🏫 RTE Web Dashboard & Status Checker — v3.5

A professional, real-time monitoring system for tracking RTE (Right to Education) application statuses from the government portal.

## 🚀 Key Features

### 1. High-Performance Bulk Checker (`rte_checker.py`)
- **Parallel Processing**: Uses a multi-threaded `ThreadPoolExecutor` to check hundreds of records in minutes (5x faster than sequential scripts).
- **Comprehensive Extraction**: Fetches not just the status, but also:
  - Child Name
  - Gender (`Kanya` / `Kumar`)
  - Area & Pincode
  - Mobile Number
  - Village (Gam)
- **Resume Support**: Automatically skips already `APPROVED` records and resumes from the last saved state using `RTE_Status_Results.xlsx`.
- **Live Sync Server**: Includes an embedded HTTP server (port 5001) that handles "Force Sync" requests from the web dashboard.

### 2. Interactive Live Dashboard (`dashboard.html`)
- **Real-Time Polling**: Automatically reloads data from the generated `data.js` every 3 seconds.
- **Smart Filtering**: Quickly filter by Approved, Submitted, Pending, or Rejected states.
- **WhatsApp Integration**: Single-click to send pre-filled Gujarati messages to parents using their real mobile numbers.
- **Searchable Table**: Instant search across Token No, Application ID, Name, and Mobile.

### 3. Strategic Analysis Module (`analysis.html`)
- **Gender Ratio**: Visual breakdown of Boy vs Girl distributions using Chart.js.
- **Geographical Heatmap**: Interactive Leaflet.js map showing application hotspots based on Surat area coordinates.
- **Top Clusters**: Automatic ranking of areas with the highest application density.
- **Volume Metrics**: Global view of submission trends and status metrics.

## 🛠 Setup & Usage

### Prerequisites
```powershell
pip install requests beautifulsoup4 openpyxl pandas
```

### Running the System
1. **Start the Engine**: Run the Python script to begin fetching live data.
   ```powershell
   python rte_checker.py
   ```
2. **Open the Dashboard**: Open `dashboard.html` in your web browser.
3. **Advanced Insights**: Click the **"📊 Full Analysis"** button in the dashboard header to view the strategic breakdown.

## 📁 Project Structure
- `rte_checker.py`: The core multi-threaded Python engine.
- `dashboard.html`: The main user interface.
- `analysis.html`: The advanced statistics and map page.
- `data.js`: Auto-generated database file (do not edit manually).
- `RTE_Status_Results.xlsx`: Formatted Excel report with color-coded results.

---
*Developed with focus on speed, accuracy, and professional aesthetics.*
