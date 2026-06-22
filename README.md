# 🇲🇦 Morocco European Solidarity Corps (ESC) Opportunities Dashboard 🇪🇺

A premium, interactive web dashboard and automated scraper that fetches, filters, and formats European Youth Portal volunteer opportunities open to Moroccan residents.

This project completely automates the tracking of ESC opportunities, converting raw, hard-to-navigate API data into a clean, modern, and searchable interface.

---

## ✨ Features

- **🚀 Live Subprocess Sync**: Instantly update opportunities by clicking **"تحديث البيانات"** on the dashboard. It runs the scraper script via Flask SSE (Server-Sent Events) and streams live console logs to the webpage.
- **🎨 Premium Responsive UI/UX**: Designed with a clean glassmorphism theme, custom dark/light modes, Cairo (Arabic) and Outfit (English) fonts, and smooth CSS micro-animations.
- **🔍 Advanced Real-Time Filters**:
  - Search by title, city, organization, or topics.
  - Filter by Country (with flags) and Topics.
  - Filter by required documents (e.g., find opportunities that **do not require CVs or Motivation Letters**).
  - Filter by deadline (All, Ongoing, Expiring soon in < 7 days).
- **🌐 In-App Arabic Translation**: English description fields can be dynamically translated to Arabic on-demand using an integrated client-side translation button.
- **📂 Multi-Format Exports**: Auto-generates structured Excel spreadsheets (`Morocco_Volunteer_Opportunities.xlsx`) and JSON files (`Morocco_Volunteer_Opportunities.json`) on each update.
- **💻 Standalone Mode**: The output HTML file (`Morocco_Volunteer_Opportunities.html`) has all opportunities data embedded as JSON and can be opened offline by double-clicking it (`file://` protocol) without needing to run any server.

---

## 🛠️ Tech Stack

- **Backend**: Python, Flask, Flask-CORS (for Server-Sent Events logging)
- **Scraper / Parser**: pandas, openpyxl, requests
- **Frontend**: HTML5, Vanilla CSS3 (custom CSS variables & themes), JavaScript (ES6+), Bootstrap 5 (RTL layout), FontAwesome 6

---

## 🚀 Getting Started

### 📋 Prerequisites

Make sure you have **Python 3.8+** installed on your system.

### 🔧 Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/bilalelou/morocco-esc-dashboard.git
   cd morocco-esc-dashboard
   ```

2. **Install the dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Run the local server**:
   ```bash
   python server.py
   ```

4. **Access the dashboard**:
   Open your browser and navigate to **`http://localhost:5000`**.

---

## 📁 Project Structure

- `server.py`: Flask local server that serves the HTML page and streams live scraping console logs.
- `scrape_europa_opportunities.py`: Core python scraper that queries the European Youth Portal API, filters for Moroccan residents, and exports data.
- `Morocco_Volunteer_Opportunities.html`: The generated web dashboard (standalone dynamic SPA).
- `Morocco_Volunteer_Opportunities.xlsx`: Excel output for structured offline review.
- `Morocco_Volunteer_Opportunities.json`: JSON output containing the raw database.
- `requirements.txt`: Python dependencies list.
- `README.md`: Project documentation.

---

## 📸 Dashboard Preview

*Dashboard includes Light/Dark mode toggling, real-time filters sidebar, statistics count, card grids, pagination, and modal details with built-in instant translation.*

---

## 🤝 Contributing

Contributions, issues, and feature requests are welcome! Feel free to check the [issues page](https://github.com/bilalelou/morocco-esc-dashboard/issues).

---

## 📝 License

Distributed under the MIT License. See `LICENSE` for more information.
