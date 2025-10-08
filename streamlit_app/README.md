# 🤖 eMMA Procurement Opportunity Scraper - Streamlit Dashboard

A beautiful, feature-rich Streamlit dashboard for viewing and analyzing Maryland eMMA procurement opportunities.

## ✨ Features

### 📊 Dashboard Mode
- **Beautiful Metric Cards**: View key statistics at a glance
  - Total opportunities count
  - Urgent opportunities (≤7 days until due)
  - Open opportunities count
  - Unique agencies count

- **Interactive Visualizations**:
  - Opportunities by procurement type (horizontal bar chart)
  - Top agencies pie chart
  - Timeline of upcoming deadlines

- **Urgent Opportunities Widget**: Highlights opportunities due within 7 days

### 📋 Data Explorer
- Browse all Excel sheets (Master, Log, Archive, Refs)
- Advanced filtering:
  - Filter by status, procurement type, and agency
  - Full-text search in titles and descriptions
- Real-time record counting
- Export filtered data to Excel

### 🔍 Advanced Search
- Multi-field search across:
  - Opportunity titles
  - Issuing agencies
  - Categories
  - Additional information
- Custom field selection
- Export search results directly to Excel

### 📈 Analytics
- Category distribution charts
- Data quality score distribution
- Status breakdown analysis
- Recent activity log viewer

### 🚀 Run New Scrape
- Configure scraper settings:
  - Days ago to scrape
  - Max pages to scrape
  - Skip detail pages option
- Real-time progress tracking
- View scraper logs
- Instant download of generated Excel file

## 🎨 UI/UX Features

- **Modern Design**: Gradient cards, smooth animations, and clean typography
- **Responsive Layout**: Wide layout optimized for data viewing
- **Tab Navigation**: Easy switching between different views
- **Color-Coded Cards**: Different colors for different metric types
- **Interactive Charts**: Powered by Plotly for rich visualizations
- **Download Buttons**: Export any view or filtered dataset
- **Progress Indicators**: Visual feedback during scraping operations

## 🚀 Getting Started

### Installation

1. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

2. **Run the app**:
   ```bash
   streamlit run app.py
   ```

3. **Open in browser**:
   - The app will automatically open at `http://localhost:8501`
   - Or use the Network URL for access from other devices

### Quick Start

1. **View Existing Data**:
   - Select "📊 View Existing Data" in the sidebar
   - Navigate through Dashboard, Data Explorer, Advanced Search, and Analytics tabs
   - Use filters and search to find specific opportunities

2. **Run New Scrape**:
   - Select "🚀 Run New Scrape" in the sidebar
   - Configure scraper settings (days ago, max pages, skip details)
   - Click "Run Scraper and Generate Excel File"
   - Wait for completion and download the results

## 📁 File Structure

```
streamlit_app/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
└── README.md          # This file
```

## 🔧 Configuration

### Environment Variables

- `EMMA_XLSX`: Path to the Excel file (default: `consolidated_opportunities.xlsx`)

### Customization

You can customize the app by modifying:

- **Styling**: Edit the CSS in the `st.markdown()` section at the top of `app.py`
- **Metrics**: Modify `create_metrics_cards()` function
- **Visualizations**: Edit `create_visualizations()` function
- **Filters**: Customize `filter_dataframe()` function

## 📊 Data Requirements

The app expects an Excel file with the following sheets:

- **Master**: Main opportunities data with columns like:
  - `unique_id`
  - `solicitation_number`
  - `opportunity_title`
  - `issuing_agency`
  - `category`
  - `procurement_type`
  - `status`
  - `response_deadline`
  - `days_until_due`
  - `contact_email`
  - `additional_information`

- **Log**: Activity log (optional)
- **Archive**: Archived opportunities (optional)
- **Refs**: Reference data (optional)

## 🎯 Features Breakdown

### Caching
- Excel data is cached for 5 minutes using `@st.cache_data`
- Improves performance for repeated views

### Error Handling
- Comprehensive try-catch blocks
- User-friendly error messages
- Detailed error logs in expandable sections

### Data Quality
- Handles missing columns gracefully
- Type conversions with error handling
- Null-safe operations throughout

## 🐛 Troubleshooting

### Common Issues

1. **Excel file not found**:
   - Ensure the file path is correct in the sidebar
   - Check the `EMMA_XLSX` environment variable
   - Run a scrape to generate a new file

2. **Charts not displaying**:
   - Ensure plotly is installed: `pip install plotly>=5.0.0`
   - Check for missing data in required columns

3. **Filters not working**:
   - Verify column names match the expected format
   - Check for null values in filter columns

4. **App won't start**:
   - Verify all dependencies are installed
   - Check Python version (3.8+ required)
   - Look for import errors in the terminal

## 🔄 Updates & Maintenance

### Clearing Cache
- Click "Clear cache" in the Streamlit menu (top-right ⋮)
- Or restart the app

### Updating Dependencies
```bash
pip install -r requirements.txt --upgrade
```

## 📝 Tips & Best Practices

1. **Performance**: Use filters to reduce dataset size for faster loading
2. **Exports**: Download filtered views for focused analysis
3. **Search**: Use Advanced Search for complex queries
4. **Visualization**: Check Analytics tab for insights
5. **Monitoring**: Review the Log sheet for scraper activity

## 🤝 Contributing

Feel free to enhance this dashboard with:
- Additional visualizations
- More filter options
- Custom export formats
- Email notifications
- Scheduled scraping
- Data validation

## 📄 License

This project is part of the eMMA scraper system.

## 🙋 Support

For issues or questions:
- Check the error logs in expandable sections
- Review the scraper output logs
- Verify data file integrity
- Consult the main project documentation

---

**Developed with ❤️ for streamlining procurement tracking**
