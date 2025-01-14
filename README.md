# Excel Fuzzy Deduplication Tool 🔍

A Streamlit application that helps you find and group similar records in Excel files using fuzzy matching. Try it out at: [fuzzydeduplicator.streamlit.app](https://fuzzydeduplicator.streamlit.app/)

## Features

- Upload Excel files (.xlsx, .xls)
- Fuzzy matching using Jaro-Winkler similarity algorithm
- Adjustable similarity threshold
- Performance optimization through leading character matching
- Progress tracking during processing
- Excel output with duplicate groups identified
- Interactive data preview

## How It Works

1. **Data Loading**: Upload your Excel file containing the records you want to deduplicate
2. **Text Processing**: All columns for each record are combined into a single text string
3. **Smart Comparison**: Records are first grouped by leading characters to reduce comparison space
4. **Fuzzy Matching**: Each record is compared to others using the Jaro-Winkler similarity algorithm
5. **Output Generation**: Results are provided in an Excel file with duplicate groups identified

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/excel-fuzzy-dedup.git
cd excel-fuzzy-dedup
```

2. Install the required packages:
```bash
pip install -r requirements.txt
```

3. Run the app:
```bash
streamlit run app.py
```

## Usage

1. Open the app in your browser (will launch automatically when you run the command above)
2. Upload your Excel file using the file uploader
3. Adjust the similarity threshold (0.5-1.0):
   - Higher values require records to be more similar to be considered duplicates
   - Recommended starting value: 0.9
4. Set the number of leading characters to match (1-10):
   - Higher values = faster processing but might miss some duplicates
   - Recommended starting value: 3
5. Click "Find Duplicates" and wait for processing to complete
6. Download the results Excel file which includes:
   - All original data
   - `duplicate_group`: Group number for duplicate records (-1 for unique records)
   - `duplicate_rows`: List of Excel row numbers for all duplicates in the group

## Performance Optimization

The app uses two main strategies to optimize performance:

1. **Leading Character Matching**: Records are first grouped by their leading characters, dramatically reducing the number of necessary comparisons
2. **Progressive Processing**: Shows progress in real-time and estimates completion time

## Requirements

- Python 3.8+
- See requirements.txt for full list of dependencies

## Contributing

Pull requests are welcome! For major changes, please open an issue first to discuss what you would like to change.

## License

[MIT](https://choosealicense.com/licenses/mit/)