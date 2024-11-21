# Video Duration Mapper

This script generates a detailed report of video durations found within a specified folder and its subdirectories. The report aggregates video durations by folder, calculates total and average durations, and formats the results in an Excel file with professional formatting.

---

## Features

- Scans a folder for video files (e.g., `.mp4`, `.mkv`, `.avi`).
- Calculates individual, total, and average durations.
- Allows exclusion of specific folders from the scan.
- Supports hierarchical folder mapping up to specified levels.
- Generates an Excel report with formatted cells, merged ranges, and easy readability.
- Outputs total and average durations in a human-readable `HH:MM:SS` format.

---

## Requirements

- **Python version**: Python 3.7+
- **Required packages**:
  - `pandas`
  - `moviepy`
  - `openpyxl`

Install dependencies using:
```
bash
pip install pandas moviepy openpyxl
```

---

## Usage Instructions

### 1. Run the Script
Execute the script via terminal or an IDE:
```
bash
python video_mapping.py
```

### 2. Input Prompts
When you run the script, you will be asked to provide the following inputs:

- **Root Folder Path**  
  The directory where your videos are stored. The script will scan this folder and its subfolders.

- **Subfolder Level**  
  Specify how deep the script should scan.  
  - `0`: Only the root folder.  
  - `1`: Root and first-level subfolders. 
  - `N`: etc. 

- **Exclude Paths**  
  List of folders to exclude from the scan, separated by commas (e.g., `path1,path2`). Leave blank if none.

- **Export Report**  
  Optionally save the results in an Excel file. The report is saved in the root folder as `video_report.xlsx`.

---

## Error Handling

- **Unsupported Formats**  
  Videos with unsupported formats or corrupt files will be skipped. The script will display an error message but continue processing other videos.

- **Missing Dependencies**  
  If a required package is not installed, install it using `pip`. For example:
```
bash
pip install moviepy
```

- **File Write Issues**  
  Ensure you have permission to write files in the specified directory. If you encounter errors, try running the script with administrator privileges.

---

## Contributing

Contributions are welcome! Follow these steps to contribute:
1. Fork this repository.
2. Create a new branch for your feature or fix: `git checkout -b feature-name`.
3. Commit your changes: `git commit -m 'Add feature/fix issue'`.
4. Push to your branch: `git push origin feature-name`.
5. Submit a pull request for review.
