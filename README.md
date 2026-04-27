# 🚀 Excel Admin Automation Suite

A collection of high-performance VBA utilities designed to eliminate repetitive office tasks, streamline HR operations, and automate file management directly from Microsoft Excel.

## 🛠️ Tools Included

### 1. Document Folder Automator
Instantly create complex folder hierarchies from an Excel list. 
* **Best for:** Onboarding new employees, setting up project folders, or organizing client directories.
* **Feature:** One-click generation from a root path.

### 2. File Renamer & Archive Tool
Batch rename files based on Excel mapping to standardize messy naming conventions.
* **Best for:** Cleaning up downloaded invoices, standardize photo names, or organizing document versions.
* **Feature:** Real-time status logging and error handling.

### 3. Bulk File Mover
Transfer specific files from a source directory to a destination directory based on your Excel data.
* **Best for:** Monthly archiving, moving processed documents to "Finished" folders, or sorting files across network drives.
* **Feature:** Uses high-performance FileSystemObject (FSO) for stable cross-drive transfers.



## 📂 Repository Structure

* `/src`: Contains the `.bas` source files for modular import.
* `Admin_Automation_Template.xlsm`: A ready-to-use Excel template with a professional UI.
* `README.md`: Documentation and user guide.



## 🚀 Getting Started

### Option A: Using the Template (Recommended)
1.  Download `Admin_Automation_Template.xlsm`.
2.  Open the file and **Enable Macros** when prompted.
3.  Navigate to the desired tool tab (e.g., File_Renamer).
4.  Input your file paths and filenames.
5.  Click the **Action Button** to run the automation.

### Option B: Integrating into your own Workbook
1.  Open your Excel file and press `ALT + F11` to open the VBA Editor.
2.  Go to `File > Import File...` and select the `.bas` files from the `/src` folder of this repo.
3.  Ensure your spreadsheet layout matches the expected ranges (Paths in B2/B3, Data starting at Row 5/6).



## 📋 Technical Requirements
* **OS:** Windows (VBA FileSystemObject is Windows-based).
* **Software:** Microsoft Excel (2016 or newer recommended).
* **Permissions:** Ensure you have read/write permissions for the folders you are targeting.



## 🤝 Contribution & Feedback
I am a strategic partner focused on simplifying complexity. If there is a repetitive office task slowing you down, feel free to open an issue or suggest a new feature!


*Developed with a focus on precision, efficiency, and professional aesthetics.*
