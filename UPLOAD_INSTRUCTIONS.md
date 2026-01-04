# How to Upload Your Data Folder to GitHub

Since you're using GitHub's web interface directly, here's how to upload your data folder:

## Option 1: Upload via GitHub Web Interface

1. **Navigate to your repository** on GitHub
2. **Click "Add file"** → **"Upload files"**
3. **Create the data folder structure:**
   - In the file path box at the top, type: `data/`
   - This will create a folder structure
4. **Drag and drop your database file:**
   - Drag `delphi_projects.json` from your local `data/` folder
   - Or click "choose your files" and select it
5. **Commit the changes:**
   - Add a commit message like "Add initial database file"
   - Click "Commit changes"

## Option 2: Upload Folder Structure

If you have multiple files in your data folder:

1. **Click "Add file"** → **"Upload files"**
2. **Drag your entire `data/` folder** into the upload area
3. GitHub will preserve the folder structure
4. **Commit the changes**

## After Upload

The code has been updated to automatically look for the database in:
1. `data/delphi_projects.json` (if data folder exists)
2. `delphi_projects.json` (root directory - fallback)

So your app will work whether the file is in the data folder or root directory.

## Important Notes

- The `.gitignore` has been updated to allow data files
- Make sure your database file is named `delphi_projects.json`
- If you want to keep the data private, consider making the repository private
- For Streamlit Cloud deployment, the data file will be included automatically

