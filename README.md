# Organization Chart Generator - Google Apps Script

This Google Apps Script reads employee data from a Google Sheet and generates a clear, hierarchical organization chart in a Google Doc format.

## Features

- **Hierarchical Structure**: Builds proper org chart based on manager-employee relationships
- **Rich Employee Information**: Displays name, title, organization, location, email, and phone
- **Smart Sorting**: Orders employees by seniority and title hierarchy
- **Professional Formatting**: Clean, readable document with proper indentation and styling
- **Summary Statistics**: Includes breakdowns by organization, location, and title level
- **Error Handling**: Validates data and provides helpful error messages

## File Structure

```
organization-chart-generator/
├── orgchart-generator.gs      # Main script file
├── setup-helper.gs           # Setup and validation helpers
├── config.sample.gs          # Configuration template (safe to commit)
├── config.gs                # Your actual configuration (DO NOT COMMIT)
├── .gitignore               # Excludes config.gs from source control
└── README.md                # This documentation
```

## Setup Instructions

### 1. Prepare Your Google Sheet

1. Create a new Google Sheet or upload your CSV data
2. Ensure the first row contains column headers matching the expected format:
   - `Name` - Employee full name
   - `User ID` - Unique identifier for the employee
   - `Job Title` - Employee's job title
   - `Organization Name` - Department/organization name
   - `Location` - Work location
   - `Email` - Employee email address
   - `Telephone` / `Mobile` / `Home Phone` - Phone numbers (script will use first available)
   - `Manager UID` - User ID of the employee's manager
   - `Status` - Employment status (script filters for "Current Employee")

### 2. Create the Google Apps Script

1. Go to [Google Apps Script](https://script.google.com)
2. Click "New project"
3. Create the following files in your project:
   - Delete the default `Code.gs` file
   - Add `orgchart-generator.gs` (paste the main script contents)
   - Add `setup-helper.gs` (paste the setup helper contents)
   - Add `config.gs` (copy from `config.sample.gs` and customize)
4. Save the project with a meaningful name like "Organization Chart Generator"

### 3. Configure the Script

1. **Create your configuration file**:
   - Copy the contents of `config.sample.gs`
   - Create a new file in Google Apps Script called `config.gs`
   - Paste the contents and update with your actual IDs

2. **Update the configuration in `config.gs`**:
   ```javascript
   const SHEET_ID = 'your_actual_sheet_id_here';
   const DOC_ID = 'your_actual_doc_id_here'; // Optional
   ```

3. **To find your Google Sheet ID**:
   - Open your Google Sheet
   - Look at the URL: `https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit`
   - Copy the long string between `/d/` and `/edit`

4. **For the Google Doc ID** (optional):
   - If you want to update an existing document, create a new Google Doc first
   - Copy its ID from the URL in the same way
   - If you leave this as the default, a new document will be created each time

> **🔒 Security Note**: The `config.gs` file contains sensitive IDs and should not be committed to source control. It's included in `.gitignore` for this reason.

### 4. Set Up Permissions

1. Click the "Run" button (or select `generateOrgChart` function and run)
2. Grant the necessary permissions when prompted:
   - Access to Google Sheets (to read employee data)
   - Access to Google Docs (to create/update the organization chart)

## Usage

### Basic Usage

1. Run the `generateOrgChart()` function
2. The script will:
   - Read data from your Google Sheet
   - Build the hierarchical structure
   - Create or update a Google Doc with the organization chart
   - Print the document URL to the console

### Testing

1. Run `testDataReading()` to validate your data format
2. This will show you the first few employee records and total count

### Advanced Usage

#### Creating a New Document Each Time
Run `createNewOrgChart()` to always generate a new document regardless of the DOC_ID setting.

#### Scheduling Regular Updates
1. In Google Apps Script, go to "Triggers" (clock icon)
2. Add a new trigger for `generateOrgChart`
3. Set it to run daily, weekly, or on a custom schedule

## Output Format

The generated Google Doc includes:

### Main Organization Chart
- Hierarchical structure using proper unordered lists
- **Employee names in bold** with all information on one line
- Complete employee details following the name (title, organization, location, email, phone)
- Proper hierarchical nesting using bullet points
- Clean, readable single-line format for each employee

### Summary Statistics
- Total employee count
- Breakdown by organization
- Breakdown by location
- Breakdown by title level

## Troubleshooting

### Common Issues

1. **Configuration errors**:
   - Make sure you've created the `config.gs` file from `config.sample.gs`
   - Verify your SHEET_ID and DOC_ID are correctly set in `config.gs`
   - Check that the configuration file is in the same Google Apps Script project

2. **"Missing required columns" error**:
   - Ensure your sheet has all required column headers
   - Check for exact spelling and case sensitivity

3. **"Sheet appears to be empty" error**:
   - Verify your Google Sheet ID is correct in `config.gs`
   - Ensure the sheet has data beyond just headers

4. **Permission errors**:
   - Re-run the script and grant all requested permissions
   - Check that your Google account has access to both the sheet and ability to create docs

5. **Hierarchy looks wrong**:
   - Verify Manager UID values match existing User ID values
   - Check that the Manager UID column contains the correct identifiers

6. **"SHEET_ID is not defined" error**:
   - You haven't created the `config.gs` file
   - Copy `config.sample.gs` to `config.gs` and update with your IDs

### Data Quality Tips

- Ensure Manager UID values exactly match User ID values
- Remove any extra spaces or formatting from ID fields
- Use consistent naming for organizations and locations
- Verify all "Current Employee" status entries are correct

## Customization

### Modifying the Output Format

You can customize the organization chart by modifying these functions:

- `renderHierarchyTree()`: Change how each employee is displayed
- `getTitleRank()`: Adjust the hierarchy ranking of job titles
- `addSummaryStatistics()`: Modify or add new statistical breakdowns

### Adding New Data Fields

1. Update the `readEmployeeData()` function to include new fields
2. Modify the `renderHierarchyTree()` function to display the new information
3. Update the column validation in `readEmployeeData()` if needed

## Sample Data Format

Your Google Sheet should look like this:

| Name | User ID | Job Title | Organization Name | Location | Email | Manager UID | Status |
|------|---------|-----------|-------------------|----------|-------|-------------|---------|
| John Doe | jdoe | Senior Director | Engineering | Boston, MA | jdoe@company.com | | Current Employee |
| Jane Smith | jsmith | Director | Engineering | Boston, MA | jsmith@company.com | jdoe | Current Employee |

## License

This script is provided as-is for educational and business use. Feel free to modify and adapt it for your organization's needs. 
