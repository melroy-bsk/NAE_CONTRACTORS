# User Creation Automation (Security Department Only)

This project automates the creation of users from an Excel file using Selenium WebDriver, filtered for **Security department contractors only**.

## Data Summary

- **Total contractors in Excel**: 164
- **Security department contractors**: 29
- **Departments available**: Canteen, Crossfit, Facilities (Caretakers), Facilities (Cleaners), IT, Marketing, Operations, Security, Transportation, Vending Machine

## Prerequisites

1. **Node.js** (version 14 or higher)
2. **Firefox browser** installed on your system
3. **geckodriver** for Firefox automation
4. Excel file named `Copy of Contractors staff list.xlsx` in the same directory

## Setup

### 1. Install Node.js Dependencies

```bash
npm install selenium-webdriver xlsx mocha
```

Or use the provided script:

```bash
npm run install-deps
```

### 2. Install GeckoDriver

**Option A: Using npm (Recommended)**

```bash
npm install -g geckodriver
```

**Option B: Manual Installation**

- Download geckodriver from: https://github.com/mozilla/geckodriver/releases
- Extract and add to your system PATH

**Option C: Using package managers**

```bash
# macOS with Homebrew
brew install geckodriver

# Ubuntu/Debian
sudo apt-get install firefox-geckodriver

# Windows with Chocolatey
choco install selenium-gecko-driver
```

### 3. Verify Setup

Test if geckodriver is accessible:

```bash
geckodriver --version
```

## File Structure

```
project-folder/
├── user-creation-automation.js    # Main automation script
├── test-first-user.js            # Test script for first user only
├── package.json                  # Node.js dependencies
├── Copy of Contractors staff list.xlsx  # Your Excel data file
└── README.md                     # This file
```

## Excel File Format

Your Excel file should have these columns:

- **Column A**: Contractor code (e.g., "OC_AS069")
- **Column B**: Full Name (e.g., "Abdul Gafoor Karumbil")
- **Column C**: First Name (can be empty, will be parsed from Full Name)
- **Column D**: Last Name (can be empty, will be parsed from Full Name)
- **Column E**: Department (e.g., "Transportation")
- **Column F**: Contracted Company (ignored)

## Usage

### Option 1: Test with First User Only

```bash
# Run as a test
npm test

# Or run directly
npm run create-first
```

This will:

1. Read the Excel file
2. Create only the first user
3. Show you the process step by step

### Option 2: Create All Users

```bash
npm run create-all
```

This will:

1. Load all contractors from Excel
2. Show you the first 5 users for verification
3. Ask if you want to create just the first user (testing) or all users
4. Proceed with user creation

## Username/Password Generation

- **Format**: `{contractor_code}_{first_name}_{last_name}` (all lowercase)
- **Example**: `oc_as069_abdul_gafoor_karumbil`
- **Password**: Same as username

## How It Works

1. **Excel Reading**: Parses the Excel file and extracts contractor data
2. **Name Parsing**: Splits "Full Name" into first and last names
3. **Navigation**: Automates browser navigation to the user creation page
4. **Form Filling**: Fills in all required fields with contractor data
5. **User Creation**: Completes the user creation process
6. **Loop**: For multiple users, navigates back to add the next user

## Important Notes

### Browser Behavior

- The script uses Firefox browser
- Browser window will open and you can watch the automation
- Do not interact with the browser while the script is running

### Error Handling

- If a user creation fails, the script will ask if you want to continue
- Errors are logged to console with details
- The script includes retries for common issues

### Performance

- There's a 2-3 second delay between each user creation
- Total time depends on number of users and system performance
- For 29 Security contractors, expect ~4-6 minutes

### Troubleshooting

**Common Issues:**

1. **"geckodriver not found"**

   - Install geckodriver and ensure it's in PATH
   - Try: `npm install -g geckodriver`

2. **"Excel file not found"**

   - Ensure `Copy of Contractors staff list.xlsx` is in the same directory
   - Check file name spelling and case

3. **"Element not found"**

   - Website layout may have changed
   - Check if selectors in the script match current website
   - Try running with first user only to debug

4. **"Timeout errors"**
   - Increase timeout in script if needed
   - Check internet connection
   - Ensure website is accessible

**Debug Mode:**
Add more logging to see what's happening:

```javascript
console.log("Current step: ...", elementFound);
```

## Customization

### Modify User Data

Edit the contractor object creation in the scripts to change filtering:

```javascript
// Current filter (Security only)
if (department.toLowerCase() === "security") {
  // Create contractor object
}

// To include all departments, remove the filter:
// if (row && row[0] && row[1]) {

// To filter for different department:
// if (department.toLowerCase() === 'transportation') {
```

### Change Browser

Replace `'firefox'` with `'chrome'` in Builder() calls:

```javascript
driver = await new Builder().forBrowser("chrome").build();
```

(Requires chromedriver installation)

### Modify Selectors

If website changes, update CSS selectors in the scripts:

```javascript
await driver.findElement(By.css(".new-selector")).click();
```

## Security Note

- Passwords are visible in console output during execution
- Consider running in a secure environment
- Review generated usernames/passwords before deployment

## Support

If you encounter issues:

1. Check the troubleshooting section above
2. Verify all prerequisites are installed
3. Test with the first user only option
4. Check browser console for additional error messages
