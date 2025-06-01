// SIMPLE CHROME PROFILE AUTOMATION - Ready to Use
// This uses your existing Chrome profile with all your logins

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs');

// === STEP 1: CREATE DRIVER WITH YOUR CHROME PROFILE ===
async function createDriverWithYourProfile() {
    console.log('🔑 Opening Chrome with your existing profile...');

    const chromeOptions = new chrome.Options();

    // === HIDE AUTOMATION ===
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');
    chromeOptions.addArguments('--disable-web-security');

    // === USE YOUR CHROME PROFILE ===
    // Automatically detect your Chrome profile path
    let userDataDir;

    if (os.platform() === 'win32') {
        // Windows
        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else if (os.platform() === 'darwin') {
        // macOS
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
    } else {
        // Linux
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    console.log(`📁 Using Chrome profile from: ${userDataDir}`);

    // Use your default profile (where you're already logged in)
    chromeOptions.addArguments(`--user-data-dir=${userDataDir}`);
    chromeOptions.addArguments('--profile-directory=Default');

    // If you have multiple profiles and want to use a different one:
    // chromeOptions.addArguments('--profile-directory=Profile 1');  // Second profile
    // chromeOptions.addArguments('--profile-directory=Profile 2');  // Third profile

    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    // Remove automation detection
    await driver.executeScript(`
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
        });
    `);

    console.log('✅ Chrome opened with your profile and existing logins!');
    return driver;
}

// === STEP 2: YOUR AUTOMATION WITH EXISTING LOGIN ===
async function runAutomationWithProfile() {
    let driver;

    try {
        console.log('🚀 Starting automation with your Chrome profile...\n');

        // IMPORTANT: Make sure ALL Chrome windows are closed first!
        console.log('⚠️  Make sure ALL Chrome windows are closed before continuing...');
        console.log('⏰ Waiting 3 seconds...\n');
        await new Promise(resolve => setTimeout(resolve, 3000));

        // Load your Excel data
        const contractor = loadFirstContractor();

        // Create driver with your profile
        driver = await createDriverWithYourProfile();
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        // Navigate to your site (you should already be logged in!)
        console.log('🌐 Navigating to Rivo Safeguard...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await driver.sleep(3000);

        // Check if you're logged in
        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Already logged in! Starting user creation...');
        } catch (error) {
            throw new Error('❌ Not logged in. Please log in to Chrome first.');
        }

        console.log(`\n👤 Creating user: ${contractor.username} (${contractor.fullName})`);

        // === YOUR EXISTING AUTOMATION CODE (no changes needed) ===

        // Navigate to user creation
        await driver.findElement(By.css(".sch-container-left")).click();
        await driver.sleep(1500);

        await driver.findElement(By.css(".sch-app-launcher-button")).click();
        await driver.sleep(2000);

        await driver.findElement(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text")).click();
        await driver.sleep(2500);

        await driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();
        await driver.sleep(3000);

        // Switch to frame
        await driver.switchTo().frame(0);
        await driver.sleep(2000);

        // Fill user form
        console.log('📝 Filling user information...');

        await fillField(driver, By.id("Username"), contractor.username, 'Username');
        await fillField(driver, By.name("Password"), contractor.password, 'Password');
        await fillField(driver, By.name("JobTitle"), contractor.department, 'Job Title');
        await fillField(driver, By.id("Attributes.People.Forename"), contractor.firstName, 'First Name');
        await fillField(driver, By.id("Attributes.People.Surname"), contractor.lastName, 'Last Name');
        await fillField(driver, By.id("Attributes.Users.EmployeeNumber"), contractor.code, 'Employee Number');

        // Handle hierarchy (simplified)
        console.log('🏢 Configuring hierarchy...');
        await handleHierarchy(driver);

        // Handle groups
        console.log('👥 Setting up user groups...');
        await handleGroups(driver, contractor.password);

        // Final settings
        console.log('⚙️  Final settings...');
        await handleFinalSettings(driver);

        // Save user
        console.log('💾 Saving user...');
        await driver.findElement(By.name("save")).click();
        await driver.sleep(5000);

        console.log(`\n🎉 SUCCESS! User created: ${contractor.username}`);

    } catch (error) {
        console.error(`\n❌ Automation failed: ${error.message}`);

        if (error.message.includes('user data directory is already in use')) {
            console.log('\n💡 SOLUTION:');
            console.log('1. Close ALL Chrome browser windows');
            console.log('2. Wait 10 seconds');
            console.log('3. Run this script again');
        } else if (error.message.includes('Not logged in')) {
            console.log('\n💡 SOLUTION:');
            console.log('1. Open Chrome browser');
            console.log('2. Go to https://www.rivosafeguard.com/insight/');
            console.log('3. Log in to your account');
            console.log('4. Close ALL Chrome windows');
            console.log('5. Run this script again');
        }

    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 10 seconds to see results...');
            await driver.sleep(10000);
            await driver.quit();
            console.log('✅ Browser closed');
        }
    }
}

// === HELPER FUNCTIONS ===

// Fill form field helper
async function fillField(driver, selector, value, description) {
    try {
        const field = await driver.findElement(selector);
        await field.click();
        await field.clear();
        await field.sendKeys(value);
        console.log(`   ✅ ${description}: ${value}`);
        await driver.sleep(800);
    } catch (error) {
        console.log(`   ⚠️  Failed to fill ${description}: ${error.message}`);
    }
}

// Handle hierarchy configuration
async function handleHierarchy(driver) {
    try {
        await driver.findElement(By.id("Hierarchies01zzz")).click();
        await driver.sleep(1500);

        await driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
        await driver.sleep(1200);

        await driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();
        await driver.sleep(2000);

        try {
            await driver.findElement(By.css("#\\31 657_HomeLocationHierarchy > option:nth-child(2)")).click();
            await driver.sleep(1500);
        } catch (e) {
            console.log('   ⚠️  Home location not found, continuing...');
        }

        await driver.findElement(By.name("addHierarchyGroupsLocations")).click();
        await driver.sleep(2000);

        await driver.findElement(By.css("#OverlayContainer > .StandardButton")).click();
        await driver.sleep(1500);

        console.log('   ✅ Hierarchy configured');

    } catch (error) {
        console.log('   ⚠️  Hierarchy configuration failed, continuing...');
    }
}

// Handle group settings
async function handleGroups(driver, password) {
    try {
        // Password confirmation
        await fillField(driver, By.name("Password"), password, 'Password confirmation');

        await driver.findElement(By.name("addHierarchyGroupsLocations")).click();
        await driver.sleep(2000);

        await driver.findElement(By.css("#ManageHierarchyUGroupLocationsForm > .GroupTab:nth-child(1)")).click();
        await driver.sleep(1500);

        // Select Third Party Staff
        const userGroupDropdown = await driver.findElement(By.id("UGroupID"));
        const thirdPartyOption = await userGroupDropdown.findElement(By.xpath("//option[. = 'Third Party Staff']"));
        await thirdPartyOption.click();
        await driver.sleep(2000);

        console.log('   ✅ User group set to Third Party Staff');

    } catch (error) {
        console.log('   ⚠️  Group configuration failed, continuing...');
    }
}

// Handle final settings
async function handleFinalSettings(driver) {
    try {
        // Employment status
        const statusDropdown = await driver.findElement(By.id("Attributes.Users.StatusOfEmployment"));
        const currentOption = await statusDropdown.findElement(By.xpath("//option[. = 'Current']"));
        await currentOption.click();
        await driver.sleep(1000);

        // User type
        const userTypeDropdown = await driver.findElement(By.id("UserTypeID"));
        const limitedOption = await userTypeDropdown.findElement(By.xpath("//option[. = 'Limited access user']"));
        await limitedOption.click();
        await driver.sleep(2000);

        console.log('   ✅ Employment status and user type set');

    } catch (error) {
        console.log('   ⚠️  Final settings failed, continuing...');
    }
}

// Load contractor data from Excel
function loadFirstContractor() {
    console.log('📁 Loading contractor from Excel...');

    const excelFile = 'Copy of Contractors staff list.xlsx';

    if (!fs.existsSync(excelFile)) {
        throw new Error(`❌ Excel file not found: ${excelFile}`);
    }

    function findColumnIndex(headers, possibleNames) {
        for (let i = 0; i < headers.length; i++) {
            const header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
            if (possibleNames.some(name => header.includes(name))) {
                return i;
            }
        }
        return -1;
    }

    const workbook = XLSX.readFile(excelFile);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = rawData[0] || [];
    const codeIndex = 0;
    const firstNameIndex = findColumnIndex(headers, ['first name', 'firstname', 'forename']);
    const lastNameIndex = findColumnIndex(headers, ['last name', 'lastname', 'surname']);
    const departmentIndex = findColumnIndex(headers, ['department', 'dept']);

    if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
        throw new Error('❌ Required columns not found');
    }

    // Find first Security contractor
    for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row[codeIndex] && row[firstNameIndex] && row[lastNameIndex]) {
            const department = row[departmentIndex] ? row[departmentIndex].toString().trim() : '';

            if (department.toLowerCase() === 'security') {
                const code = row[codeIndex].toString().trim();
                const firstName = row[firstNameIndex].toString().trim();
                const lastName = row[lastNameIndex].toString().trim();
                const fullName = `${firstName} ${lastName}`;
                const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                const contractor = {
                    code: code,
                    fullName: fullName,
                    firstName: firstName,
                    lastName: lastName,
                    department: department || 'Security',
                    username: username,
                    password: username
                };

                console.log(`👤 Selected: ${contractor.fullName} (${contractor.username})`);
                return contractor;
            }
        }
    }

    throw new Error('❌ No Security contractors found');
}

// === RUN THE AUTOMATION ===
if (require.main === module) {
    console.log('🎯 CHROME PROFILE AUTOMATION');
    console.log('🔑 This will use your existing Chrome profile with all logins\n');

    console.log('📋 BEFORE RUNNING:');
    console.log('1. ✅ Make sure you are logged into https://www.rivosafeguard.com/insight/');
    console.log('2. ✅ Close ALL Chrome browser windows');
    console.log('3. ✅ Press Enter to start...\n');

    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.once('data', () => {
        process.stdin.setRawMode(false);

        runAutomationWithProfile()
            .then(() => {
                console.log('\n🏁 Automation completed!');
                process.exit(0);
            })
            .catch((error) => {
                console.error('\n💥 Automation failed!');
                console.error(`Error: ${error.message}`);
                process.exit(1);
            });
    });
}

module.exports = { createDriverWithYourProfile, runAutomationWithProfile };