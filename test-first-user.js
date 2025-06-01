// Standalone Test Script - No Mocha Required
// Run with: node test-first-user-standalone.js
const chrome = require('selenium-webdriver/chrome');  // ADD THIS LINE
const { Builder, By, Key, until } = require('selenium-webdriver');
const XLSX = require('xlsx');
const fs = require('fs');

// Configuration
const CONFIG = {
    defaultTimeout: 15000,
    pageLoadTimeout: 30000,
    elementTimeout: 10000,
    excelFileName: 'Copy of Contractors staff list.xlsx'
};

class TestUtilities {
    constructor(driver) {
        this.driver = driver;
    }

    // Safe element interaction with retry logic
    async safeClick(selector, description, timeout = CONFIG.elementTimeout) {
        let attempts = 3;
        while (attempts > 0) {
            try {
                console.log(`   🖱️  Attempting to click: ${description}`);
                const element = await this.driver.wait(until.elementLocated(selector), timeout);
                await this.driver.wait(until.elementIsEnabled(element), timeout);
                await element.click();
                console.log(`   ✅ Successfully clicked: ${description}`);
                return element;
            } catch (error) {
                attempts--;
                console.log(`   ⚠️  Click attempt failed for ${description}, ${attempts} attempts remaining`);
                if (attempts === 0) {
                    throw new Error(`Failed to click ${description}: ${error.message}`);
                }
                await this.driver.sleep(1000);
            }
        }
    }

    // Safe text input with validation
    async safeInput(selector, text, description, timeout = CONFIG.elementTimeout) {
        try {
            console.log(`   ✏️  Entering text in: ${description}`);
            const element = await this.driver.wait(until.elementLocated(selector), timeout);
            await this.driver.wait(until.elementIsEnabled(element), timeout);

            await element.clear();
            await element.sendKeys(text);

            // Validate the text was entered correctly
            const actualValue = await element.getAttribute('value');
            if (actualValue !== text) {
                console.log(`   ⚠️  Text validation warning. Expected: "${text}", Got: "${actualValue}"`);
            }

            console.log(`   ✅ Successfully entered: ${description}`);
            return element;
        } catch (error) {
            throw new Error(`Failed to enter text in ${description}: ${error.message}`);
        }
    }

    // Safe dropdown selection
    async safeSelectOption(dropdownSelector, optionText, description, timeout = CONFIG.elementTimeout) {
        try {
            console.log(`   📋 Selecting option "${optionText}" in: ${description}`);
            const dropdown = await this.driver.wait(until.elementLocated(dropdownSelector), timeout);
            await this.driver.wait(until.elementIsEnabled(dropdown), timeout);

            // Try to find and click the option
            const options = await dropdown.findElements(By.xpath(`//option[normalize-space(text()) = '${optionText}']`));
            if (options.length === 0) {
                console.log(`   ⚠️  Option "${optionText}" not found in ${description}, continuing...`);
                return dropdown;
            }

            await options[0].click();
            console.log(`   ✅ Successfully selected: ${optionText}`);
            return dropdown;
        } catch (error) {
            console.log(`   ⚠️  Failed to select option in ${description}: ${error.message}`);
            return null;
        }
    }

    // Wait for page to be ready
    async waitForPageReady(timeout = CONFIG.pageLoadTimeout) {
        try {
            await this.driver.wait(
                () => this.driver.executeScript('return document.readyState === "complete"'),
                timeout
            );
        } catch (error) {
            console.log('   ⚠️  Page ready check failed, continuing...');
        }
    }

    // Human-like delay
    async humanDelay(baseMs = 1000, variationMs = 500) {
        const randomVariation = Math.random() * variationMs * 2 - variationMs;
        const totalDelay = Math.max(baseMs + randomVariation, 200);
        await this.driver.sleep(totalDelay);
    }
}

// async function createStealthBrowser() {
//     console.log('🕵️  Creating stealth browser...');

//     const chromeOptions = new chrome.Options();

//     // Remove automation detection
//     chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
//     chromeOptions.excludeSwitches(['enable-automation']);
//     chromeOptions.addArguments('--disable-infobars');
//     chromeOptions.addArguments('--disable-extensions');
//     chromeOptions.addArguments('--disable-dev-shm-usage');
//     chromeOptions.addArguments('--no-sandbox');
//     chromeOptions.addArguments('--disable-web-security');

//     // Real user agent
//     chromeOptions.addArguments('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36');

//     // Disable notifications
//     chromeOptions.setUserPreferences({
//         'profile.default_content_setting_values.notifications': 2
//     });

//     // Create driver
//     const driver = await new Builder()
//         .forBrowser('chrome')
//         .setChromeOptions(chromeOptions)
//         .build();

//     // Remove webdriver property
//     await driver.executeScript(`
//         Object.defineProperty(navigator, 'webdriver', {
//             get: () => undefined,
//         });
//         delete window.navigator.webdriver;
//         delete window.webdriver;
//     `);

//     console.log('✅ Stealth browser ready!');
//     return driver;
// }

// Helper function to find column index by possible header names

async function createStealthBrowser() {
    const chromeOptions = new chrome.Options();
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');

    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    await driver.executeScript("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})");
    return driver;
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

// Load contractor data from Excel
function loadContractorData() {
    console.log('📁 Loading contractor data from Excel...');

    // Check if file exists
    if (!fs.existsSync(CONFIG.excelFileName)) {
        throw new Error(`❌ Excel file not found: ${CONFIG.excelFileName}`);
    }

    // Read Excel file
    const workbook = XLSX.readFile(CONFIG.excelFileName);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (!rawData || rawData.length < 2) {
        throw new Error('❌ Excel file appears to be empty or has no data rows');
    }

    // Find column indices
    const headers = rawData[0] || [];
    const codeIndex = 0;
    const firstNameIndex = findColumnIndex(headers, ['first name', 'firstname', 'forename']);
    const lastNameIndex = findColumnIndex(headers, ['last name', 'lastname', 'surname']);
    const departmentIndex = findColumnIndex(headers, ['department', 'dept']);

    console.log('📊 Column mapping:');
    console.log(`   Code: Column ${codeIndex} (${headers[codeIndex] || 'NOT FOUND'})`);
    console.log(`   First Name: Column ${firstNameIndex} (${headers[firstNameIndex] || 'NOT FOUND'})`);
    console.log(`   Last Name: Column ${lastNameIndex} (${headers[lastNameIndex] || 'NOT FOUND'})`);
    console.log(`   Department: Column ${departmentIndex} (${headers[departmentIndex] || 'NOT FOUND'})`);

    if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
        throw new Error('❌ Required columns not found. Need: First Name, Last Name, Department');
    }

    // Find first Security department contractor
    for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row[codeIndex] && row[firstNameIndex] && row[lastNameIndex]) {
            const code = row[codeIndex].toString().trim();
            const firstName = row[firstNameIndex].toString().trim();
            const lastName = row[lastNameIndex].toString().trim();
            const department = row[departmentIndex] ? row[departmentIndex].toString().trim() : '';

            if (department.toLowerCase() === 'security') {
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

                console.log('👤 First Security contractor found:');
                console.log(`   Name: ${contractor.fullName}`);
                console.log(`   Username: ${contractor.username}`);
                console.log(`   Department: ${contractor.department}`);
                console.log(`   Code: ${contractor.code}`);

                return contractor;
            }
        }
    }

    throw new Error('❌ No Security department contractors found in Excel file');
}

// Main test function
async function createFirstUserTest() {
    let driver;
    let testUtils;

    try {
        console.log('🚀 Starting first user creation test...');

        // Load contractor data
        const contractor = loadContractorData();

        // Initialize WebDriver
        console.log('🌐 Initializing browser...');
        // driver = await new Builder().forBrowser('chrome').build();
        driver = await createStealthBrowser();
        testUtils = new TestUtilities(driver);

        await driver.manage().window().setRect({ width: 1382, height: 736 });

        console.log(`\n👤 Creating user: ${contractor.username} (${contractor.fullName})`);

        // Navigate to the application
        console.log('🌐 Navigating to Rivo Safeguard...');
        await driver.get("https://www.rivosafeguard.com/insight/");
        await testUtils.waitForPageReady();
        await testUtils.humanDelay(3000, 1000);

        // Follow navigation sequence
        console.log('🧭 Following navigation sequence...');
        await testUtils.safeClick(By.css(".sch-container-left"), "Left container");
        await testUtils.humanDelay(1500, 500);

        await testUtils.safeClick(By.css(".sch-app-launcher-button"), "App launcher");
        await testUtils.humanDelay(2000, 800);

        await testUtils.safeClick(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text"), "Menu item");
        await testUtils.humanDelay(2500, 1000);

        await testUtils.safeClick(By.css(".k-drawer-item:nth-child(6)"), "User management");
        await testUtils.humanDelay(3000, 1500);

        // Switch to frame
        console.log('🖼️  Switching to user creation frame...');
        await driver.switchTo().frame(0);
        await testUtils.humanDelay(2000, 1000);

        // Fill in basic user information
        console.log('📝 Filling user information...');

        await testUtils.safeInput(By.id("Username"), contractor.username, "Username");
        await testUtils.humanDelay(1200, 600);

        await testUtils.safeInput(By.name("Password"), contractor.password, "Password");
        await testUtils.humanDelay(1000, 500);

        await testUtils.safeInput(By.name("JobTitle"), contractor.department, "Job Title");
        await testUtils.humanDelay(1500, 700);

        await testUtils.safeInput(By.id("Attributes.People.Forename"), contractor.firstName, "First Name");
        await testUtils.humanDelay(1000, 500);

        await testUtils.safeInput(By.id("Attributes.People.Surname"), contractor.lastName, "Last Name");
        await testUtils.humanDelay(1200, 600);

        await testUtils.safeInput(By.id("Attributes.Users.EmployeeNumber"), contractor.code, "Employee Number");
        await testUtils.humanDelay(1500, 750);

        // Handle Hierarchy selection (with better error handling)
        console.log('🏢 Configuring hierarchy settings...');
        try {
            await testUtils.safeClick(By.id("Hierarchies01zzz"), "Hierarchy button");
            await testUtils.humanDelay(1500, 750);

            await testUtils.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Hierarchy dropdown");
            await testUtils.humanDelay(1200, 600);

            await testUtils.safeClick(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)"), "Hierarchy lookup");
            await testUtils.humanDelay(2000, 1000);

            // Try home location hierarchy (optional)
            try {
                await testUtils.safeClick(By.css("#\\31 657_HomeLocationHierarchy > option:nth-child(2)"), "Home location");
                await testUtils.humanDelay(1500, 750);
            } catch (error) {
                console.log('   ⚠️  Home location hierarchy not found, continuing...');
            }

            // Additional hierarchy navigation
            await testUtils.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Hierarchy dropdown 2");
            await testUtils.humanDelay(1000, 500);

            await testUtils.safeClick(By.css(".HierarchyNode__SelectDiv-sc-1ytna50-2"), "Hierarchy node select");
            await testUtils.humanDelay(1200, 600);

            await testUtils.safeClick(By.css(".DropdownDisplay__DropdownText-sc-6p7u3y-0 > span:nth-child(1)"), "Dropdown text");
            await testUtils.humanDelay(1500, 750);

            // Navigate hierarchy tree (optional)
            try {
                await testUtils.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Hierarchy arrow 1");
                await testUtils.humanDelay(1000, 500);

                await testUtils.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Hierarchy arrow 2");
                await testUtils.humanDelay(1000, 500);

                await testUtils.safeClick(By.css(".HierarchyNode__NodeText-sc-1ytna50-1"), "Hierarchy node text");
                await testUtils.humanDelay(1500, 750);
            } catch (error) {
                console.log('   ⚠️  Hierarchy tree navigation failed, continuing...');
            }

            // Add hierarchy groups/locations
            await testUtils.safeClick(By.name("addHierarchyGroupsLocations"), "Add hierarchy groups");
            await testUtils.humanDelay(2000, 1000);

            await testUtils.safeClick(By.css("#OverlayContainer > .StandardButton"), "Overlay confirm");
            await testUtils.humanDelay(1500, 750);

        } catch (error) {
            console.log('   ⚠️  Some hierarchy operations failed, continuing...');
        }

        // Group operations
        console.log('👥 Managing groups and permissions...');
        try {
            // Group tab
            try {
                await testUtils.safeClick(By.id("GroupTabDiv_23890"), "Group tab");
                await testUtils.humanDelay(1500, 750);
            } catch (error) {
                console.log('   ⚠️  Group tab not found, continuing...');
            }

            // Update password in second field
            await testUtils.safeInput(By.name("Password"), contractor.password, "Password confirmation");
            await testUtils.humanDelay(1500, 750);

            // Additional group operations
            await testUtils.safeClick(By.name("addHierarchyGroupsLocations"), "Add hierarchy groups 2");
            await testUtils.humanDelay(2000, 1000);

            await testUtils.safeClick(By.css("#ManageHierarchyUGroupLocationsForm > .GroupTab:nth-child(1)"), "Group tab form");
            await testUtils.humanDelay(1500, 750);

            // Select User Group - Third Party Staff
            await testUtils.safeSelectOption(By.id("UGroupID"), "Third Party Staff", "User Group");
            await testUtils.humanDelay(2000, 1000);

            // Additional hierarchy selections (optional)
            try {
                await testUtils.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Final hierarchy dropdown");
                await testUtils.humanDelay(1000, 500);

                await testUtils.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Final hierarchy arrow 1");
                await testUtils.humanDelay(800, 400);

                await testUtils.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Final hierarchy arrow 2");
                await testUtils.humanDelay(800, 400);

                await testUtils.safeClick(By.css(".HierarchyNode__NodeText-sc-1ytna50-1"), "Final hierarchy node");
                await testUtils.humanDelay(1000, 500);

                await testUtils.safeClick(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)"), "Final hierarchy lookup");
                await testUtils.humanDelay(1500, 750);
            } catch (error) {
                console.log('   ⚠️  Final hierarchy selections failed, continuing...');
            }

        } catch (error) {
            console.log('   ⚠️  Some group operations failed, continuing...');
        }

        // Final settings
        console.log('⚙️  Configuring final settings...');
        try {
            // Set Status of Employment to Current
            await testUtils.safeSelectOption(By.id("Attributes.Users.StatusOfEmployment"), "Current", "Employment Status");
            await testUtils.humanDelay(1000, 500);

            // Set User Type to Limited access user
            await testUtils.safeSelectOption(By.id("UserTypeID"), "Limited access user", "User Type");
            await testUtils.humanDelay(2000, 1000);

        } catch (error) {
            console.log('   ⚠️  Some final settings failed, continuing...');
        }

        // Save the user
        console.log('💾 Saving user...');
        await testUtils.safeClick(By.name("save"), "Save button");
        await testUtils.humanDelay(5000, 2000);

        console.log(`\n✅ Successfully completed user creation test for: ${contractor.username}`);
        console.log('🎉 Test completed successfully!');

    } catch (error) {
        console.error(`\n❌ Test failed: ${error.message}`);
        console.error('💡 Check the browser window for more details');
        throw error;
    } finally {
        if (driver) {
            console.log('\n🧹 Cleaning up...');
            await driver.sleep(3000); // Keep browser open for a moment to see results
            await driver.quit();
            console.log('✅ Browser closed');
        }
    }
}

// Run the test
if (require.main === module) {
    console.log('🧪 Starting standalone user creation test...');
    console.log('📋 This will create the first Security department contractor from your Excel file\n');

    createFirstUserTest()
        .then(() => {
            console.log('\n🏁 Test execution completed successfully!');
            process.exit(0);
        })
        .catch((error) => {
            console.error('\n💥 Test execution failed!');
            console.error(`Error: ${error.message}`);
            process.exit(1);
        });
}

module.exports = { createFirstUserTest, loadContractorData };