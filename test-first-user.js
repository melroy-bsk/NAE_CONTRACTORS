// SINGLE USER TEST - Use Your Existing Login
// This will create ONLY ONE user for testing purposes

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs');

// === CREATE BROWSER WITH YOUR EXISTING LOGIN ===
async function createTestBrowser() {
    console.log('🔑 Opening test browser with your existing login...');

    const chromeOptions = new chrome.Options();

    // Stealth settings
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');
    chromeOptions.addArguments('--disable-web-security');

    // Use your existing Chrome profile
    const userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    chromeOptions.addArguments(`--user-data-dir=${userDataDir}`);
    chromeOptions.addArguments('--profile-directory=Default');

    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    // Hide automation
    await driver.executeScript(`
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
        });
    `);

    console.log('✅ Test browser ready with your login!');
    return driver;
}

// === LOAD JUST THE FIRST SECURITY CONTRACTOR ===
function loadSingleTestUser() {
    console.log('📁 Loading ONE test user from Excel...');

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

    console.log('📊 Column mapping:');
    console.log(`   Code: Column ${codeIndex} (${headers[codeIndex] || 'NOT FOUND'})`);
    console.log(`   First Name: Column ${firstNameIndex} (${headers[firstNameIndex] || 'NOT FOUND'})`);
    console.log(`   Last Name: Column ${lastNameIndex} (${headers[lastNameIndex] || 'NOT FOUND'})`);
    console.log(`   Department: Column ${departmentIndex} (${headers[departmentIndex] || 'NOT FOUND'})`);

    if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
        throw new Error('❌ Required columns not found');
    }

    // Find FIRST Security department contractor
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

                const testUser = {
                    code: code,
                    fullName: fullName,
                    firstName: firstName,
                    lastName: lastName,
                    department: department || 'Security',
                    username: username,
                    password: username
                };

                console.log('👤 Test user selected:');
                console.log(`   Name: ${testUser.fullName}`);
                console.log(`   Username: ${testUser.username}`);
                console.log(`   Code: ${testUser.code}`);

                return testUser;
            }
        }
    }

    throw new Error('❌ No Security department contractors found');
}

// === SINGLE USER TEST AUTOMATION ===
async function testCreateSingleUser() {
    let driver;

    try {
        console.log('🧪 SINGLE USER TEST - Starting...\n');

        // Load test user
        const testUser = loadSingleTestUser();

        // Create browser with existing login
        driver = await createTestBrowser();
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        // Navigate to site (should already be logged in)
        console.log('\n🌐 Navigating to Rivo Safeguard...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await driver.sleep(3000);

        // Check if logged in
        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Successfully using existing login!');
        } catch (error) {
            throw new Error('❌ Not logged in. Please log in to Chrome first.');
        }

        console.log(`\n👤 Creating TEST USER: ${testUser.username} (${testUser.fullName})`);

        // === NAVIGATION SEQUENCE ===
        console.log('🧭 Navigating to user creation...');

        await driver.findElement(By.css(".sch-container-left")).click();
        console.log('   ✅ Clicked left container');
        await driver.sleep(1500);

        await driver.findElement(By.css(".sch-app-launcher-button")).click();
        console.log('   ✅ Clicked app launcher');
        await driver.sleep(2000);

        await driver.findElement(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text")).click();
        console.log('   ✅ Clicked menu item');
        await driver.sleep(2500);

        await driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();
        console.log('   ✅ Clicked user management');
        await driver.sleep(3000);

        // Switch to frame
        console.log('🖼️  Switching to user creation frame...');
        await driver.switchTo().frame(0);
        await driver.sleep(2000);

        // === FILL BASIC INFORMATION ===
        console.log('📝 Filling basic user information...');

        // Username
        const usernameField = await driver.findElement(By.id("Username"));
        await usernameField.click();
        await usernameField.clear();
        await usernameField.sendKeys(testUser.username);
        console.log(`   ✅ Username: ${testUser.username}`);
        await driver.sleep(1000);

        // Password
        const passwordField = await driver.findElement(By.name("Password"));
        await passwordField.click();
        await passwordField.clear();
        await passwordField.sendKeys(testUser.password);
        console.log(`   ✅ Password: ${testUser.password}`);
        await driver.sleep(1000);

        // Job Title
        const jobTitleField = await driver.findElement(By.name("JobTitle"));
        await jobTitleField.click();
        await jobTitleField.clear();
        await jobTitleField.sendKeys(testUser.department);
        console.log(`   ✅ Job Title: ${testUser.department}`);
        await driver.sleep(1200);

        // First Name
        const forenameField = await driver.findElement(By.id("Attributes.People.Forename"));
        await forenameField.click();
        await forenameField.clear();
        await forenameField.sendKeys(testUser.firstName);
        console.log(`   ✅ First Name: ${testUser.firstName}`);
        await driver.sleep(1000);

        // Last Name
        const surnameField = await driver.findElement(By.id("Attributes.People.Surname"));
        await surnameField.click();
        await surnameField.clear();
        await surnameField.sendKeys(testUser.lastName);
        console.log(`   ✅ Last Name: ${testUser.lastName}`);
        await driver.sleep(1200);

        // Employee Number
        const employeeNumberField = await driver.findElement(By.id("Attributes.Users.EmployeeNumber"));
        await employeeNumberField.click();
        await employeeNumberField.clear();
        await employeeNumberField.sendKeys(testUser.code);
        console.log(`   ✅ Employee Number: ${testUser.code}`);
        await driver.sleep(1500);

        // === HIERARCHY CONFIGURATION ===
        console.log('🏢 Configuring hierarchy (simplified for test)...');
        try {
            await driver.findElement(By.id("Hierarchies01zzz")).click();
            await driver.sleep(1500);

            await driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await driver.sleep(1200);

            await driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();
            await driver.sleep(2000);

            // Try home location (optional)
            try {
                await driver.findElement(By.css("#\\31 657_HomeLocationHierarchy > option:nth-child(2)")).click();
                await driver.sleep(1500);
                console.log('   ✅ Home location set');
            } catch (error) {
                console.log('   ⚠️  Home location not found, skipping...');
            }

            await driver.findElement(By.name("addHierarchyGroupsLocations")).click();
            await driver.sleep(2000);

            await driver.findElement(By.css("#OverlayContainer > .StandardButton")).click();
            await driver.sleep(1500);
            console.log('   ✅ Basic hierarchy configured');

        } catch (error) {
            console.log('   ⚠️  Some hierarchy operations failed, continuing...');
        }

        // === USER GROUP SELECTION ===
        console.log('👥 Setting user group...');
        try {
            // Password confirmation
            const passwordField2 = await driver.findElement(By.name("Password"));
            await passwordField2.click();
            await passwordField2.clear();
            await passwordField2.sendKeys(testUser.password);
            await driver.sleep(1000);

            await driver.findElement(By.name("addHierarchyGroupsLocations")).click();
            await driver.sleep(2000);

            await driver.findElement(By.css("#ManageHierarchyUGroupLocationsForm > .GroupTab:nth-child(1)")).click();
            await driver.sleep(1500);

            // Select Third Party Staff
            const userGroupDropdown = await driver.findElement(By.id("UGroupID"));
            const thirdPartyOption = await userGroupDropdown.findElement(By.xpath("//option[. = 'Third Party Staff']"));
            await thirdPartyOption.click();
            await driver.sleep(2000);
            console.log('   ✅ User group: Third Party Staff');

        } catch (error) {
            console.log('   ⚠️  Some group operations failed, continuing...');
        }

        // === FINAL SETTINGS ===
        console.log('⚙️  Setting final options...');
        try {
            // Employment Status
            const statusDropdown = await driver.findElement(By.id("Attributes.Users.StatusOfEmployment"));
            const currentOption = await statusDropdown.findElement(By.xpath("//option[. = 'Current']"));
            await currentOption.click();
            await driver.sleep(1000);
            console.log('   ✅ Employment Status: Current');

            // User Type
            const userTypeDropdown = await driver.findElement(By.id("UserTypeID"));
            const limitedOption = await userTypeDropdown.findElement(By.xpath("//option[. = 'Limited access user']"));
            await limitedOption.click();
            await driver.sleep(2000);
            console.log('   ✅ User Type: Limited access user');

        } catch (error) {
            console.log('   ⚠️  Some final settings failed, continuing...');
        }

        // === ASK BEFORE SAVING ===
        console.log('\n🤔 READY TO SAVE USER');
        console.log(`📋 User Summary:`);
        console.log(`   Username: ${testUser.username}`);
        console.log(`   Name: ${testUser.fullName}`);
        console.log(`   Department: ${testUser.department}`);
        console.log(`   Code: ${testUser.code}`);

        console.log('\n❓ Do you want to SAVE this test user? (y/n)');

        // Wait for user input
        const readline = require('readline');
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        const answer = await new Promise(resolve => {
            rl.question('Enter y to save, n to cancel: ', resolve);
        });
        rl.close();

        if (answer.toLowerCase() === 'y') {
            // === SAVE USER ===
            console.log('💾 Saving test user...');
            await driver.findElement(By.name("save")).click();
            await driver.sleep(5000);

            console.log(`\n🎉 SUCCESS! Test user created: ${testUser.username}`);
            console.log('✅ Single user test completed!');
        } else {
            console.log('\n🚫 Test cancelled - user NOT saved');
            console.log('💡 You can review the filled form in the browser');
        }

    } catch (error) {
        console.error(`\n❌ Test failed: ${error.message}`);

        if (error.message.includes('Not logged in')) {
            console.log('\n💡 SOLUTION:');
            console.log('1. Open Chrome browser');
            console.log('2. Go to https://www.rivosafeguard.com/insight/');
            console.log('3. Log in to your account');
            console.log('4. Close ALL Chrome windows');
            console.log('5. Run this test again');
        }

        throw error;
    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 30 seconds to review...');
            await driver.sleep(30000); // Keep open for 30 seconds
            await driver.quit();
            console.log('✅ Browser closed');
        }
    }
}

// === RUN SINGLE USER TEST ===
if (require.main === module) {
    console.log('🧪 SINGLE USER TEST - Automation with Existing Login');
    console.log('📋 This will create ONLY ONE user for testing\n');

    console.log('⚠️  BEFORE RUNNING:');
    console.log('1. ✅ Log into https://www.rivosafeguard.com/insight/ in Chrome');
    console.log('2. ✅ Close ALL Chrome browser windows');
    console.log('3. ✅ Make sure Excel file is in this folder');
    console.log('4. ✅ Press Enter to start test...\n');

    // Wait for user to press Enter
    process.stdin.setRawMode(true);
    process.stdin.resume();
    process.stdin.once('data', () => {
        process.stdin.setRawMode(false);

        testCreateSingleUser()
            .then(() => {
                console.log('\n🏁 Single user test completed!');
                process.exit(0);
            })
            .catch((error) => {
                console.error('\n💥 Test failed!');
                console.error(`Error: ${error.message}`);
                process.exit(1);
            });
    });
}

module.exports = { testCreateSingleUser, loadSingleTestUser };