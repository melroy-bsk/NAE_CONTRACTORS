// USE EXISTING CHROME PROFILE WITH GOOGLE SSO
// Connects to your already-running Chrome or starts it with your profile

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs');
const readline = require('readline');
const { exec } = require('child_process');

// === CONFIGURATION ===
const CONFIG = {
    excelFileName: 'Copy of Contractors staff list.xlsx',
    defaultTimeout: 30000,
    humanDelayBase: 1500,
    humanDelayVariation: 1000,
    debugPort: 9222
};

// === CHECK IF CHROME IS RUNNING WITH DEBUGGING ===
async function isChromeDebuggingEnabled() {
    return new Promise((resolve) => {
        exec(`curl -s http://localhost:${CONFIG.debugPort}/json/version`, (error, stdout) => {
            if (!error && stdout) {
                try {
                    const info = JSON.parse(stdout);
                    console.log('✅ Found Chrome with debugging enabled');
                    console.log(`   Browser: ${info.Browser}`);
                    resolve(true);
                } catch (e) {
                    resolve(false);
                }
            } else {
                resolve(false);
            }
        });
    });
}

// === INSTRUCTIONS FOR MANUAL CHROME START ===
async function showChromeStartInstructions() {
    console.log('\n' + '═'.repeat(70));
    console.log('⚠️  CHROME SETUP REQUIRED');
    console.log('═'.repeat(70));
    console.log('\n📋 Please start Chrome with remote debugging:\n');

    const platform = os.platform();

    if (platform === 'win32') {
        console.log('🖥️  WINDOWS Instructions:');
        console.log('1. Close all Chrome windows');
        console.log('2. Press Win+R, then paste this command:\n');
        console.log('   chrome.exe --remote-debugging-port=9222');
        console.log('\n   Or use full path:');
        console.log('   "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe" --remote-debugging-port=9222\n');
    } else if (platform === 'darwin') {
        console.log('🍎 MAC Instructions:');
        console.log('1. Close all Chrome windows');
        console.log('2. Open Terminal and paste:\n');
        console.log('   /Applications/Google\\ Chrome.app/Contents/MacOS/Google\\ Chrome --remote-debugging-port=9222\n');
    } else {
        console.log('🐧 LINUX Instructions:');
        console.log('1. Close all Chrome windows');
        console.log('2. Open Terminal and paste:\n');
        console.log('   google-chrome --remote-debugging-port=9222\n');
    }

    console.log('3. Chrome will open with YOUR profile and saved logins');
    console.log('4. You can browse normally - your Google session will work\n');
    console.log('═'.repeat(70));

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    await new Promise(resolve => {
        rl.question('\n✅ Press Enter after starting Chrome with the command above: ', () => {
            rl.close();
            resolve();
        });
    });
}

// === CONNECT TO YOUR EXISTING CHROME ===
async function connectToYourChrome() {
    console.log('🔌 Connecting to your Chrome profile...\n');

    // Check if Chrome is running with debugging
    let isDebuggingEnabled = await isChromeDebuggingEnabled();

    if (!isDebuggingEnabled) {
        await showChromeStartInstructions();

        // Check again
        isDebuggingEnabled = await isChromeDebuggingEnabled();
        if (!isDebuggingEnabled) {
            throw new Error('Chrome debugging port not found. Please follow the instructions above.');
        }
    }

    try {
        // Connect to existing Chrome
        const chromeOptions = new chrome.Options();
        chromeOptions.debuggerAddress(`localhost:${CONFIG.debugPort}`);

        console.log('🔗 Connecting to Chrome...');

        const driver = await new Builder()
            .forBrowser('chrome')
            .setChromeOptions(chromeOptions)
            .build();

        // Test connection
        try {
            const currentUrl = await driver.getCurrentUrl();
            console.log('✅ Successfully connected to your Chrome!');
            console.log(`📍 Current page: ${currentUrl}`);
            console.log('🔐 Your Google sessions are active\n');
        } catch (e) {
            console.log('✅ Connected to Chrome!\n');
        }

        // Remove automation indicators
        await driver.executeScript(`
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
        `).catch(() => { });

        return driver;

    } catch (error) {
        console.error('❌ Connection failed:', error.message);
        console.log('\n💡 Make sure you:');
        console.log('1. Started Chrome with --remote-debugging-port=9222');
        console.log('2. Chrome is running and accessible');
        throw error;
    }
}

// === NATURAL HUMAN AUTOMATOR ===
class NaturalHumanAutomator {
    constructor(driver) {
        this.driver = driver;
    }

    async randomDelay(min = 500, max = 2000) {
        const delay = Math.floor(Math.random() * (max - min + 1)) + min;
        await this.driver.sleep(delay);
    }

    async humanType(element, text) {
        await element.click();
        await this.randomDelay(200, 500);

        // Clear field naturally
        await element.sendKeys(Key.chord(Key.CONTROL, "a"));
        await this.randomDelay(100, 300);
        await element.sendKeys(Key.DELETE);
        await this.randomDelay(200, 400);

        // Type with natural rhythm
        for (let i = 0; i < text.length; i++) {
            await element.sendKeys(text[i]);

            if (Math.random() < 0.1) {
                // Occasional pause
                await this.driver.sleep(300 + Math.random() * 200);
            } else {
                // Normal speed
                await this.driver.sleep(50 + Math.random() * 100);
            }
        }

        await this.randomDelay(300, 600);
    }

    async naturalClick(element) {
        await this.driver.executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element);
        await this.randomDelay(500, 1000);

        const actions = this.driver.actions({ async: true });
        await actions.move({ origin: element }).perform();
        await this.randomDelay(200, 400);
        await element.click();
        await this.randomDelay(300, 600);
    }

    async waitAndType(selector, text, description) {
        try {
            console.log(`   ✏️  ${description}: "${text}"`);

            const element = await this.driver.wait(
                until.elementLocated(selector),
                CONFIG.defaultTimeout
            );

            await this.driver.wait(until.elementIsVisible(element), 5000);
            await this.driver.wait(until.elementIsEnabled(element), 5000);

            await this.humanType(element, text);
            console.log(`   ✅ Entered ${description}`);

            return element;

        } catch (error) {
            console.log(`   ⚠️  Failed to enter ${description}: ${error.message}`);
            throw error;
        }
    }

    async waitAndClick(selector, description) {
        try {
            console.log(`   🖱️  Clicking: ${description}`);

            const element = await this.driver.wait(
                until.elementLocated(selector),
                CONFIG.defaultTimeout
            );

            await this.driver.wait(until.elementIsVisible(element), 5000);
            await this.driver.wait(until.elementIsEnabled(element), 5000);

            await this.naturalClick(element);
            console.log(`   ✅ Clicked ${description}`);

            return element;

        } catch (error) {
            console.log(`   ⚠️  Failed to click ${description}: ${error.message}`);
            throw error;
        }
    }

    async selectDropdownOption(dropdownSelector, optionText, description) {
        try {
            console.log(`   📋 Selecting "${optionText}" from ${description}`);

            const dropdown = await this.driver.wait(
                until.elementLocated(dropdownSelector),
                CONFIG.defaultTimeout
            );

            await this.naturalClick(dropdown);
            await this.randomDelay(500, 1000);

            const option = await this.driver.wait(
                until.elementLocated(By.xpath(`//option[normalize-space(text())='${optionText}']`)),
                5000
            );

            await option.click();
            console.log(`   ✅ Selected "${optionText}"`);

            await this.randomDelay(300, 600);
            return dropdown;

        } catch (error) {
            console.log(`   ⚠️  Could not select option: ${error.message}`);
            return null;
        }
    }
}

// === SIMPLE NAVIGATION PROMPT ===
async function waitForUserNavigation(message = "Navigate to the desired page") {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('\n' + '═'.repeat(60));
    console.log('🧭 MANUAL NAVIGATION');
    console.log('═'.repeat(60));
    console.log(`\n${message}\n`);

    await new Promise(resolve => {
        rl.question('✅ Press Enter when ready: ', () => {
            rl.close();
            resolve();
        });
    });

    console.log('\n');
}

// === DETECT USER CREATION FORM ===
async function detectUserCreationForm(driver) {
    console.log('🔍 Looking for user creation form...');

    try {
        await driver.switchTo().defaultContent();

        // Check main content
        try {
            await driver.findElement(By.id("Username"));
            console.log('✅ Form found in main content');
            return true;
        } catch (e) {
            // Not in main content
        }

        // Check iframes
        const frames = await driver.findElements(By.css('iframe'));
        for (let i = 0; i < frames.length; i++) {
            try {
                await driver.switchTo().defaultContent();
                await driver.switchTo().frame(i);

                await driver.findElement(By.id("Username"));
                console.log(`✅ Form found in iframe ${i}`);
                return true;
            } catch (e) {
                // Continue
            }
        }

        await driver.switchTo().defaultContent();
        return false;

    } catch (error) {
        return false;
    }
}

// === FILL USER FORM ===
async function fillUserForm(driver, contractor, automator) {
    console.log(`\n👤 Creating user: ${contractor.fullName}`);
    console.log(`📝 Username: ${contractor.username}`);

    try {
        // Fill required fields
        await automator.waitAndType(By.id("Username"), contractor.username, "Username");
        await automator.randomDelay(1000, 2000);

        await automator.waitAndType(By.name("Password"), contractor.password, "Password");
        await automator.randomDelay(1000, 2000);

        await automator.waitAndType(By.name("JobTitle"), contractor.department, "Job Title");
        await automator.randomDelay(1000, 2000);

        await automator.waitAndType(By.id("Attributes.People.Forename"), contractor.firstName, "First Name");
        await automator.randomDelay(1000, 2000);

        await automator.waitAndType(By.id("Attributes.People.Surname"), contractor.lastName, "Last Name");
        await automator.randomDelay(1000, 2000);

        await automator.waitAndType(By.id("Attributes.Users.EmployeeNumber"), contractor.code, "Employee Number");
        await automator.randomDelay(1000, 2000);

        // Try optional fields
        try {
            await automator.selectDropdownOption(
                By.id("Attributes.Users.StatusOfEmployment"),
                "Current",
                "Employment Status"
            );
        } catch (e) { }

        try {
            await automator.selectDropdownOption(
                By.id("UserTypeID"),
                "Limited access user",
                "User Type"
            );
        } catch (e) { }

        // Confirm save
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        console.log('\n📋 Review and confirm:');
        const shouldSave = await new Promise(resolve => {
            rl.question('💾 Save this user? (y/n): ', (answer) => {
                rl.close();
                resolve(answer);
            });
        });

        if (shouldSave.toLowerCase() === 'y') {
            console.log('💾 Saving...');
            await automator.waitAndClick(By.name("save"), "Save button");
            await automator.randomDelay(3000, 5000);
            console.log(`✅ Created: ${contractor.username}`);
            return true;
        } else {
            console.log('🚫 Skipped');
            return false;
        }

    } catch (error) {
        console.error(`❌ Error: ${error.message}`);
        return false;
    }
}

// === LOAD CONTRACTORS ===
function loadContractorsFromExcel(mode = 'single') {
    console.log('📁 Loading contractors from Excel...');

    if (!fs.existsSync(CONFIG.excelFileName)) {
        throw new Error(`Excel file not found: ${CONFIG.excelFileName}`);
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

    const workbook = XLSX.readFile(CONFIG.excelFileName);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = rawData[0] || [];
    const codeIndex = 0;
    const firstNameIndex = findColumnIndex(headers, ['first name', 'firstname', 'forename']);
    const lastNameIndex = findColumnIndex(headers, ['last name', 'lastname', 'surname']);
    const departmentIndex = findColumnIndex(headers, ['department', 'dept']);

    if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
        throw new Error('Required columns not found');
    }

    const contractors = [];

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

                contractors.push({
                    code, fullName, firstName, lastName,
                    department: department || 'Security',
                    username, password: username
                });

                if (mode === 'single') {
                    console.log(`✅ Selected: ${fullName}\n`);
                    return contractors[0];
                }
            }
        }
    }

    if (contractors.length === 0) {
        throw new Error('No Security contractors found');
    }

    console.log(`✅ Found ${contractors.length} contractors\n`);
    return mode === 'all' ? contractors : contractors[0];
}

// === MAIN FUNCTION ===
async function runWithExistingProfile(mode = 'single') {
    let driver;

    try {
        console.log('🌟 USING YOUR EXISTING CHROME PROFILE\n');
        console.log('This preserves your Google login and all sessions\n');

        // Load contractors
        const contractors = mode === 'single' ?
            [loadContractorsFromExcel('single')] :
            loadContractorsFromExcel('all');

        // Connect to Chrome
        driver = await connectToYourChrome();
        const automator = new NaturalHumanAutomator(driver);

        // Guide user
        console.log('📍 NAVIGATION STEPS:');
        console.log('1. Go to: https://www.rivosafeguard.com/insight/');
        console.log('2. Your Google login should work automatically');
        console.log('3. Navigate to User Management → Create User\n');

        await waitForUserNavigation("Please navigate to the user creation form in Rivo Safeguard");

        // Process contractors
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            console.log(`\n${'═'.repeat(50)}`);
            console.log(`📊 Contractor ${i + 1} of ${contractors.length}`);
            console.log('═'.repeat(50));

            // Check for form
            let formFound = await detectUserCreationForm(driver);

            if (!formFound) {
                console.log('⚠️  Form not found');
                await waitForUserNavigation("Please navigate to the user creation form");
                formFound = await detectUserCreationForm(driver);
            }

            if (formFound) {
                await fillUserForm(driver, contractor, automator);

                if (i < contractors.length - 1) {
                    await waitForUserNavigation("Navigate back to create another user");
                }
            }
        }

        console.log('\n' + '═'.repeat(50));
        console.log('🎉 COMPLETED!');
        console.log('═'.repeat(50));
        console.log('✅ Your Chrome remains open with all sessions intact\n');

    } catch (error) {
        console.error('❌ Error:', error.message);
    }
}

// === ENTRY POINT ===
if (require.main === module) {
    const args = process.argv.slice(2);

    console.log('🔐 EXISTING CHROME PROFILE AUTOMATION\n');

    if (args.includes('--help')) {
        console.log('📚 Usage:');
        console.log('  node script.js        Single contractor');
        console.log('  node script.js --all  All contractors\n');

        console.log('📋 How it works:');
        console.log('1. You start Chrome with debugging enabled');
        console.log('2. Script connects to YOUR Chrome');
        console.log('3. Your Google login works automatically');
        console.log('4. You navigate, script fills forms\n');

    } else if (args.includes('--all')) {
        runWithExistingProfile('all');
    } else {
        runWithExistingProfile('single');
    }
}

module.exports = { runWithExistingProfile };