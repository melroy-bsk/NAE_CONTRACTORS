// CHROME PROFILE AUTOMATION WITH GOOGLE SSO
// Uses existing Chrome profile to maintain Google login sessions

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
    chromeProfile: 'Default', // Change to your profile name if different
    debugPort: 9222,
    useExistingChrome: true // Always use existing Chrome to maintain sessions
};

// === CHECK IF CHROME IS RUNNING ===
async function isChromeRunning() {
    return new Promise((resolve) => {
        exec('curl -s http://localhost:9222/json/version', (error, stdout) => {
            if (!error && stdout) {
                try {
                    JSON.parse(stdout);
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

// === GET CHROME EXECUTABLE PATH ===
function getChromePath() {
    const platform = os.platform();

    if (platform === 'darwin') {
        return '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
    } else if (platform === 'win32') {
        const paths = [
            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
            path.join(os.homedir(), 'AppData\\Local\\Google\\Chrome\\Application\\chrome.exe')
        ];

        for (const chromePath of paths) {
            if (fs.existsSync(chromePath)) {
                return chromePath;
            }
        }
        return 'chrome.exe'; // Fallback to PATH
    } else {
        // Linux
        const paths = ['/usr/bin/google-chrome', '/usr/bin/google-chrome-stable', '/usr/bin/chromium'];
        for (const chromePath of paths) {
            if (fs.existsSync(chromePath)) {
                return chromePath;
            }
        }
        return 'google-chrome'; // Fallback
    }
}

// === START CHROME WITH YOUR PROFILE ===
async function startChromeWithProfile() {
    console.log('🚀 Starting Chrome with your existing profile...\n');

    const chromePath = getChromePath();
    console.log(`📍 Chrome path: ${chromePath}`);

    // Get user data directory
    let userDataDir;
    if (os.platform() === 'darwin') {
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
    } else if (os.platform() === 'win32') {
        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else {
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    console.log(`📁 Profile location: ${userDataDir}`);
    console.log(`👤 Using profile: ${CONFIG.chromeProfile}`);

    // Build command based on platform
    let command;
    if (os.platform() === 'win32') {
        command = `"${chromePath}" --remote-debugging-port=${CONFIG.debugPort} --user-data-dir="${userDataDir}" --profile-directory="${CONFIG.chromeProfile}"`;
    } else {
        command = `"${chromePath}" --remote-debugging-port=${CONFIG.debugPort} --user-data-dir="${userDataDir}" --profile-directory="${CONFIG.chromeProfile}"`;
    }

    return new Promise((resolve, reject) => {
        console.log('📌 Starting Chrome...');

        exec(command, (error) => {
            // Chrome will continue running, so we don't wait for it to exit
            if (error && error.code !== null && error.code !== 0) {
                // Only reject on actual errors, not on Chrome staying open
                console.error('Error starting Chrome:', error);
            }
        });

        // Give Chrome time to start
        console.log('⏳ Waiting for Chrome to start...');
        setTimeout(async () => {
            const running = await isChromeRunning();
            if (running) {
                console.log('✅ Chrome started successfully!');
                console.log('🔐 Your Google sessions are preserved\n');
                resolve();
            } else {
                console.log('⚠️  Chrome may be starting slowly, waiting more...');
                setTimeout(resolve, 3000);
            }
        }, 3000);
    });
}

// === CONNECT TO CHROME WITH PROFILE ===
async function connectToChromeWithProfile() {
    console.log('🔌 Connecting to Chrome with your profile...\n');

    // First check if Chrome is running with debugging
    const isRunning = await isChromeRunning();

    if (!isRunning) {
        console.log('📌 Chrome not running with debugging port');
        console.log('🔄 Starting Chrome for you...\n');

        await startChromeWithProfile();

        // Extra wait to ensure Chrome is ready
        await new Promise(resolve => setTimeout(resolve, 2000));
    } else {
        console.log('✅ Found existing Chrome with debugging enabled\n');
    }

    // Now connect to Chrome
    try {
        const chromeOptions = new chrome.Options();
        chromeOptions.debuggerAddress(`localhost:${CONFIG.debugPort}`);

        console.log('🔗 Attempting connection...');

        const driver = await new Builder()
            .forBrowser('chrome')
            .setChromeOptions(chromeOptions)
            .build();

        // Verify connection by getting current URL
        try {
            const currentUrl = await driver.getCurrentUrl();
            console.log('✅ Successfully connected to Chrome!');
            console.log(`📍 Current page: ${currentUrl}\n`);
        } catch (e) {
            console.log('✅ Connected to Chrome!\n');
        }

        // Remove automation indicators
        await driver.executeScript(`
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined
            });
        `).catch(() => { }); // Ignore errors if script fails

        return driver;

    } catch (error) {
        console.error('❌ Failed to connect to Chrome:', error.message);
        console.log('\n💡 Troubleshooting tips:');
        console.log('1. Close ALL Chrome windows');
        console.log('2. Run the script again');
        console.log('3. The script will start Chrome with the correct settings\n');
        throw error;
    }
}

// === NATURAL HUMAN AUTOMATOR (Same as before) ===
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

        // Clear field
        await element.sendKeys(Key.chord(Key.CONTROL, "a"));
        await this.randomDelay(100, 300);
        await element.sendKeys(Key.DELETE);
        await this.randomDelay(200, 400);

        // Type naturally
        for (let i = 0; i < text.length; i++) {
            await element.sendKeys(text[i]);

            if (Math.random() < 0.1) {
                await this.driver.sleep(300 + Math.random() * 200);
            } else {
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

// === INTERACTIVE NAVIGATION ===
async function interactiveNavigation(message = "Navigate to the desired page") {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('\n' + '═'.repeat(60));
    console.log('🧭 MANUAL NAVIGATION REQUIRED');
    console.log('═'.repeat(60));
    console.log(`\n📋 ${message}`);
    console.log('\n💡 Important:');
    console.log('   • Your Google login session is preserved');
    console.log('   • Navigate normally - take your time');
    console.log('   • Complete any authentication if needed');
    console.log('   • The script will wait for you');
    console.log('\n');

    await new Promise(resolve => {
        rl.question('✅ Press Enter when ready to continue: ', () => {
            rl.close();
            resolve();
        });
    });

    console.log('\n✨ Continuing with automation...\n');
}

// === SMART NAVIGATION GUIDE ===
async function guidedNavigation() {
    console.log('\n📍 NAVIGATION GUIDE FOR RIVO SAFEGUARD');
    console.log('━'.repeat(50));
    console.log('\nTypical steps:');
    console.log('1. Go to: https://www.rivosafeguard.com/insight/');
    console.log('2. If prompted, log in with your Google account');
    console.log('3. Click the menu/app launcher (usually top-left)');
    console.log('4. Navigate to User Management or similar');
    console.log('5. Click "Create New User" or similar button');
    console.log('\n');

    await interactiveNavigation("Please navigate to the user creation form");
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

        // Check all iframes
        const frames = await driver.findElements(By.css('iframe'));
        console.log(`   Checking ${frames.length} iframe(s)...`);

        for (let i = 0; i < frames.length; i++) {
            try {
                await driver.switchTo().defaultContent();
                await driver.switchTo().frame(i);

                await driver.findElement(By.id("Username"));
                console.log(`✅ Form found in iframe ${i}`);
                return true;
            } catch (e) {
                // Continue checking
            }
        }

        await driver.switchTo().defaultContent();
        console.log('❌ Form not found');
        return false;

    } catch (error) {
        console.log('❌ Error detecting form:', error.message);
        return false;
    }
}

// === FILL USER FORM ===
async function fillUserFormNaturally(driver, contractor, automator) {
    console.log(`\n👤 Creating user: ${contractor.fullName}`);
    console.log(`📝 Username: ${contractor.username}`);

    try {
        // Fill basic fields
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

        // Optional fields
        console.log('\n🔧 Checking for optional fields...');

        try {
            await automator.selectDropdownOption(
                By.id("Attributes.Users.StatusOfEmployment"),
                "Current",
                "Employment Status"
            );
        } catch (e) {
            console.log('   ℹ️  Employment Status not available');
        }

        try {
            await automator.selectDropdownOption(
                By.id("UserTypeID"),
                "Limited access user",
                "User Type"
            );
        } catch (e) {
            console.log('   ℹ️  User Type not available');
        }

        // Review and confirm
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        console.log('\n' + '─'.repeat(50));
        console.log('📋 REVIEW USER DETAILS:');
        console.log(`   Name: ${contractor.fullName}`);
        console.log(`   Username: ${contractor.username}`);
        console.log(`   Department: ${contractor.department}`);
        console.log(`   Employee #: ${contractor.code}`);
        console.log('─'.repeat(50) + '\n');

        const shouldSave = await new Promise(resolve => {
            rl.question('💾 Save this user? (y/n): ', (answer) => {
                rl.close();
                resolve(answer);
            });
        });

        if (shouldSave.toLowerCase() === 'y') {
            console.log('\n💾 Saving user...');
            await automator.waitAndClick(By.name("save"), "Save button");
            await automator.randomDelay(3000, 5000);

            console.log(`✅ Successfully created: ${contractor.username}`);
            return true;
        } else {
            console.log('🚫 User creation skipped');
            return false;
        }

    } catch (error) {
        console.error(`❌ Error: ${error.message}`);
        return false;
    }
}

// === LOAD CONTRACTORS FROM EXCEL ===
function loadContractorsFromExcel(mode = 'single') {
    console.log('📁 Loading contractor data...');

    if (!fs.existsSync(CONFIG.excelFileName)) {
        throw new Error(`❌ Excel file not found: ${CONFIG.excelFileName}`);
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
        throw new Error('❌ Required columns not found');
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
        throw new Error('❌ No Security contractors found');
    }

    console.log(`✅ Found ${contractors.length} Security contractors\n`);
    return mode === 'all' ? contractors : contractors[0];
}

// === MAIN AUTOMATION ===
async function runChromeProfileAutomation(mode = 'single') {
    let driver;

    try {
        console.log('🌟 CHROME PROFILE AUTOMATION WITH GOOGLE SSO\n');
        console.log('This script preserves your Google login sessions\n');

        // Load contractors
        const contractors = mode === 'single' ?
            [loadContractorsFromExcel('single')] :
            loadContractorsFromExcel('all');

        // Connect to Chrome with profile
        driver = await connectToChromeWithProfile();
        const automator = new NaturalHumanAutomator(driver);

        // Initial navigation guide
        await guidedNavigation();

        // Process each contractor
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            console.log(`\n${'═'.repeat(60)}`);
            console.log(`📊 Processing contractor ${i + 1} of ${contractors.length}`);
            console.log('═'.repeat(60));

            // Detect form
            let formFound = await detectUserCreationForm(driver);

            if (!formFound) {
                console.log('\n⚠️  User creation form not detected');
                await interactiveNavigation("Please navigate to the user creation form");

                formFound = await detectUserCreationForm(driver);
                if (!formFound) {
                    console.log('❌ Could not find form, skipping this contractor');
                    continue;
                }
            }

            // Fill form
            const success = await fillUserFormNaturally(driver, contractor, automator);

            if (success && i < contractors.length - 1) {
                await interactiveNavigation("Navigate back to create another user");
            }
        }

        // Completion
        console.log('\n' + '═'.repeat(60));
        console.log('🎉 AUTOMATION COMPLETED!');
        console.log('═'.repeat(60));
        console.log(`\n✅ Processed ${contractors.length} contractor(s)`);
        console.log('📌 Chrome remains open with your session intact');

    } catch (error) {
        console.error('\n❌ Error:', error.message);

        if (error.message.includes('connect')) {
            console.log('\n💡 Connection Tips:');
            console.log('1. Close ALL Chrome windows');
            console.log('2. Run this script again');
            console.log('3. Let the script start Chrome for you');
        }
    }

    console.log('\n👋 Script finished. Chrome remains under your control.\n');
}

// === COMMAND LINE INTERFACE ===
if (require.main === module) {
    const args = process.argv.slice(2);

    console.log('🔐 CHROME PROFILE AUTOMATION\n');

    if (args.includes('--help')) {
        console.log('📚 Usage:');
        console.log('  node script.js           Process single contractor');
        console.log('  node script.js --all     Process all contractors');
        console.log('  node script.js --help    Show this help\n');

        console.log('✨ Features:');
        console.log('  • Uses your existing Chrome profile');
        console.log('  • Preserves Google login sessions');
        console.log('  • Natural human-like automation');
        console.log('  • Interactive navigation prompts\n');

        console.log('🔧 Configuration:');
        console.log('  Edit CONFIG.chromeProfile if using non-default profile');
        console.log('  Current profile:', CONFIG.chromeProfile);

    } else if (args.includes('--all')) {
        console.log('👥 Mode: ALL contractors\n');
        runChromeProfileAutomation('all');

    } else {
        console.log('👤 Mode: SINGLE contractor\n');
        runChromeProfileAutomation('single');
    }
}

module.exports = {
    runChromeProfileAutomation,
    connectToChromeWithProfile,
    NaturalHumanAutomator
};