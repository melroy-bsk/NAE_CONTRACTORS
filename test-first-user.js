// NATURAL USER-LIKE AUTOMATION
// Uses real Chrome profile and mimics human behavior

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
    profileMethod: 'remote-debugging', // 'remote-debugging' or 'profile-copy'
    debugPort: 9222
};

// === CHECK IF CHROME IS RUNNING ===
async function isChromeRunning() {
    return new Promise((resolve) => {
        let command;

        switch (os.platform()) {
            case 'win32':
                command = 'tasklist /FI "IMAGENAME eq chrome.exe" 2>nul | find /I "chrome.exe" >nul';
                break;
            case 'darwin':
                command = 'pgrep -x "Google Chrome" > /dev/null';
                break;
            case 'linux':
                command = 'pgrep -x chrome > /dev/null';
                break;
        }

        exec(command, (error) => {
            resolve(!error);
        });
    });
}

// === KILL ALL CHROME PROCESSES ===
async function killAllChrome() {
    console.log('🔄 Closing any existing Chrome instances...');

    return new Promise((resolve) => {
        let command;

        switch (os.platform()) {
            case 'win32':
                command = 'taskkill /F /IM chrome.exe /T 2>nul';
                break;
            case 'darwin':
                command = 'pkill -f "Google Chrome" 2>/dev/null';
                break;
            case 'linux':
                command = 'pkill -f chrome 2>/dev/null';
                break;
        }

        exec(command, () => {
            setTimeout(resolve, 2000);
        });
    });
}

// === GET CHROME PATHS ===
function getChromePaths() {
    let chromePath, userDataDir;

    if (os.platform() === 'darwin') {
        chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
    } else if (os.platform() === 'win32') {
        chromePath = 'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe';
        if (!fs.existsSync(chromePath)) {
            chromePath = 'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe';
        }
        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else {
        chromePath = '/usr/bin/google-chrome';
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    return { chromePath, userDataDir };
}

// === START CHROME WITH REMOTE DEBUGGING ===
async function startChromeWithRemoteDebugging() {
    console.log('🚀 Starting Chrome with remote debugging...\n');

    const { chromePath, userDataDir } = getChromePaths();

    // Kill any existing Chrome
    await killAllChrome();

    // Start Chrome with remote debugging
    const chromeArgs = [
        `--remote-debugging-port=${CONFIG.debugPort}`,
        `--user-data-dir=${userDataDir}`,
        '--no-first-run',
        '--no-default-browser-check'
    ];

    return new Promise((resolve, reject) => {
        exec(`"${chromePath}" ${chromeArgs.join(' ')}`, (error) => {
            if (error && error.code !== null) {
                reject(error);
            }
        });

        // Give Chrome time to start
        setTimeout(() => {
            console.log('✅ Chrome started with remote debugging on port', CONFIG.debugPort);
            console.log('📌 You can now use the browser normally\n');
            resolve();
        }, 3000);
    });
}

// === CONNECT TO EXISTING CHROME ===
async function connectToExistingChrome() {
    console.log('🔌 Connecting to existing Chrome instance...\n');

    const chromeOptions = new chrome.Options();
    chromeOptions.addArguments(`--remote-debugging-port=${CONFIG.debugPort}`);
    chromeOptions.debuggerAddress(`127.0.0.1:${CONFIG.debugPort}`);

    try {
        const driver = await new Builder()
            .forBrowser('chrome')
            .setChromeOptions(chromeOptions)
            .build();

        console.log('✅ Successfully connected to Chrome!\n');
        return driver;

    } catch (error) {
        console.error('❌ Failed to connect:', error.message);
        console.log('\n💡 Try these steps:');
        console.log('1. Close all Chrome windows');
        console.log('2. Run the script again');
        console.log('3. Let the script start Chrome for you\n');
        throw error;
    }
}

// === CREATE PROFILE COPY METHOD ===
async function createProfileCopyDriver() {
    console.log('📁 Creating Chrome with profile copy...\n');

    const { userDataDir } = getChromePaths();
    const tempDir = path.join(os.tmpdir(), `chrome-automation-${Date.now()}`);

    // Create temp directory
    fs.mkdirSync(tempDir, { recursive: true });

    // Copy only essential files
    const filesToCopy = ['Cookies', 'Preferences', 'Local State'];
    const sourceDir = path.join(userDataDir, 'Default');
    const destDir = path.join(tempDir, 'Default');

    fs.mkdirSync(destDir, { recursive: true });

    for (const file of filesToCopy) {
        const src = path.join(sourceDir, file);
        const dest = path.join(destDir, file);

        if (fs.existsSync(src)) {
            try {
                fs.copyFileSync(src, dest);
                console.log(`   ✅ Copied ${file}`);
            } catch (e) {
                console.log(`   ⚠️  Skipped ${file}`);
            }
        }
    }

    const chromeOptions = new chrome.Options();
    chromeOptions.addArguments(`--user-data-dir=${tempDir}`);
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--start-maximized');

    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    console.log('✅ Chrome started with profile copy\n');
    return { driver, tempDir };
}

// === CREATE NATURAL CHROME DRIVER ===
async function createNaturalChromeDriver() {
    console.log('🌟 Setting up natural Chrome automation...\n');

    try {
        if (CONFIG.profileMethod === 'remote-debugging') {
            // Check if Chrome is already running with debugging
            try {
                return await connectToExistingChrome();
            } catch (e) {
                // Chrome not running with debugging, start it
                console.log('📌 Chrome not running with remote debugging');
                console.log('🔄 Starting Chrome for you...\n');

                await startChromeWithRemoteDebugging();

                // Wait and then connect
                await new Promise(resolve => setTimeout(resolve, 2000));
                return await connectToExistingChrome();
            }
        } else {
            // Use profile copy method
            const result = await createProfileCopyDriver();
            return result.driver;
        }

    } catch (error) {
        console.error('❌ Setup failed:', error.message);
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
        // Click the element first
        await element.click();
        await this.randomDelay(200, 500);

        // Clear existing content naturally
        await element.sendKeys(Key.chord(Key.CONTROL, "a"));
        await this.randomDelay(100, 300);
        await element.sendKeys(Key.DELETE);
        await this.randomDelay(200, 400);

        // Type character by character with natural rhythm
        for (let i = 0; i < text.length; i++) {
            await element.sendKeys(text[i]);

            // Natural typing rhythm
            if (Math.random() < 0.1) {
                // Occasional pause (thinking)
                await this.driver.sleep(300 + Math.random() * 200);
            } else {
                // Normal typing speed
                await this.driver.sleep(50 + Math.random() * 100);
            }
        }

        await this.randomDelay(300, 600);
    }

    async naturalClick(element) {
        // Scroll element into view
        await this.driver.executeScript("arguments[0].scrollIntoView({behavior: 'smooth', block: 'center'});", element);
        await this.randomDelay(500, 1000);

        // Move to element and click
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

            // Find and click the option
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
async function interactiveNavigation(driver, message = "Navigate to the desired page") {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('\n' + '═'.repeat(50));
    console.log('🧭 MANUAL NAVIGATION MODE');
    console.log('═'.repeat(50));
    console.log(`\n📋 ${message}`);
    console.log('\n💡 Tips:');
    console.log('   • Take your time - act like a normal user');
    console.log('   • Complete any login steps needed');
    console.log('   • Navigate naturally through the interface');
    console.log('   • The automation will wait for you');
    console.log('\n');

    await new Promise(resolve => {
        rl.question('✅ Press Enter when ready to continue: ', () => {
            rl.close();
            resolve();
        });
    });

    console.log('\n✨ Continuing with automation...\n');
}

// === DETECT USER CREATION FORM ===
async function detectUserCreationForm(driver) {
    console.log('🔍 Checking for user creation form...');

    try {
        // Check all frames
        await driver.switchTo().defaultContent();

        // Try main content first
        try {
            await driver.findElement(By.id("Username"));
            console.log('✅ Form found in main content');
            return true;
        } catch (e) {
            // Not in main content, check frames
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
                // Not in this frame
            }
        }

        await driver.switchTo().defaultContent();
        return false;

    } catch (error) {
        console.log('❌ Form not detected');
        return false;
    }
}

// === FILL USER FORM NATURALLY ===
async function fillUserFormNaturally(driver, contractor, automator) {
    console.log(`\n👤 Creating user: ${contractor.fullName}`);
    console.log(`📝 Username: ${contractor.username}`);

    try {
        // Basic fields with natural delays between each
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
        console.log('\n🔧 Attempting optional fields...');

        try {
            await automator.selectDropdownOption(
                By.id("Attributes.Users.StatusOfEmployment"),
                "Current",
                "Employment Status"
            );
            await automator.randomDelay(1000, 2000);
        } catch (e) {
            console.log('   ℹ️  Employment Status field not available');
        }

        try {
            await automator.selectDropdownOption(
                By.id("UserTypeID"),
                "Limited access user",
                "User Type"
            );
            await automator.randomDelay(1000, 2000);
        } catch (e) {
            console.log('   ℹ️  User Type field not available');
        }

        // Confirm before saving
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        console.log('\n' + '─'.repeat(40));
        console.log('📋 REVIEW USER DETAILS:');
        console.log(`   Name: ${contractor.fullName}`);
        console.log(`   Username: ${contractor.username}`);
        console.log(`   Department: ${contractor.department}`);
        console.log('─'.repeat(40) + '\n');

        const shouldSave = await new Promise(resolve => {
            rl.question('💾 Save this user? (y/n): ', (answer) => {
                rl.close();
                resolve(answer);
            });
        });

        if (shouldSave.toLowerCase() === 'y') {
            console.log('\n💾 Saving user...');
            await automator.waitAndClick(By.name("save"), "Save button");

            // Wait for save to complete
            await automator.randomDelay(3000, 5000);

            console.log(`\n✅ Successfully created user: ${contractor.username}`);
            return true;
        } else {
            console.log('\n🚫 User creation cancelled');
            return false;
        }

    } catch (error) {
        console.error(`\n❌ Error creating user: ${error.message}`);
        return false;
    }
}

// === LOAD CONTRACTORS (unchanged) ===
function loadContractorsFromExcel(mode = 'single') {
    console.log('📁 Loading contractor data from Excel...');

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
        throw new Error('❌ Required columns not found. Need: First Name, Last Name, Department');
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
                    console.log(`✅ Found contractor: ${fullName}\n`);
                    return contractors[0];
                }
            }
        }
    }

    if (contractors.length === 0) {
        throw new Error('❌ No Security department contractors found in Excel file');
    }

    if (mode === 'all') {
        console.log(`✅ Found ${contractors.length} Security contractors\n`);
        return contractors;
    }

    return contractors[0];
}

// === MAIN NATURAL AUTOMATION ===
async function runNaturalAutomation(mode = 'single') {
    let driver;
    let tempDir;

    try {
        console.log('🌟 NATURAL USER AUTOMATION\n');
        console.log(`📋 Using method: ${CONFIG.profileMethod}\n`);

        // Load contractor data
        const contractors = mode === 'single' ?
            [loadContractorsFromExcel('single')] :
            loadContractorsFromExcel('all');

        // Create natural driver
        if (CONFIG.profileMethod === 'remote-debugging') {
            driver = await createNaturalChromeDriver();
        } else {
            const result = await createProfileCopyDriver();
            driver = result.driver;
            tempDir = result.tempDir;
        }

        const automator = new NaturalHumanAutomator(driver);

        // Initial navigation
        await interactiveNavigation(driver,
            "Please navigate to Rivo Safeguard and log in"
        );

        // Process each contractor
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            console.log(`\n${'═'.repeat(50)}`);
            console.log(`📊 Processing ${i + 1} of ${contractors.length}`);
            console.log('═'.repeat(50));

            // Ensure we're on the user creation form
            let formFound = false;
            let attempts = 0;

            while (!formFound && attempts < 3) {
                formFound = await detectUserCreationForm(driver);

                if (!formFound) {
                    if (attempts === 0) {
                        console.log('\n⚠️  User creation form not detected');
                    }

                    await interactiveNavigation(driver,
                        "Please navigate to the user creation form"
                    );

                    attempts++;
                }
            }

            if (!formFound) {
                console.log('❌ Could not find user creation form after 3 attempts');
                continue;
            }

            // Fill the form naturally
            const success = await fillUserFormNaturally(driver, contractor, automator);

            if (success && i < contractors.length - 1) {
                await interactiveNavigation(driver,
                    "Please navigate back to create another user"
                );
            }
        }

        console.log('\n' + '═'.repeat(50));
        console.log('🎉 AUTOMATION COMPLETED');
        console.log('═'.repeat(50));
        console.log(`\n✅ Processed ${contractors.length} contractor(s)`);

    } catch (error) {
        console.error('\n❌ Automation error:', error.message);
        console.error(error.stack);
    } finally {
        if (driver && CONFIG.profileMethod !== 'remote-debugging') {
            console.log('\n⏰ Browser will remain open for review...');
            console.log('📌 Close the browser manually when done\n');

            // Keep browser open for profile copy method
            await new Promise(() => { });
        } else if (CONFIG.profileMethod === 'remote-debugging') {
            console.log('\n📌 Browser remains under your control');
            console.log('✅ You can continue using Chrome normally\n');
        }

        // Cleanup temp directory if used
        if (tempDir && fs.existsSync(tempDir)) {
            try {
                fs.rmSync(tempDir, { recursive: true, force: true });
            } catch (e) {
                // Ignore cleanup errors
            }
        }
    }
}

// === COMMAND LINE INTERFACE ===
if (require.main === module) {
    const args = process.argv.slice(2);

    console.log('🌟 NATURAL USER AUTOMATION\n');

    if (args.includes('--help')) {
        console.log('📚 Usage:');
        console.log('  node natural-automation.js         Process single user');
        console.log('  node natural-automation.js --all   Process all users');
        console.log('  node natural-automation.js --help  Show this help\n');

        console.log('📋 Features:');
        console.log('  • Two methods: remote-debugging or profile-copy');
        console.log('  • Natural typing and clicking patterns');
        console.log('  • Interactive navigation prompts');
        console.log('  • Human-like delays and behavior\n');

        console.log('🔧 Configuration:');
        console.log('  Edit CONFIG.profileMethod to switch between:');
        console.log('  - "remote-debugging": Connect to existing Chrome');
        console.log('  - "profile-copy": Create a copy of your profile\n');

    } else if (args.includes('--all')) {
        console.log('👥 Processing ALL Security contractors\n');
        runNaturalAutomation('all');

    } else {
        console.log('👤 Processing SINGLE contractor\n');
        runNaturalAutomation('single');
    }
}

module.exports = {
    runNaturalAutomation,
    createNaturalChromeDriver,
    NaturalHumanAutomator
};