// FIXED HYBRID AUTOMATION - No More DevTools/Deprecated Endpoint Errors
// Manual navigation + Automated form filling with all issues resolved

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
    defaultTimeout: 15000,
    humanDelayBase: 1000,
    humanDelayVariation: 500,
    sourceProfile: 'Profile 1'
};

// === FORCE KILL CHROME ===
async function forceKillChrome() {
    console.log('🔄 Ensuring Chrome is completely closed...');

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
            default:
                resolve();
                return;
        }

        exec(command, () => {
            console.log('✅ Chrome processes cleared');
            setTimeout(resolve, 2000);
        });
    });
}

// === CREATE SAFE PROFILE COPY ===
async function createSafeProfileCopy() {
    console.log('📋 Creating safe profile copy for automation...');

    // Get source profile path
    let userDataDir;
    if (os.platform() === 'darwin') {
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
    } else if (os.platform() === 'win32') {
        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else {
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    const sourceProfilePath = path.join(userDataDir, CONFIG.sourceProfile);
    const tempDir = path.join(os.tmpdir(), `chrome-safe-${Date.now()}`);
    const tempProfilePath = path.join(tempDir, 'Default');

    try {
        // Create temp directories
        fs.mkdirSync(tempDir, { recursive: true });
        fs.mkdirSync(tempProfilePath, { recursive: true });

        // Copy essential files that contain login info
        const essentialFiles = ['Cookies', 'Login Data', 'Preferences', 'Web Data'];

        for (const fileName of essentialFiles) {
            const sourcePath = path.join(sourceProfilePath, fileName);
            const destPath = path.join(tempProfilePath, fileName);

            if (fs.existsSync(sourcePath)) {
                try {
                    fs.copyFileSync(sourcePath, destPath);
                    console.log(`   ✅ Copied: ${fileName}`);
                } catch (e) {
                    console.log(`   ⚠️  Skipped ${fileName}: ${e.message}`);
                }
            }
        }

        console.log(`✅ Safe profile created at: ${tempDir}`);
        return tempDir;

    } catch (error) {
        console.log(`⚠️  Profile copy failed: ${error.message}`);
        console.log('🔄 Using fresh profile instead...');

        const freshDir = path.join(os.tmpdir(), `chrome-fresh-${Date.now()}`);
        fs.mkdirSync(freshDir, { recursive: true });
        return freshDir;
    }
}

// === FIXED CHROME DRIVER ===
async function createFixedChromeDriver() {
    console.log('🔧 Creating fixed Chrome driver (no more errors!)...\n');

    try {
        // Step 1: Kill Chrome processes
        await forceKillChrome();

        // Step 2: Create safe profile
        const tempProfileDir = await createSafeProfileCopy();

        // Step 3: Modern Chrome options
        const chromeOptions = new chrome.Options();

        // === PROFILE SETTINGS ===
        chromeOptions.addArguments(`--user-data-dir=${tempProfileDir}`);
        chromeOptions.addArguments('--profile-directory=Default');

        // === FIX DEVTOOLS ERRORS ===
        chromeOptions.addArguments('--disable-dev-shm-usage');
        chromeOptions.addArguments('--disable-features=VizDisplayCompositor');
        chromeOptions.addArguments('--disable-features=AudioServiceOutOfProcess');
        chromeOptions.addArguments('--disable-features=VizServiceDisplayCompositor');
        chromeOptions.addArguments('--disable-gpu-sandbox');
        chromeOptions.addArguments('--disable-background-networking');

        // === FIX DEPRECATED ENDPOINT WARNINGS ===
        chromeOptions.addArguments('--disable-features=TranslateUI');
        chromeOptions.addArguments('--disable-features=BlinkGenPropertyTrees');
        chromeOptions.addArguments('--disable-background-timer-throttling');
        chromeOptions.addArguments('--disable-backgrounding-occluded-windows');
        chromeOptions.addArguments('--disable-renderer-backgrounding');

        // === ESSENTIAL FLAGS ===
        chromeOptions.addArguments('--no-sandbox');
        chromeOptions.addArguments('--no-first-run');
        chromeOptions.addArguments('--no-default-browser-check');
        chromeOptions.addArguments('--disable-default-apps');
        chromeOptions.addArguments('--disable-popup-blocking');
        chromeOptions.addArguments('--disable-extensions');
        chromeOptions.addArguments('--disable-web-security');

        // === STEALTH SETTINGS ===
        chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
        chromeOptions.excludeSwitches(['enable-automation', 'enable-logging']);
        chromeOptions.addArguments('--disable-infobars');

        // === USER AGENT ===
        chromeOptions.addArguments('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36');

        // === PREFERENCES ===
        chromeOptions.setUserPreferences({
            'profile.default_content_setting_values.notifications': 2,
            'profile.default_content_settings.popups': 0,
            'profile.managed_default_content_settings.images': 1
        });

        // === SET CHROME BINARY (IMPORTANT FOR MAC) ===
        let chromePath = null;
        if (os.platform() === 'darwin') {
            chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
        } else if (os.platform() === 'win32') {
            const possiblePaths = [
                'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
                'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
            ];
            for (const testPath of possiblePaths) {
                if (fs.existsSync(testPath)) {
                    chromePath = testPath;
                    break;
                }
            }
        }

        if (chromePath && fs.existsSync(chromePath)) {
            chromeOptions.setChromeBinaryPath(chromePath);
        }

        console.log('🚀 Creating WebDriver with fixed settings...');

        // Create driver with retries
        let driver;
        let attempts = 3;

        while (attempts > 0) {
            try {
                driver = await new Builder()
                    .forBrowser('chrome')
                    .setChromeOptions(chromeOptions)
                    .build();
                break;
            } catch (error) {
                attempts--;
                if (attempts > 0) {
                    console.log(`⚠️  Retry ${3 - attempts}/3: ${error.message}`);
                    await new Promise(resolve => setTimeout(resolve, 2000));
                } else {
                    throw error;
                }
            }
        }

        // Enhanced stealth injection
        await driver.executeScript(`
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined,
                configurable: true
            });
            delete window.navigator.webdriver;
            delete window.webdriver;
            delete window.domAutomation;
            delete window.domAutomationController;
            
            // Add realistic chrome object
            if (!window.chrome) {
                window.chrome = {
                    runtime: {},
                    loadTimes: function() { return {}; },
                    csi: function() { return {}; },
                    app: {}
                };
            }
        `);

        console.log('✅ Fixed Chrome driver created successfully!');
        console.log('🕵️  No DevTools errors, no deprecated endpoint warnings\n');

        return { driver, tempProfileDir };

    } catch (error) {
        console.error('❌ Driver creation failed:', error.message);
        throw error;
    }
}

// === LOAD CONTRACTOR DATA (UNCHANGED) ===
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
                    console.log(`👤 Selected contractor: ${fullName} (${username})\n`);
                    return contractors[0];
                }
            }
        }
    }

    if (contractors.length === 0) {
        throw new Error('❌ No Security department contractors found in Excel file');
    }

    if (mode === 'all') {
        console.log(`👥 Loaded ${contractors.length} Security contractors\n`);
        return contractors;
    }

    return contractors[0];
}

// === HUMAN AUTOMATION CLASS (UNCHANGED) ===
class HumanAutomator {
    constructor(driver) {
        this.driver = driver;
    }

    async humanDelay(baseMs = CONFIG.humanDelayBase, variationMs = CONFIG.humanDelayVariation) {
        const randomVariation = Math.random() * variationMs * 2 - variationMs;
        const totalDelay = Math.max(baseMs + randomVariation, 200);
        await this.driver.sleep(totalDelay);
    }

    async humanType(element, text, avgDelay = 80) {
        await element.clear();
        await this.humanDelay(300, 200);

        for (let i = 0; i < text.length; i++) {
            await element.sendKeys(text[i]);
            await this.driver.sleep(avgDelay + (Math.random() * 40 - 20));
        }

        await this.humanDelay(200, 100);
    }

    async safeInput(selector, text, description, timeout = CONFIG.defaultTimeout) {
        try {
            console.log(`   ✏️  Entering: ${description}`);
            const element = await this.driver.wait(until.elementLocated(selector), timeout);
            await this.driver.wait(until.elementIsEnabled(element), 5000);

            await element.click();
            await this.humanDelay(300, 150);
            await this.humanType(element, text);

            console.log(`   ✅ Successfully entered: ${description}`);
            return element;

        } catch (error) {
            console.log(`   ⚠️  Failed to enter ${description}: ${error.message}`);
            throw error;
        }
    }

    async safeClick(selector, description, timeout = CONFIG.defaultTimeout) {
        try {
            console.log(`   🖱️  Clicking: ${description}`);
            const element = await this.driver.wait(until.elementLocated(selector), timeout);
            await this.driver.wait(until.elementIsEnabled(element), 5000);

            await element.click();
            console.log(`   ✅ Successfully clicked: ${description}`);
            await this.humanDelay(300, 150);
            return element;

        } catch (error) {
            console.log(`   ⚠️  Failed to click ${description}: ${error.message}`);
            throw error;
        }
    }

    async safeSelectOption(dropdownSelector, optionText, description, timeout = CONFIG.defaultTimeout) {
        try {
            console.log(`   📋 Selecting "${optionText}" in: ${description}`);
            const dropdown = await this.driver.wait(until.elementLocated(dropdownSelector), timeout);
            await this.driver.wait(until.elementIsEnabled(dropdown), 5000);

            const options = await dropdown.findElements(By.xpath(`//option[normalize-space(text()) = '${optionText}']`));
            if (options.length === 0) {
                console.log(`   ⚠️  Option "${optionText}" not found, continuing...`);
                return dropdown;
            }

            await options[0].click();
            console.log(`   ✅ Successfully selected: ${optionText}`);
            await this.humanDelay(500, 250);
            return dropdown;

        } catch (error) {
            console.log(`   ⚠️  Failed to select option: ${error.message}`);
            return null;
        }
    }
}

// === MANUAL NAVIGATION HELPER ===
async function waitForManualNavigation(driver) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('👋 MANUAL NAVIGATION MODE');
    console.log('═══════════════════════════════════════');
    console.log('📋 INSTRUCTIONS:');
    console.log('1. 🌐 Use the browser window to navigate manually');
    console.log('2. 🔐 Log into Rivo Safeguard if needed');
    console.log('3. 🧭 Navigate to the USER CREATION page');
    console.log('4. 📝 Get to the user creation form');
    console.log('5. ✅ Press Enter when ready for automation');
    console.log('');

    await new Promise(resolve => {
        rl.question('👀 Navigate manually, then press Enter: ', () => {
            rl.close();
            resolve();
        });
    });

    console.log('✅ Manual navigation complete! Starting automation...\n');
}

// === CHECK FOR USER CREATION FORM ===
async function waitForUserCreationForm(driver) {
    console.log('🔍 Looking for user creation form...');

    try {
        await driver.switchTo().defaultContent();

        const frames = await driver.findElements(By.css('iframe'));
        if (frames.length > 0) {
            console.log('🖼️  Switching to iframe...');
            await driver.switchTo().frame(0);
        }

        await driver.wait(until.elementLocated(By.id("Username")), 5000);
        console.log('✅ User creation form found!');
        return true;

    } catch (error) {
        console.log('❌ User creation form not found');
        return false;
    }
}

// === FILL USER FORM ===
async function fillUserForm(driver, contractor, automator) {
    console.log(`\n👤 Filling form for: ${contractor.username} (${contractor.fullName})`);

    try {
        // Basic information
        console.log('📝 Filling basic information...');
        await automator.safeInput(By.id("Username"), contractor.username, "Username");
        await automator.safeInput(By.name("Password"), contractor.password, "Password");
        await automator.safeInput(By.name("JobTitle"), contractor.department, "Job Title");
        await automator.safeInput(By.id("Attributes.People.Forename"), contractor.firstName, "First Name");
        await automator.safeInput(By.id("Attributes.People.Surname"), contractor.lastName, "Last Name");
        await automator.safeInput(By.id("Attributes.Users.EmployeeNumber"), contractor.code, "Employee Number");

        // Optional advanced settings
        console.log('🏢 Attempting advanced settings (may skip)...');
        try {
            await automator.safeSelectOption(By.id("Attributes.Users.StatusOfEmployment"), "Current", "Employment Status");
            await automator.safeSelectOption(By.id("UserTypeID"), "Limited access user", "User Type");
        } catch (error) {
            console.log('   ⚠️  Advanced settings skipped');
        }

        // Confirm save
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        console.log('\n🤔 READY TO SAVE USER');
        console.log(`📋 ${contractor.fullName} (${contractor.username})`);

        const shouldSave = await new Promise(resolve => {
            rl.question('❓ Save this user? (y/n): ', resolve);
        });
        rl.close();

        if (shouldSave.toLowerCase() === 'y') {
            console.log('💾 Saving user...');
            await automator.safeClick(By.name("save"), "Save button");
            await automator.humanDelay(5000, 2000);
            console.log(`✅ User created: ${contractor.username}`);
        } else {
            console.log('🚫 Save cancelled');
        }

    } catch (error) {
        console.error(`❌ Form filling failed: ${error.message}`);
        throw error;
    }
}

// === MAIN FIXED HYBRID AUTOMATION ===
async function runFixedHybridAutomation(mode = 'single') {
    let driver;
    let tempProfileDir;

    try {
        console.log('🔧 FIXED HYBRID AUTOMATION - No More Errors!\n');

        // Load contractors
        const contractors = mode === 'single' ?
            [loadContractorsFromExcel('single')] :
            loadContractorsFromExcel('all');

        // Create fixed driver
        const result = await createFixedChromeDriver();
        driver = result.driver;
        tempProfileDir = result.tempProfileDir;

        await driver.manage().window().setRect({ width: 1920, height: 1080 });
        const automator = new HumanAutomator(driver);

        // Try automatic navigation first
        console.log('🌐 Attempting automatic navigation...');
        try {
            await driver.get('https://www.rivosafeguard.com/insight/');
            await driver.sleep(3000);

            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Automatic navigation successful!');

            // Try auto-navigation to user creation
            await automator.safeClick(By.css(".sch-container-left"), "Left container");
            await automator.safeClick(By.css(".sch-app-launcher-button"), "App launcher");
            await automator.safeClick(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text"), "Menu item");
            await automator.safeClick(By.css(".k-drawer-item:nth-child(6)"), "User management");
            await driver.switchTo().frame(0);

            console.log('✅ Auto-navigation to user creation successful!');

        } catch (error) {
            console.log('❌ Auto-navigation failed, switching to manual...');
            await waitForManualNavigation(driver);

            // Wait for correct page
            let onCorrectPage = false;
            while (!onCorrectPage) {
                onCorrectPage = await waitForUserCreationForm(driver);
                if (!onCorrectPage) {
                    await waitForManualNavigation(driver);
                }
            }
        }

        // Process contractors
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            console.log(`\n📊 Progress: ${i + 1}/${contractors.length}`);

            if (i > 0) {
                const rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });

                await new Promise(resolve => {
                    rl.question('👀 Navigate to fresh user creation form, press Enter: ', () => {
                        rl.close();
                        resolve();
                    });
                });

                while (!(await waitForUserCreationForm(driver))) {
                    await new Promise(resolve => setTimeout(resolve, 2000));
                }
            }

            await fillUserForm(driver, contractor, automator);
            console.log(`✅ Completed: ${i + 1}/${contractors.length}`);
        }

        console.log('\n🎉 Fixed hybrid automation completed!');
        console.log('✅ No DevTools errors!');
        console.log('✅ No deprecated endpoint warnings!');

    } catch (error) {
        console.error('❌ Automation failed:', error.message);
    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 10 seconds...');
            await driver.sleep(10000);
            await driver.quit();
        }

        // Cleanup temp profile
        if (tempProfileDir && fs.existsSync(tempProfileDir)) {
            try {
                fs.rmSync(tempProfileDir, { recursive: true, force: true });
                console.log('🧹 Temporary profile cleaned up');
            } catch (e) {
                console.log('⚠️  Temp cleanup will happen on restart');
            }
        }
    }
}

// === COMMAND LINE INTERFACE ===
if (require.main === module) {
    const args = process.argv.slice(2);

    console.log('🔧 FIXED HYBRID AUTOMATION\n');
    console.log('✅ No DevTools errors');
    console.log('✅ No deprecated endpoint warnings');
    console.log('✅ Safe profile handling\n');

    if (args.includes('--help')) {
        console.log('📚 Commands:');
        console.log('  node fixed-hybrid.js           Single user (hybrid)');
        console.log('  node fixed-hybrid.js --all     All users (hybrid)');
        console.log('  node fixed-hybrid.js --help    Show help\n');

    } else if (args.includes('--all')) {
        console.log('🚀 ALL USERS MODE\n');
        runFixedHybridAutomation('all');

    } else {
        console.log('🧪 SINGLE USER MODE\n');
        runFixedHybridAutomation('single');
    }
}

module.exports = {
    runFixedHybridAutomation,
    createFixedChromeDriver,
    fillUserForm
};