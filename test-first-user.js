// HYBRID AUTOMATION - Manual Navigation + Auto Form Filling
// Perfect solution: You navigate manually, automation handles form filling

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs');
const readline = require('readline');

// === CONFIGURATION ===
const CONFIG = {
    excelFileName: 'Copy of Contractors staff list.xlsx',
    defaultTimeout: 15000,
    humanDelayBase: 1000,
    humanDelayVariation: 500,
    defaultProfile: 'Profile 1'
};

// === WORKING PROFILE DRIVER (FROM YOUR SUCCESSFUL TEST) ===
async function createWorkingProfileDriver(profileName = CONFIG.defaultProfile) {
    console.log(`🔧 Creating Chrome driver with Profile: "${profileName}"...`);

    const chromeOptions = new chrome.Options();

    // Cross-platform paths
    let chromePath = null;
    let userDataDir = null;

    if (os.platform() === 'darwin') {
        chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
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
        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else {
        chromePath = 'google-chrome';
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    // Set Chrome binary
    if (chromePath && chromePath !== 'google-chrome') {
        chromeOptions.setChromeBinaryPath(chromePath);
    }

    // Profile settings
    chromeOptions.addArguments(`--user-data-dir=${userDataDir}`);
    chromeOptions.addArguments(`--profile-directory=${profileName}`);
    chromeOptions.addArguments('--new-window');
    chromeOptions.addArguments('--no-first-run');
    chromeOptions.addArguments('--no-default-browser-check');

    // Stealth settings
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');

    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    // Stealth injection
    await driver.executeScript(`
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
        });
        delete window.navigator.webdriver;
    `);

    console.log('✅ Profile driver created!');
    return driver;
}

// === LOAD CONTRACTOR DATA ===
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
                    code: code,
                    fullName: fullName,
                    firstName: firstName,
                    lastName: lastName,
                    department: department || 'Security',
                    username: username,
                    password: username
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

// === MANUAL NAVIGATION HELPER ===
async function waitForManualNavigation(driver) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('👋 MANUAL NAVIGATION MODE');
    console.log('═══════════════════════════════════════');
    console.log('');
    console.log('📋 INSTRUCTIONS:');
    console.log('1. 🌐 Use the browser window to navigate manually');
    console.log('2. 🔐 Log into Rivo Safeguard if needed');
    console.log('3. 🧭 Navigate to the USER CREATION page');
    console.log('4. 📝 Get to the point where you see the user creation form');
    console.log('5. ✅ Press Enter here when you\'re ready');
    console.log('');
    console.log('🎯 TARGET: Get to the user creation form page');
    console.log('📍 URL should be something like: .../insight/...');
    console.log('');

    await new Promise(resolve => {
        rl.question('👀 Navigate manually, then press Enter when ready to continue automation: ', () => {
            rl.close();
            resolve();
        });
    });

    console.log('✅ Manual navigation complete! Starting automation...\n');
}

// === WAIT FOR SPECIFIC PAGE ===
async function waitForUserCreationPage(driver) {
    console.log('🔍 Checking if we\'re on the user creation page...');

    try {
        // Check if we're already in the frame or need to switch
        try {
            await driver.switchTo().defaultContent();
        } catch (e) {
            // Already in default content
        }

        // Look for frame first
        const frames = await driver.findElements(By.css('iframe'));
        if (frames.length > 0) {
            console.log('🖼️  Found iframe, switching to it...');
            await driver.switchTo().frame(0);
        }

        // Check for user creation form elements
        await driver.wait(until.elementLocated(By.id("Username")), 5000);
        console.log('✅ Found user creation form - ready to automate!');
        return true;

    } catch (error) {
        console.log('❌ User creation form not found');
        console.log('💡 Please navigate to the user creation page manually');
        return false;
    }
}

// === HUMAN-LIKE AUTOMATION ===
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

// === FILL USER FORM (MAIN AUTOMATION) ===
async function fillUserForm(driver, contractor, automator) {
    console.log(`\n👤 Filling form for: ${contractor.username} (${contractor.fullName})`);

    try {
        // Fill basic information
        console.log('📝 Filling basic user information...');
        await automator.safeInput(By.id("Username"), contractor.username, "Username");
        await automator.safeInput(By.name("Password"), contractor.password, "Password");
        await automator.safeInput(By.name("JobTitle"), contractor.department, "Job Title");
        await automator.safeInput(By.id("Attributes.People.Forename"), contractor.firstName, "First Name");
        await automator.safeInput(By.id("Attributes.People.Surname"), contractor.lastName, "Last Name");
        await automator.safeInput(By.id("Attributes.Users.EmployeeNumber"), contractor.code, "Employee Number");

        // Handle hierarchy configuration (simplified)
        console.log('🏢 Configuring hierarchy (attempting, may skip if not available)...');
        try {
            await automator.safeClick(By.id("Hierarchies01zzz"), "Hierarchy button");
            await automator.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Hierarchy dropdown");
            await automator.safeClick(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)"), "Hierarchy lookup");
            await automator.safeClick(By.name("addHierarchyGroupsLocations"), "Add hierarchy groups");
            await automator.safeClick(By.css("#OverlayContainer > .StandardButton"), "Overlay confirm");
        } catch (error) {
            console.log('   ⚠️  Hierarchy configuration skipped - will continue without it');
        }

        // Group management (simplified)
        console.log('👥 Setting user groups (attempting, may skip if not available)...');
        try {
            await automator.safeInput(By.name("Password"), contractor.password, "Password confirmation");
            await automator.safeClick(By.name("addHierarchyGroupsLocations"), "Add hierarchy groups 2");
            await automator.safeSelectOption(By.id("UGroupID"), "Third Party Staff", "User Group");
        } catch (error) {
            console.log('   ⚠️  Group configuration skipped - will continue without it');
        }

        // Final settings
        console.log('⚙️  Setting final options...');
        try {
            await automator.safeSelectOption(By.id("Attributes.Users.StatusOfEmployment"), "Current", "Employment Status");
            await automator.safeSelectOption(By.id("UserTypeID"), "Limited access user", "User Type");
        } catch (error) {
            console.log('   ⚠️  Some final settings skipped - will continue');
        }

        // Ask before saving
        console.log('\n🤔 READY TO SAVE USER');
        console.log('═══════════════════════════════════════');
        console.log(`📋 User Summary:`);
        console.log(`   Username: ${contractor.username}`);
        console.log(`   Name: ${contractor.fullName}`);
        console.log(`   Department: ${contractor.department}`);
        console.log(`   Code: ${contractor.code}`);
        console.log('');

        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        const shouldSave = await new Promise(resolve => {
            rl.question('❓ Save this user? (y/n): ', resolve);
        });
        rl.close();

        if (shouldSave.toLowerCase() === 'y') {
            console.log('💾 Saving user...');
            await automator.safeClick(By.name("save"), "Save button");
            await automator.humanDelay(5000, 2000);
            console.log(`✅ Successfully created user: ${contractor.username}`);
        } else {
            console.log('🚫 User creation cancelled - form filled but not saved');
        }

    } catch (error) {
        console.error(`❌ Failed to fill form: ${error.message}`);
        throw error;
    }
}

// === MAIN HYBRID AUTOMATION ===
async function runHybridAutomation(mode = 'single') {
    let driver;

    try {
        console.log('🤝 HYBRID AUTOMATION - Manual Navigation + Auto Form Filling\n');

        // Load contractor data
        const contractors = mode === 'single' ?
            [loadContractorsFromExcel('single')] :
            loadContractorsFromExcel('all');

        // Create driver
        driver = await createWorkingProfileDriver();
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        const automator = new HumanAutomator(driver);

        // Try automatic navigation first
        console.log('🌐 Attempting automatic navigation...');
        try {
            await driver.get('https://www.rivosafeguard.com/insight/');
            await driver.sleep(3000);

            // Check if we can find login elements
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Automatic navigation successful!');

            // Try to navigate to user creation automatically
            console.log('🧭 Attempting to navigate to user creation...');
            await automator.safeClick(By.css(".sch-container-left"), "Left container");
            await automator.safeClick(By.css(".sch-app-launcher-button"), "App launcher");
            await automator.safeClick(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text"), "Menu item");
            await automator.safeClick(By.css(".k-drawer-item:nth-child(6)"), "User management");
            await driver.switchTo().frame(0);

            console.log('✅ Automatic navigation to user creation successful!');

        } catch (error) {
            console.log('❌ Automatic navigation failed, switching to manual mode...');

            // Manual navigation fallback
            await waitForManualNavigation(driver);

            // Wait for user to get to the right page
            let onCorrectPage = false;
            while (!onCorrectPage) {
                onCorrectPage = await waitForUserCreationPage(driver);
                if (!onCorrectPage) {
                    console.log('\n❓ Not on user creation page yet...');
                    await waitForManualNavigation(driver);
                }
            }
        }

        // Process contractors
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            console.log(`\n📊 Progress: ${i + 1}/${contractors.length}`);

            if (i > 0) {
                console.log('🔄 Ready for next user...');
                const rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });

                await new Promise(resolve => {
                    rl.question('👀 Navigate to a fresh user creation form, then press Enter: ', () => {
                        rl.close();
                        resolve();
                    });
                });

                // Wait for correct page again
                let onCorrectPage = false;
                while (!onCorrectPage) {
                    onCorrectPage = await waitForUserCreationPage(driver);
                    if (!onCorrectPage) {
                        console.log('❓ Please navigate to user creation page...');
                        await new Promise(resolve => setTimeout(resolve, 2000));
                    }
                }
            }

            // Fill the form
            await fillUserForm(driver, contractor, automator);

            console.log(`✅ Completed: ${i + 1}/${contractors.length}`);
        }

        console.log('\n🎉 Hybrid automation completed!');

    } catch (error) {
        console.error('❌ Hybrid automation failed:', error.message);
    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 10 seconds...');
            await driver.sleep(10000);
            await driver.quit();
            console.log('✅ Browser closed');
        }
    }
}

// === COMMAND LINE INTERFACE ===
if (require.main === module) {
    const args = process.argv.slice(2);

    console.log('🤝 HYBRID AUTOMATION TOOL\n');
    console.log('📋 Manual navigation + Automated form filling\n');

    if (args.includes('--help') || args.includes('-h')) {
        console.log('📚 AVAILABLE COMMANDS:\n');
        console.log('  node hybrid-automation.js           Create single user (hybrid mode)');
        console.log('  node hybrid-automation.js --all     Create all users (hybrid mode)');
        console.log('  node hybrid-automation.js --help    Show this help\n');
        console.log('🤝 HOW IT WORKS:');
        console.log('1. Opens Chrome with your profile');
        console.log('2. You navigate manually to the user creation page');
        console.log('3. Automation fills the forms automatically');
        console.log('4. You confirm before saving each user\n');

    } else if (args.includes('--all')) {
        console.log('🚀 ALL USERS MODE (Hybrid)\n');
        runHybridAutomation('all');

    } else {
        console.log('🧪 SINGLE USER MODE (Hybrid)\n');
        console.log('💡 This is perfect for testing - you navigate manually, automation fills forms\n');
        runHybridAutomation('single');
    }
}

module.exports = {
    runHybridAutomation,
    createWorkingProfileDriver,
    fillUserForm,
    waitForManualNavigation
};