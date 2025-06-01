// COMPLETE CHROME PROFILE AUTOMATION SCRIPT
// Uses your existing Chrome profile, stealth mode, and Excel processing
// Ready to run - just save as automation.js and run: node automation.js

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs');

// === CONFIGURATION ===
const CONFIG = {
    excelFileName: 'Copy of Contractors staff list.xlsx',
    defaultTimeout: 15000,
    humanDelayBase: 1000,
    humanDelayVariation: 500
};

// === CHROME PROFILE DETECTION AND LISTING ===
function findChromeProfilePath() {
    let userDataDir;

    switch (os.platform()) {
        case 'win32': // Windows
            userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
            break;
        case 'darwin': // macOS
            userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
            break;
        case 'linux': // Linux
            userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
            break;
        default:
            throw new Error('❌ Unsupported operating system');
    }

    return { userDataDir };
}

function listChromeProfiles() {
    console.log('📋 Your Chrome Profiles:\n');

    const { userDataDir } = findChromeProfilePath();

    try {
        const localStatePath = path.join(userDataDir, 'Local State');

        if (fs.existsSync(localStatePath)) {
            const localState = JSON.parse(fs.readFileSync(localStatePath, 'utf8'));
            const profiles = localState.profile?.info_cache || {};

            console.log('🎭 Found these profiles:');
            console.log('═══════════════════════════════════════');

            Object.keys(profiles).forEach((profileKey, index) => {
                const profile = profiles[profileKey];
                const name = profile.name || 'Unnamed Profile';
                const email = profile.user_name || 'No account linked';

                console.log(`${index + 1}. Profile: "${profileKey}"`);
                console.log(`   📛 Display Name: ${name}`);
                console.log(`   📧 Account: ${email}`);
                console.log(`   🔧 Usage: --profile-directory=${profileKey}`);
                console.log('───────────────────────────────────────');
            });

            console.log('\n💡 TIP: "Default" is usually your main profile\n');
            return Object.keys(profiles);

        } else {
            console.log('📋 Common Chrome profiles (auto-detected):');
            console.log('═══════════════════════════════════');
            console.log('1. "Default" - Your main profile');
            console.log('2. "Profile 1" - Second profile');
            console.log('3. "Profile 2" - Third profile');
            console.log('4. "Profile 3" - Fourth profile\n');

            return ['Default', 'Profile 1', 'Profile 2', 'Profile 3'];
        }

    } catch (error) {
        console.log('⚠️  Could not read profile information:', error.message);
        console.log('📋 Using default profile names\n');
        return ['Default', 'Profile 1', 'Profile 2'];
    }
}

// === STEALTH CHROME DRIVER CREATION ===
async function createStealthChromeDriver(profileName = 'Default') {
    console.log(`🕵️  Creating stealth Chrome driver with profile: "${profileName}"...`);

    const chromeOptions = new chrome.Options();

    // === STEALTH SETTINGS ===
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');
    chromeOptions.addArguments('--disable-extensions');
    chromeOptions.addArguments('--disable-dev-shm-usage');
    chromeOptions.addArguments('--no-sandbox');
    chromeOptions.addArguments('--disable-web-security');
    chromeOptions.addArguments('--disable-features=VizDisplayCompositor');
    chromeOptions.addArguments('--disable-ipc-flooding-protection');

    // === USER AGENT SPOOFING ===
    chromeOptions.addArguments('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36');

    // === USE YOUR CHROME PROFILE ===
    const { userDataDir } = findChromeProfilePath();
    chromeOptions.addArguments(`--user-data-dir=${userDataDir}`);
    chromeOptions.addArguments(`--profile-directory=${profileName}`);

    // === PREFERENCES ===
    chromeOptions.setUserPreferences({
        'credentials_enable_service': false,
        'profile.password_manager_enabled': false,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_settings.popups': 0
    });

    console.log(`📁 Profile path: ${userDataDir}`);
    console.log(`📂 Profile directory: ${profileName}`);

    // === CREATE DRIVER ===
    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    // === INJECT STEALTH SCRIPTS ===
    await driver.executeScript(`
        // Remove webdriver property
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
        });
        
        // Delete automation indicators
        delete window.navigator.webdriver;
        delete window.webdriver;
        delete window.domAutomation;
        delete window.domAutomationController;
        
        // Override navigator properties
        Object.defineProperty(navigator, 'plugins', {
            get: () => [1, 2, 3, 4, 5],
        });
        
        Object.defineProperty(navigator, 'languages', {
            get: () => ['en-US', 'en'],
        });
        
        // Add chrome object if needed
        if (navigator.userAgent.includes('Chrome')) {
            window.chrome = {
                runtime: {},
                loadTimes: function() { return {}; },
                csi: function() { return {}; },
                app: {}
            };
        }
    `);

    console.log('✅ Stealth Chrome driver ready with your profile!\n');
    return driver;
}

// === EXCEL DATA PROCESSING ===
function findColumnIndex(headers, possibleNames) {
    for (let i = 0; i < headers.length; i++) {
        const header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
        if (possibleNames.some(name => header.includes(name))) {
            return i;
        }
    }
    return -1;
}

function loadContractorsFromExcel(mode = 'single') {
    console.log('📁 Loading contractor data from Excel...');

    if (!fs.existsSync(CONFIG.excelFileName)) {
        throw new Error(`❌ Excel file not found: ${CONFIG.excelFileName}`);
    }

    const workbook = XLSX.readFile(CONFIG.excelFileName);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    if (!rawData || rawData.length < 2) {
        throw new Error('❌ Excel file appears to be empty or has no data rows');
    }

    // Find column indices
    const headers = rawData[0] || [];
    const codeIndex = 0; // Assuming first column is always code
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

    const contractors = [];

    // Process data rows
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

                // If single mode, return first contractor
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

    return contractors[0]; // Default to first contractor
}

// === HUMAN-LIKE AUTOMATION UTILITIES ===
class HumanAutomator {
    constructor(driver) {
        this.driver = driver;
    }

    async humanDelay(baseMs = CONFIG.humanDelayBase, variationMs = CONFIG.humanDelayVariation) {
        const randomVariation = Math.random() * variationMs * 2 - variationMs;
        let totalDelay = Math.max(baseMs + randomVariation, 200);

        // 5% chance of longer "thinking" pause
        if (Math.random() < 0.05) {
            totalDelay += Math.random() * 2000;
            console.log(`   🤔 Taking a thinking pause... (${Math.round(totalDelay)}ms)`);
        }

        await this.driver.sleep(totalDelay);
    }

    async humanType(element, text, avgDelay = 80) {
        await element.clear();
        await this.humanDelay(300, 200);

        for (let i = 0; i < text.length; i++) {
            const char = text[i];

            // 1% chance of typing mistake
            if (Math.random() < 0.01 && i > 0) {
                const wrongChar = String.fromCharCode(65 + Math.floor(Math.random() * 26));
                await element.sendKeys(wrongChar);
                await this.driver.sleep(100 + Math.random() * 200);
                await element.sendKeys(Key.BACK_SPACE);
                await this.driver.sleep(50 + Math.random() * 100);
            }

            await element.sendKeys(char);
            await this.driver.sleep(avgDelay + (Math.random() * 40 - 20));
        }

        await this.humanDelay(200, 100);
    }

    async safeClick(selector, description, timeout = CONFIG.defaultTimeout) {
        let attempts = 3;
        while (attempts > 0) {
            try {
                console.log(`   🖱️  Clicking: ${description}`);
                const element = await this.driver.wait(until.elementLocated(selector), timeout);
                await this.driver.wait(until.elementIsEnabled(element), 5000);

                // Move mouse to element area
                const actions = this.driver.actions({ bridge: true });
                await actions.move({ origin: element, x: Math.random() * 10 - 5, y: Math.random() * 10 - 5 }).perform();
                await this.humanDelay(150, 75);

                await element.click();
                console.log(`   ✅ Successfully clicked: ${description}`);
                await this.humanDelay(300, 150);
                return element;

            } catch (error) {
                attempts--;
                console.log(`   ⚠️  Click failed for ${description}, ${attempts} attempts remaining`);
                if (attempts === 0) {
                    throw new Error(`Failed to click ${description}: ${error.message}`);
                }
                await this.humanDelay(1000, 500);
            }
        }
    }

    async safeInput(selector, text, description, timeout = CONFIG.defaultTimeout) {
        try {
            console.log(`   ✏️  Entering: ${description}`);
            const element = await this.driver.wait(until.elementLocated(selector), timeout);
            await this.driver.wait(until.elementIsEnabled(element), 5000);

            await element.click();
            await this.humanDelay(300, 150);
            await this.humanType(element, text);

            // Validate input
            const actualValue = await element.getAttribute('value');
            if (actualValue !== text) {
                console.log(`   ⚠️  Input validation warning. Expected: "${text}", Got: "${actualValue}"`);
            }

            console.log(`   ✅ Successfully entered: ${description}`);
            return element;

        } catch (error) {
            throw new Error(`Failed to enter text in ${description}: ${error.message}`);
        }
    }

    async safeSelectOption(dropdownSelector, optionText, description, timeout = CONFIG.defaultTimeout) {
        try {
            console.log(`   📋 Selecting "${optionText}" in: ${description}`);
            const dropdown = await this.driver.wait(until.elementLocated(dropdownSelector), timeout);
            await this.driver.wait(until.elementIsEnabled(dropdown), 5000);

            const options = await dropdown.findElements(By.xpath(`//option[normalize-space(text()) = '${optionText}']`));
            if (options.length === 0) {
                console.log(`   ⚠️  Option "${optionText}" not found in ${description}, continuing...`);
                return dropdown;
            }

            await options[0].click();
            console.log(`   ✅ Successfully selected: ${optionText}`);
            await this.humanDelay(500, 250);
            return dropdown;

        } catch (error) {
            console.log(`   ⚠️  Failed to select option in ${description}: ${error.message}`);
            return null;
        }
    }
}

// === MAIN USER CREATION AUTOMATION ===
async function createUser(driver, contractor, automator) {
    const startTime = Date.now();

    try {
        console.log(`\n👤 Creating user: ${contractor.username} (${contractor.fullName})`);

        // Navigate to user creation
        console.log('🧭 Navigating to user creation...');
        await automator.safeClick(By.css(".sch-container-left"), "Left container");
        await automator.safeClick(By.css(".sch-app-launcher-button"), "App launcher");
        await automator.safeClick(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text"), "Menu item");
        await automator.safeClick(By.css(".k-drawer-item:nth-child(6)"), "User management");

        // Switch to frame
        console.log('🖼️  Switching to user creation frame...');
        await driver.switchTo().frame(0);
        await automator.humanDelay(2000, 1000);

        // Fill basic information
        console.log('📝 Filling basic user information...');
        await automator.safeInput(By.id("Username"), contractor.username, "Username");
        await automator.safeInput(By.name("Password"), contractor.password, "Password");
        await automator.safeInput(By.name("JobTitle"), contractor.department, "Job Title");
        await automator.safeInput(By.id("Attributes.People.Forename"), contractor.firstName, "First Name");
        await automator.safeInput(By.id("Attributes.People.Surname"), contractor.lastName, "Last Name");
        await automator.safeInput(By.id("Attributes.Users.EmployeeNumber"), contractor.code, "Employee Number");

        // Handle hierarchy configuration
        console.log('🏢 Configuring hierarchy...');
        try {
            await automator.safeClick(By.id("Hierarchies01zzz"), "Hierarchy button");
            await automator.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Hierarchy dropdown");
            await automator.safeClick(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)"), "Hierarchy lookup");

            // Optional home location
            try {
                await automator.safeClick(By.css("#\\31 657_HomeLocationHierarchy > option:nth-child(2)"), "Home location");
            } catch (e) {
                console.log('   ⚠️  Home location not found, continuing...');
            }

            await automator.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Hierarchy dropdown 2");
            await automator.safeClick(By.css(".HierarchyNode__SelectDiv-sc-1ytna50-2"), "Hierarchy node select");
            await automator.safeClick(By.css(".DropdownDisplay__DropdownText-sc-6p7u3y-0 > span:nth-child(1)"), "Dropdown text");

            // Hierarchy tree navigation
            try {
                await automator.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Hierarchy arrow 1");
                await automator.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Hierarchy arrow 2");
                await automator.safeClick(By.css(".HierarchyNode__NodeText-sc-1ytna50-1"), "Hierarchy node text");
            } catch (e) {
                console.log('   ⚠️  Hierarchy tree navigation failed, continuing...');
            }

            await automator.safeClick(By.name("addHierarchyGroupsLocations"), "Add hierarchy groups");
            await automator.safeClick(By.css("#OverlayContainer > .StandardButton"), "Overlay confirm");

        } catch (error) {
            console.log('   ⚠️  Some hierarchy operations failed, continuing...');
        }

        // Group management
        console.log('👥 Managing groups and permissions...');
        try {
            // Group tab
            try {
                await automator.safeClick(By.id("GroupTabDiv_23890"), "Group tab");
            } catch (e) {
                console.log('   ⚠️  Group tab not found, continuing...');
            }

            // Password confirmation
            await automator.safeInput(By.name("Password"), contractor.password, "Password confirmation");

            await automator.safeClick(By.name("addHierarchyGroupsLocations"), "Add hierarchy groups 2");
            await automator.safeClick(By.css("#ManageHierarchyUGroupLocationsForm > .GroupTab:nth-child(1)"), "Group tab form");

            // Select user group
            await automator.safeSelectOption(By.id("UGroupID"), "Third Party Staff", "User Group");

            // Additional hierarchy selections
            try {
                await automator.safeClick(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1"), "Final hierarchy dropdown");
                await automator.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Final hierarchy arrow 1");
                await automator.safeClick(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6"), "Final hierarchy arrow 2");
                await automator.safeClick(By.css(".HierarchyNode__NodeText-sc-1ytna50-1"), "Final hierarchy node");
                await automator.safeClick(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)"), "Final hierarchy lookup");
            } catch (e) {
                console.log('   ⚠️  Additional hierarchy selections failed, continuing...');
            }

        } catch (error) {
            console.log('   ⚠️  Some group operations failed, continuing...');
        }

        // Final settings
        console.log('⚙️  Configuring final settings...');
        await automator.safeSelectOption(By.id("Attributes.Users.StatusOfEmployment"), "Current", "Employment Status");
        await automator.safeSelectOption(By.id("UserTypeID"), "Limited access user", "User Type");

        // Save user
        console.log('💾 Saving user...');
        await automator.safeClick(By.name("save"), "Save button");
        await automator.humanDelay(5000, 2000);

        const endTime = Date.now();
        const totalTime = (endTime - startTime) / 1000;

        console.log(`   ✅ Successfully created user: ${contractor.username}`);
        console.log(`   ⏱️  Total time: ${totalTime.toFixed(1)} seconds`);

        // Ensure minimum time for human-like behavior
        if (totalTime < 35) {
            const additionalDelay = (40 - totalTime) * 1000;
            console.log(`   ⏳ Adding ${(additionalDelay / 1000).toFixed(1)}s delay for human-like timing...`);
            await automator.humanDelay(additionalDelay, 1000);
        }

    } catch (error) {
        console.error(`   ❌ Failed to create user ${contractor.username}: ${error.message}`);
        throw error;
    }
}

// === MAIN AUTOMATION FUNCTIONS ===
async function runSingleUserTest(profileName = 'Default') {
    let driver;

    try {
        console.log('🧪 SINGLE USER TEST - Starting...\n');

        // Load single contractor
        const contractor = loadContractorsFromExcel('single');

        // Create stealth driver
        driver = await createStealthChromeDriver(profileName);
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        const automator = new HumanAutomator(driver);

        // Navigate and check login
        console.log('🌐 Navigating to Rivo Safeguard...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await automator.humanDelay(3000, 1500);

        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Successfully using existing login!\n');
        } catch (error) {
            throw new Error('❌ Not logged in. Please log in to Chrome first.');
        }

        // Create user
        await createUser(driver, contractor, automator);

        console.log('\n🎉 Single user test completed successfully!');

    } catch (error) {
        console.error(`\n❌ Single user test failed: ${error.message}`);
        throw error;
    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 10 seconds to review...');
            await driver.sleep(10000);
            await driver.quit();
            console.log('✅ Browser closed');
        }
    }
}

async function runAllUsersAutomation(profileName = 'Default') {
    let driver;

    try {
        console.log('🚀 ALL USERS AUTOMATION - Starting...\n');

        // Load all contractors
        const contractors = loadContractorsFromExcel('all');
        console.log(`📊 Found ${contractors.length} Security contractors to create`);
        console.log(`⏱️  Estimated time: ${(contractors.length * 45 / 60).toFixed(1)} minutes\n`);

        // Show first few contractors
        console.log('👥 First 5 contractors to be created:');
        contractors.slice(0, 5).forEach((contractor, index) => {
            console.log(`   ${index + 1}. ${contractor.fullName} -> ${contractor.username}`);
        });
        console.log('');

        // Create stealth driver
        driver = await createStealthChromeDriver(profileName);
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        const automator = new HumanAutomator(driver);

        // Navigate and check login
        console.log('🌐 Navigating to Rivo Safeguard...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await automator.humanDelay(3000, 1500);

        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Successfully using existing login!\n');
        } catch (error) {
            throw new Error('❌ Not logged in. Please log in to Chrome first.');
        }

        // Create all users
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            try {
                console.log(`\n📊 Progress: ${i + 1}/${contractors.length}`);

                if (i > 0) {
                    // Navigate back to user creation for subsequent users
                    await driver.switchTo().defaultContent();
                    await automator.humanDelay(1000, 500);
                    await automator.safeClick(By.css(".k-drawer-item:nth-child(6)"), "User management");
                    await driver.switchTo().frame(0);
                    await automator.humanDelay(2000, 1000);
                }

                await createUser(driver, contractor, automator);

                console.log(`✅ Completed: ${i + 1}/${contractors.length} users created`);

                // Delay between users
                if (i < contractors.length - 1) {
                    console.log('   🔄 Preparing for next user...');
                    await automator.humanDelay(3000, 2000);
                }

            } catch (error) {
                console.error(`❌ Failed to create user ${i + 1}: ${contractor.username}`);

                // Ask user if they want to continue
                const readline = require('readline');
                const rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });

                const answer = await new Promise(resolve => {
                    rl.question('❓ Do you want to continue with the next user? (y/n): ', resolve);
                });
                rl.close();

                if (answer.toLowerCase() !== 'y') {
                    console.log('🛑 Stopping automation as requested.');
                    break;
                }
            }
        }

        console.log('\n🎉 All users automation completed!');

    } catch (error) {
        console.error(`\n❌ All users automation failed: ${error.message}`);
        throw error;
    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 10 seconds to review...');
            await driver.sleep(10000);
            await driver.quit();
            console.log('✅ Browser closed');
        }
    }
}

// === INTERACTIVE PROFILE SELECTOR ===
async function selectProfileInteractively() {
    const readline = require('readline');
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('🎮 Interactive Profile Selector\n');

    const availableProfiles = listChromeProfiles();

    console.log('❓ Which profile would you like to use?');
    availableProfiles.forEach((profile, index) => {
        console.log(`   ${index + 1}. ${profile}`);
    });
    console.log(`   ${availableProfiles.length + 1}. Enter custom profile name`);

    const choice = await new Promise(resolve => {
        rl.question('\nEnter your choice (number): ', resolve);
    });

    let selectedProfile;
    const choiceNum = parseInt(choice);

    if (choiceNum >= 1 && choiceNum <= availableProfiles.length) {
        selectedProfile = availableProfiles[choiceNum - 1];
    } else if (choiceNum === availableProfiles.length + 1) {
        selectedProfile = await new Promise(resolve => {
            rl.question('Enter custom profile directory name: ', resolve);
        });
    } else {
        console.log('❌ Invalid choice. Using Default profile.');
        selectedProfile = 'Default';
    }

    rl.close();

    console.log(`✅ Selected profile: "${selectedProfile}"\n`);
    return selectedProfile;
}

// === MAIN ENTRY POINT ===
async function main() {
    const args = process.argv.slice(2);

    console.log('🎯 CHROME PROFILE AUTOMATION TOOL\n');
    console.log('🔑 Uses your existing Chrome profile with all logins\n');

    try {
        if (args.includes('--list-profiles') || args.includes('-l')) {
            // List profiles only
            console.log('📋 PROFILE LISTING MODE\n');
            listChromeProfiles();

        } else if (args.includes('--help') || args.includes('-h')) {
            // Show help
            console.log('📚 AVAILABLE COMMANDS:\n');
            console.log('  node automation.js                    Run single user test (recommended first)');
            console.log('  node automation.js --all              Create ALL users from Excel');
            console.log('  node automation.js --list-profiles    List all Chrome profiles');
            console.log('  node automation.js --select-profile   Interactive profile selection');
            console.log('  node automation.js --help             Show this help\n');
            console.log('📋 REQUIREMENTS:');
            console.log('  ✅ Chrome browser installed');
            console.log('  ✅ Excel file: "Copy of Contractors staff list.xlsx" in same folder');
            console.log('  ✅ Already logged into https://www.rivosafeguard.com/insight/');
            console.log('  ✅ ALL Chrome windows closed before running\n');

        } else if (args.includes('--all')) {
            // Run all users automation
            console.log('🚀 ALL USERS MODE\n');

            let profileName = 'Default';
            if (args.includes('--select-profile')) {
                profileName = await selectProfileInteractively();
            } else {
                console.log('🔍 Available Chrome profiles:');
                listChromeProfiles();
                console.log(`🎯 Using profile: "${profileName}"\n`);
            }

            console.log('⚠️  IMPORTANT: Make sure ALL Chrome windows are closed!\n');
            await new Promise(resolve => setTimeout(resolve, 3000));

            await runAllUsersAutomation(profileName);

        } else if (args.includes('--select-profile')) {
            // Interactive profile selection for single user
            console.log('🎮 INTERACTIVE PROFILE SELECTION\n');
            const profileName = await selectProfileInteractively();

            console.log('⚠️  IMPORTANT: Make sure ALL Chrome windows are closed!\n');
            await new Promise(resolve => setTimeout(resolve, 3000));

            await runSingleUserTest(profileName);

        } else {
            // Default: Single user test
            console.log('🧪 SINGLE USER TEST MODE (Recommended for first run)\n');
            console.log('🔍 Available Chrome profiles:');
            listChromeProfiles();
            console.log('🎯 Using "Default" profile (your main profile)\n');

            console.log('⚠️  BEFORE RUNNING:');
            console.log('1. ✅ Log into https://www.rivosafeguard.com/insight/ in Chrome');
            console.log('2. ✅ Close ALL Chrome browser windows');
            console.log('3. ✅ Make sure Excel file is in this folder');
            console.log('4. ✅ Press Enter to start...\n');

            // Wait for user confirmation
            process.stdin.setRawMode(true);
            process.stdin.resume();
            process.stdin.once('data', () => {
                process.stdin.setRawMode(false);
                runSingleUserTest('Default');
            });
        }

    } catch (error) {
        console.error('💥 Application failed:', error.message);

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

        process.exit(1);
    }
}

// === RUN THE APPLICATION ===
if (require.main === module) {
    main();
}

// === EXPORTS ===
module.exports = {
    createStealthChromeDriver,
    listChromeProfiles,
    loadContractorsFromExcel,
    runSingleUserTest,
    runAllUsersAutomation,
    HumanAutomator
};