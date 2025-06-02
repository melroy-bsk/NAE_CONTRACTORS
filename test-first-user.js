// CHROME EXISTING PROFILE AUTOMATION
// Multiple methods to use your existing Chrome profile

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const XLSX = require('xlsx');
const fs = require('fs');
const readline = require('readline');
const { exec, spawn } = require('child_process');

// === CONFIGURATION ===
const CONFIG = {
    excelFileName: 'Copy of Contractors staff list.xlsx',
    defaultTimeout: 30000,
    humanDelayBase: 1500,
    humanDelayVariation: 1000,
    debugPort: 9222,
    method: 'auto' // 'auto', 'manual', or 'extension'
};

// === CHECK CHROME DEBUGGING PORT ===
async function checkChromeDebugging(port = CONFIG.debugPort) {
    return new Promise((resolve) => {
        const command = os.platform() === 'win32'
            ? `netstat -an | findstr :${port}`
            : `lsof -i :${port}`;

        exec(command, (error, stdout) => {
            if (stdout && stdout.includes(port.toString())) {
                console.log(`✅ Port ${port} is active`);

                // Try to verify it's Chrome
                exec(`curl -s http://localhost:${port}/json/version`, (error2, stdout2) => {
                    if (!error2 && stdout2) {
                        try {
                            const info = JSON.parse(stdout2);
                            console.log(`✅ Chrome debugging confirmed: ${info.Browser}`);
                            resolve(true);
                        } catch (e) {
                            console.log(`⚠️  Port ${port} is active but not Chrome debugging`);
                            resolve(false);
                        }
                    } else {
                        resolve(false);
                    }
                });
            } else {
                resolve(false);
            }
        });
    });
}

// === KILL CHROME PROCESSES ===
async function killChrome() {
    console.log('🔄 Closing Chrome processes...');

    return new Promise((resolve) => {
        const commands = {
            win32: 'taskkill /F /IM chrome.exe /T 2>nul',
            darwin: 'pkill -f "Google Chrome"',
            linux: 'pkill chrome'
        };

        const command = commands[os.platform()] || commands.linux;

        exec(command, () => {
            setTimeout(resolve, 2000);
        });
    });
}

// === GET CHROME PATHS ===
function getChromePaths() {
    const platform = os.platform();
    let chromePath, userDataDir;

    if (platform === 'darwin') {
        chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
    } else if (platform === 'win32') {
        const possiblePaths = [
            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe',
            path.join(os.homedir(), 'AppData\\Local\\Google\\Chrome\\Application\\chrome.exe')
        ];

        chromePath = possiblePaths.find(p => fs.existsSync(p)) || 'chrome.exe';
        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else {
        const possiblePaths = [
            '/usr/bin/google-chrome',
            '/usr/bin/google-chrome-stable',
            '/usr/bin/chromium-browser',
            '/usr/bin/chromium'
        ];

        chromePath = possiblePaths.find(p => fs.existsSync(p)) || 'google-chrome';
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    return { chromePath, userDataDir };
}

// === METHOD 1: AUTO START CHROME ===
async function autoStartChrome() {
    console.log('🚀 Auto-starting Chrome with debugging...\n');

    const { chromePath, userDataDir } = getChromePaths();

    if (!fs.existsSync(chromePath)) {
        throw new Error(`Chrome not found at: ${chromePath}`);
    }

    // Kill existing Chrome
    await killChrome();

    // Start Chrome with debugging
    const args = [
        `--remote-debugging-port=${CONFIG.debugPort}`,
        '--no-first-run',
        '--no-default-browser-check',
        `--user-data-dir=${userDataDir}`
    ];

    console.log(`📍 Chrome path: ${chromePath}`);
    console.log(`📁 Profile path: ${userDataDir}`);
    console.log(`🔧 Debug port: ${CONFIG.debugPort}\n`);

    return new Promise((resolve, reject) => {
        const chromeProcess = spawn(chromePath, args, {
            detached: true,
            stdio: 'ignore'
        });

        chromeProcess.unref();

        chromeProcess.on('error', (err) => {
            console.error('❌ Failed to start Chrome:', err.message);
            reject(err);
        });

        // Wait for Chrome to start
        let attempts = 0;
        const checkInterval = setInterval(async () => {
            attempts++;

            const isReady = await checkChromeDebugging();
            if (isReady) {
                clearInterval(checkInterval);
                console.log('✅ Chrome started successfully!\n');
                resolve();
            } else if (attempts > 10) {
                clearInterval(checkInterval);
                reject(new Error('Chrome failed to start with debugging'));
            }
        }, 1000);
    });
}

// === METHOD 2: MANUAL INSTRUCTIONS ===
async function showManualInstructions() {
    console.log('\n' + '═'.repeat(80));
    console.log('📋 MANUAL CHROME SETUP');
    console.log('═'.repeat(80));

    const { chromePath, userDataDir } = getChromePaths();
    const platform = os.platform();

    console.log('\n1️⃣  First, close ALL Chrome windows\n');

    console.log('2️⃣  Then run this command:\n');

    if (platform === 'win32') {
        console.log('   Option A (if Chrome is in PATH):');
        console.log('   chrome --remote-debugging-port=9222\n');

        console.log('   Option B (full path):');
        console.log(`   "${chromePath}" --remote-debugging-port=9222\n`);

        console.log('   Option C (with specific profile):');
        console.log(`   "${chromePath}" --remote-debugging-port=9222 --user-data-dir="${userDataDir}"\n`);
    } else if (platform === 'darwin') {
        console.log('   In Terminal:');
        console.log(`   "${chromePath}" --remote-debugging-port=9222\n`);
    } else {
        console.log('   In Terminal:');
        console.log(`   ${chromePath} --remote-debugging-port=9222\n`);
    }

    console.log('3️⃣  Chrome will open with YOUR profile and saved logins\n');
    console.log('═'.repeat(80));

    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    await new Promise(resolve => {
        rl.question('\n✅ Press Enter after starting Chrome: ', () => {
            rl.close();
            resolve();
        });
    });
}

// === METHOD 3: CHROME EXTENSION APPROACH ===
async function useChromeExtension() {
    console.log('🧩 Using Chrome Extension method...\n');

    const { userDataDir } = getChromePaths();

    const chromeOptions = new chrome.Options();
    chromeOptions.addArguments(`--user-data-dir=${userDataDir}`);
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');

    // This will use your existing profile
    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    return driver;
}

// === CONNECT TO CHROME ===
async function connectToChrome(method = CONFIG.method) {
    console.log('🔌 Setting up Chrome connection...\n');

    try {
        if (method === 'extension') {
            // Try extension method
            return await useChromeExtension();
        }

        // Check if Chrome is already running with debugging
        let isDebugging = await checkChromeDebugging();

        if (!isDebugging) {
            if (method === 'auto') {
                // Try auto-start
                try {
                    await autoStartChrome();
                    isDebugging = true;
                } catch (err) {
                    console.log('⚠️  Auto-start failed, switching to manual mode');
                    console.log(`   Error: ${err.message}\n`);
                    method = 'manual';
                }
            }

            if (method === 'manual' && !isDebugging) {
                await showManualInstructions();

                // Verify Chrome started
                isDebugging = await checkChromeDebugging();
                if (!isDebugging) {
                    throw new Error('Chrome debugging port still not found');
                }
            }
        }

        // Connect to Chrome
        const chromeOptions = new chrome.Options();
        chromeOptions.debuggerAddress(`localhost:${CONFIG.debugPort}`);

        console.log('🔗 Connecting to Chrome...');

        const driver = await new Builder()
            .forBrowser('chrome')
            .setChromeOptions(chromeOptions)
            .build();

        // Test connection
        try {
            const url = await driver.getCurrentUrl();
            console.log('✅ Connected successfully!');
            console.log(`📍 Current page: ${url}\n`);
        } catch (e) {
            console.log('✅ Connected to Chrome!\n');
        }

        return driver;

    } catch (error) {
        console.error('❌ Connection failed:', error.message);

        console.log('\n💡 Troubleshooting:');
        console.log('1. Make sure ALL Chrome windows are closed');
        console.log('2. Try running with --manual flag');
        console.log('3. Or try --extension flag for direct profile access\n');

        throw error;
    }
}

// === NATURAL AUTOMATOR CLASS ===
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

        await element.sendKeys(Key.chord(Key.CONTROL, "a"));
        await this.randomDelay(100, 300);
        await element.sendKeys(Key.DELETE);
        await this.randomDelay(200, 400);

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
            console.log(`   ⚠️  Failed: ${error.message}`);
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
            console.log(`   ⚠️  Failed: ${error.message}`);
            throw error;
        }
    }
}

// === NAVIGATION HELPER ===
async function waitForNavigation(message) {
    const rl = readline.createInterface({
        input: process.stdin,
        output: process.stdout
    });

    console.log('\n' + '═'.repeat(60));
    console.log(`📋 ${message}`);
    console.log('═'.repeat(60));

    await new Promise(resolve => {
        rl.question('\n✅ Press Enter when ready: ', () => {
            rl.close();
            resolve();
        });
    });
}

// === DETECT FORM ===
async function detectUserForm(driver) {
    try {
        await driver.switchTo().defaultContent();

        try {
            await driver.findElement(By.id("Username"));
            return true;
        } catch (e) { }

        const frames = await driver.findElements(By.css('iframe'));
        for (let i = 0; i < frames.length; i++) {
            try {
                await driver.switchTo().defaultContent();
                await driver.switchTo().frame(i);
                await driver.findElement(By.id("Username"));
                return true;
            } catch (e) { }
        }

        await driver.switchTo().defaultContent();
        return false;
    } catch (error) {
        return false;
    }
}

// === LOAD CONTRACTORS ===
function loadContractorsFromExcel(mode = 'single') {
    console.log('📁 Loading contractors...\n');

    if (!fs.existsSync(CONFIG.excelFileName)) {
        throw new Error(`Excel file not found: ${CONFIG.excelFileName}`);
    }

    const workbook = XLSX.readFile(CONFIG.excelFileName);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];
    const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    const headers = rawData[0] || [];
    const contractors = [];

    for (let i = 1; i < rawData.length; i++) {
        const row = rawData[i];
        if (row && row[0] && row[1] && row[2] && row[3]) {
            const code = row[0].toString().trim();
            const firstName = row[1].toString().trim();
            const lastName = row[2].toString().trim();
            const department = row[3] ? row[3].toString().trim() : '';

            if (department.toLowerCase() === 'security') {
                const fullName = `${firstName} ${lastName}`;
                const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                contractors.push({
                    code, fullName, firstName, lastName,
                    department: 'Security',
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
async function runAutomation(options = {}) {
    let driver;

    try {
        console.log('🌟 CHROME PROFILE AUTOMATION\n');

        const mode = options.mode || 'single';
        const method = options.method || CONFIG.method;

        // Load contractors
        const contractors = mode === 'single' ?
            [loadContractorsFromExcel('single')] :
            loadContractorsFromExcel('all');

        // Connect to Chrome
        console.log(`📋 Connection method: ${method}\n`);
        driver = await connectToChrome(method);
        const automator = new NaturalHumanAutomator(driver);

        // Initial navigation
        console.log('📍 Navigate to Rivo Safeguard:');
        console.log('   1. Go to https://www.rivosafeguard.com/insight/');
        console.log('   2. Log in if needed (Google SSO should work)');
        console.log('   3. Navigate to User Management → Create User\n');

        await waitForNavigation("Navigate to user creation form");

        // Process contractors
        for (let i = 0; i < contractors.length; i++) {
            const contractor = contractors[i];

            console.log(`\n${'═'.repeat(50)}`);
            console.log(`📊 Contractor ${i + 1} of ${contractors.length}`);
            console.log('═'.repeat(50));

            // Detect form
            const formFound = await detectUserForm(driver);
            if (!formFound) {
                console.log('⚠️  Form not found');
                await waitForNavigation("Navigate to user creation form");
            }

            // Fill form
            console.log(`\n👤 Creating: ${contractor.fullName}`);

            try {
                await automator.waitAndType(By.id("Username"), contractor.username, "Username");
                await automator.waitAndType(By.name("Password"), contractor.password, "Password");
                await automator.waitAndType(By.name("JobTitle"), contractor.department, "Job Title");
                await automator.waitAndType(By.id("Attributes.People.Forename"), contractor.firstName, "First Name");
                await automator.waitAndType(By.id("Attributes.People.Surname"), contractor.lastName, "Last Name");
                await automator.waitAndType(By.id("Attributes.Users.EmployeeNumber"), contractor.code, "Employee Number");

                const rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });

                const save = await new Promise(resolve => {
                    rl.question('\n💾 Save this user? (y/n): ', (answer) => {
                        rl.close();
                        resolve(answer.toLowerCase() === 'y');
                    });
                });

                if (save) {
                    await automator.waitAndClick(By.name("save"), "Save button");
                    console.log('✅ User created');
                }

                if (i < contractors.length - 1) {
                    await waitForNavigation("Navigate to create another user");
                }

            } catch (err) {
                console.error('❌ Error:', err.message);
            }
        }

        console.log('\n🎉 COMPLETED!\n');

    } catch (error) {
        console.error('❌ Fatal error:', error.message);
    }
}

// === CLI ===
if (require.main === module) {
    const args = process.argv.slice(2);

    const options = {
        mode: args.includes('--all') ? 'all' : 'single',
        method: args.includes('--manual') ? 'manual' :
            args.includes('--extension') ? 'extension' : 'auto'
    };

    if (args.includes('--help')) {
        console.log('📚 USAGE:\n');
        console.log('  node script.js              Auto mode, single user');
        console.log('  node script.js --all        Auto mode, all users');
        console.log('  node script.js --manual     Manual Chrome start');
        console.log('  node script.js --extension  Direct profile access\n');

        console.log('🔧 METHODS:');
        console.log('  auto       - Tries to start Chrome automatically');
        console.log('  manual     - Shows instructions to start Chrome');
        console.log('  extension  - Uses profile directly (may have limits)\n');
    } else {
        runAutomation(options);
    }
}

module.exports = { runAutomation };