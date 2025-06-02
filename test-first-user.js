// PROFILE FIX DIAGNOSTIC - Step by Step Solution
// This will identify and fix the Chrome profile issue

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const path = require('path');
const os = require('os');
const fs = require('fs');
const { exec } = require('child_process');

// === STEP 1: CHECK IF CHROME IS RUNNING ===
function checkChromeProcesses() {
    console.log('🔍 STEP 1: Checking if Chrome is running...\n');

    return new Promise((resolve) => {
        let command;

        switch (os.platform()) {
            case 'win32':
                command = 'tasklist /FI "IMAGENAME eq chrome.exe" /FO CSV';
                break;
            case 'darwin':
                command = 'ps aux | grep -i "Google Chrome" | grep -v grep';
                break;
            case 'linux':
                command = 'ps aux | grep -i chrome | grep -v grep';
                break;
            default:
                console.log('❓ Cannot check Chrome processes on this OS');
                resolve(false);
                return;
        }

        exec(command, (error, stdout, stderr) => {
            const isRunning = stdout.includes('chrome') || stdout.includes('Chrome');

            if (isRunning) {
                console.log('❌ Chrome is STILL RUNNING!');
                console.log('🔧 YOU MUST close ALL Chrome windows first\n');

                if (os.platform() === 'win32') {
                    console.log('💡 Windows: Check Task Manager and end all chrome.exe processes');
                    console.log('   Or run: taskkill /F /IM chrome.exe');
                } else if (os.platform() === 'darwin') {
                    console.log('💡 Mac: Command+Option+Esc → Force quit Google Chrome');
                    console.log('   Or run: killall "Google Chrome"');
                } else {
                    console.log('💡 Linux: Run: pkill chrome');
                }

                console.log('\n🛑 STOP: Close Chrome completely and run this script again!\n');
            } else {
                console.log('✅ Chrome is not running - Good!\n');
            }

            resolve(isRunning);
        });
    });
}

// === STEP 2: FIND ACTUAL CHROME PROFILE PATH ===
function findAndVerifyProfilePath() {
    console.log('🔍 STEP 2: Finding and verifying Chrome profile path...\n');

    // Try multiple possible paths
    const possiblePaths = [];

    if (os.platform() === 'win32') {
        possiblePaths.push(
            path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data'),
            path.join(os.homedir(), 'AppData', 'Roaming', 'Google', 'Chrome', 'User Data'),
            path.join('C:', 'Users', os.userInfo().username, 'AppData', 'Local', 'Google', 'Chrome', 'User Data')
        );
    } else if (os.platform() === 'darwin') {
        possiblePaths.push(
            path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome')
        );
    } else {
        possiblePaths.push(
            path.join(os.homedir(), '.config', 'google-chrome'),
            path.join(os.homedir(), '.config', 'chromium')
        );
    }

    console.log('📁 Checking possible Chrome paths:');
    console.log('═══════════════════════════════════════');

    let validPath = null;

    for (const testPath of possiblePaths) {
        const exists = fs.existsSync(testPath);
        console.log(`${exists ? '✅' : '❌'} ${testPath}`);

        if (exists && !validPath) {
            // Verify it has profiles
            try {
                const contents = fs.readdirSync(testPath);
                const hasProfiles = contents.some(item =>
                    item === 'Default' || item.startsWith('Profile')
                );

                if (hasProfiles) {
                    validPath = testPath;
                    console.log(`   📂 Contains profiles: ${contents.filter(item =>
                        item === 'Default' || item.startsWith('Profile')
                    ).join(', ')}`);
                }
            } catch (e) {
                console.log(`   ❌ Cannot read directory: ${e.message}`);
            }
        }
    }

    if (validPath) {
        console.log(`\n🎯 FOUND valid Chrome profile path: ${validPath}\n`);
        return validPath;
    } else {
        console.log('\n❌ NO valid Chrome profile path found!');
        console.log('🔧 SOLUTIONS:');
        console.log('1. Install Google Chrome');
        console.log('2. Open Chrome at least once to create profiles');
        console.log('3. Check if Chrome is installed in a custom location\n');
        return null;
    }
}

// === STEP 3: VERIFY SPECIFIC PROFILE ===
function verifySpecificProfile(userDataDir, profileName) {
    console.log(`🔍 STEP 3: Verifying profile "${profileName}"...\n`);

    const profilePath = path.join(userDataDir, profileName);

    console.log(`📂 Checking: ${profilePath}`);

    if (!fs.existsSync(profilePath)) {
        console.log('❌ Profile directory does NOT exist!');

        // List what profiles actually exist
        try {
            const actualProfiles = fs.readdirSync(userDataDir)
                .filter(item => {
                    const itemPath = path.join(userDataDir, item);
                    return fs.statSync(itemPath).isDirectory() &&
                        (item === 'Default' || item.startsWith('Profile'));
                });

            console.log('📋 Available profiles:');
            actualProfiles.forEach(profile => console.log(`   - ${profile}`));

            return { exists: false, actualProfiles };
        } catch (e) {
            console.log('❌ Cannot read profile directory');
            return { exists: false, actualProfiles: [] };
        }
    }

    // Check if profile has preferences
    const prefsPath = path.join(profilePath, 'Preferences');
    const hasPrefs = fs.existsSync(prefsPath);

    console.log(`✅ Profile directory exists: ${profilePath}`);
    console.log(`${hasPrefs ? '✅' : '❌'} Has Preferences file: ${hasPrefs}`);

    if (hasPrefs) {
        try {
            const prefs = JSON.parse(fs.readFileSync(prefsPath, 'utf8'));
            const profileInfo = prefs.profile || {};
            const accountInfo = prefs.account_info || [];

            console.log('📋 Profile details:');
            console.log(`   Name: ${profileInfo.name || 'Not set'}`);
            console.log(`   Accounts: ${accountInfo.length > 0 ?
                accountInfo.map(acc => acc.email || acc.gaia_id).join(', ') : 'None'}`);
        } catch (e) {
            console.log('⚠️  Could not read profile preferences');
        }
    }

    console.log('');
    return { exists: true, hasPrefs, profilePath };
}

// === STEP 4: TEST CHROME WITH MANUAL COMMAND ===
function generateManualCommand(userDataDir, profileName) {
    console.log('🔍 STEP 4: Manual Chrome command test...\n');

    let chromePath;
    let command;

    if (os.platform() === 'win32') {
        chromePath = '"C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"';
        command = `${chromePath} --user-data-dir="${userDataDir}" --profile-directory="${profileName}" --new-window`;
    } else if (os.platform() === 'darwin') {
        chromePath = '"/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"';
        command = `${chromePath} --user-data-dir="${userDataDir}" --profile-directory="${profileName}" --new-window`;
    } else {
        chromePath = 'google-chrome';
        command = `${chromePath} --user-data-dir="${userDataDir}" --profile-directory="${profileName}" --new-window`;
    }

    console.log('🧪 MANUAL TEST: Run this command in your terminal/command prompt:');
    console.log('═══════════════════════════════════════════════════════════════');
    console.log(command);
    console.log('═══════════════════════════════════════════════════════════════');
    console.log('');
    console.log('📋 This should:');
    console.log('1. Open Chrome with the correct profile');
    console.log('2. Show your login status');
    console.log('3. Verify the profile is working');
    console.log('');
    console.log('❓ If this command works, the automation should work too');
    console.log('❓ If this command fails, we need to fix the profile first\n');

    return command;
}

// === STEP 5: TEST SELENIUM WITH CORRECTED SETTINGS ===
async function testSeleniumWithProfile(userDataDir, profileName) {
    console.log('🔍 STEP 5: Testing Selenium with corrected profile settings...\n');

    let driver;

    try {
        const chromeOptions = new chrome.Options();

        // === CRITICAL PROFILE SETTINGS ===
        // Try different formats to see which works
        console.log('🧪 Testing different argument formats...');

        // Format 1: With quotes (Windows-friendly)
        chromeOptions.addArguments(`--user-data-dir="${userDataDir}"`);
        chromeOptions.addArguments(`--profile-directory=${profileName}`);

        // Additional flags
        chromeOptions.addArguments('--no-first-run');
        chromeOptions.addArguments('--no-default-browser-check');
        chromeOptions.addArguments('--disable-default-apps');
        chromeOptions.addArguments('--disable-infobars');
        chromeOptions.addArguments('--disable-extensions');

        // Minimal stealth (for testing)
        chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
        chromeOptions.excludeSwitches(['enable-automation']);

        console.log('📋 Chrome arguments:');
        console.log(`   --user-data-dir="${userDataDir}"`);
        console.log(`   --profile-directory=${profileName}`);
        console.log('');

        console.log('🚀 Creating Selenium driver...');
        driver = await new Builder()
            .forBrowser('chrome')
            .setChromeOptions(chromeOptions)
            .build();

        console.log('✅ Driver created successfully!');

        // Test navigation
        console.log('🌐 Testing navigation...');
        await driver.get('chrome://version/');
        await driver.sleep(2000);

        // Check profile info
        const pageText = await driver.executeScript('return document.documentElement.innerText;');

        if (pageText.includes(profileName)) {
            console.log(`✅ SUCCESS! Profile "${profileName}" is being used by Selenium!`);
        } else {
            console.log(`❌ FAILED! Profile "${profileName}" is NOT being used`);
            console.log('📋 Page shows:');
            console.log(pageText.substring(0, 500) + '...');
        }

        // Test Rivo Safeguard
        console.log('\n🌐 Testing Rivo Safeguard login...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await driver.sleep(5000);

        // Check for login elements
        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ SUCCESS! Found login elements - Profile has active session!');
            console.log('🎉 PROFILE IS WORKING WITH SELENIUM!');
            return true;
        } catch (error) {
            console.log('❌ Login elements not found - Profile not logged in or wrong profile');
            console.log('💡 Solution: Log into Rivo Safeguard with this profile first');
            return false;
        }

    } catch (error) {
        console.log(`❌ Selenium test failed: ${error.message}`);

        if (error.message.includes('user data directory is already in use')) {
            console.log('🔧 Solution: Chrome is still running - close it completely');
        } else if (error.message.includes('cannot find Chrome binary')) {
            console.log('🔧 Solution: Chrome is not installed or not in PATH');
        } else {
            console.log('🔧 Solution: Check the manual command first');
        }

        return false;

    } finally {
        if (driver) {
            console.log('\n⏰ Keeping browser open for 10 seconds to verify...');
            await driver.sleep(10000);
            await driver.quit();
            console.log('✅ Test browser closed');
        }
    }
}

function getWorkingProfileSettings() {
    let chromePath = null;
    let userDataDir = null;

    if (os.platform() === 'darwin') {
        // macOS - We know this works for you
        chromePath = '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome';
        userDataDir = path.join(os.homedir(), 'Library', 'Application Support', 'Google', 'Chrome');
    } else if (os.platform() === 'win32') {
        // Windows - Standard paths
        const possibleChromePaths = [
            'C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe',
            'C:\\Program Files (x86)\\Google\\Chrome\\Application\\chrome.exe'
        ];

        // Find which Chrome path exists
        for (const testPath of possibleChromePaths) {
            if (fs.existsSync(testPath)) {
                chromePath = testPath;
                break;
            }
        }

        userDataDir = path.join(os.homedir(), 'AppData', 'Local', 'Google', 'Chrome', 'User Data');
    } else {
        // Linux
        chromePath = 'google-chrome'; // Usually in PATH
        userDataDir = path.join(os.homedir(), '.config', 'google-chrome');
    }

    console.log(`🔍 Platform: ${os.platform()}`);
    console.log(`📁 Chrome: ${chromePath}`);
    console.log(`📂 User Data: ${userDataDir}`);

    // Verify paths exist
    if (chromePath && chromePath !== 'google-chrome' && !fs.existsSync(chromePath)) {
        console.warn(`⚠️  Chrome not found at: ${chromePath}`);
    }

    if (!fs.existsSync(userDataDir)) {
        console.warn(`⚠️  User data not found at: ${userDataDir}`);
    }

    return { chromePath, userDataDir };
}

async function createWorkingProfileDriver(profileName = 'Profile 1') {
    console.log(`🔧 Creating working Chrome driver with Profile: "${profileName}"...`);

    const { chromePath, userDataDir } = getWorkingProfileSettings();
    const chromeOptions = new chrome.Options();

    // === SET CHROME BINARY (IMPORTANT FOR MAC) ===
    if (chromePath && chromePath !== 'google-chrome') {
        chromeOptions.setChromeBinaryPath(chromePath);
        console.log(`🎯 Using Chrome binary: ${chromePath}`);
    }

    // === PROFILE SETTINGS (EXACT SAME AS YOUR WORKING MANUAL COMMAND) ===
    chromeOptions.addArguments(`--user-data-dir=${userDataDir}`);
    chromeOptions.addArguments(`--profile-directory=${profileName}`);

    // === ESSENTIAL FLAGS ===
    chromeOptions.addArguments('--new-window');
    chromeOptions.addArguments('--no-first-run');
    chromeOptions.addArguments('--no-default-browser-check');
    chromeOptions.addArguments('--disable-default-apps');

    // === STEALTH SETTINGS ===
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');
    chromeOptions.addArguments('--disable-web-security');
    chromeOptions.addArguments('--disable-extensions');

    // === PREFERENCES ===
    chromeOptions.setUserPreferences({
        'credentials_enable_service': false,
        'profile.password_manager_enabled': false,
        'profile.default_content_setting_values.notifications': 2
    });

    console.log(`📂 Profile directory: ${profileName}`);
    console.log(`📁 User data directory: ${userDataDir}`);

    try {
        const driver = await new Builder()
            .forBrowser('chrome')
            .setChromeOptions(chromeOptions)
            .build();

        // === STEALTH INJECTION ===
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
        `);

        console.log('✅ Working profile driver created successfully!\n');
        return driver;

    } catch (error) {
        console.error('❌ Failed to create driver:', error.message);

        if (error.message.includes('user data directory is already in use')) {
            console.log('🔧 Solution: Close ALL Chrome windows first');
        } else if (error.message.includes('cannot find Chrome binary')) {
            console.log('🔧 Solution: Install Chrome or check installation path');
        }

        throw error;
    }
}


// === STEP 6: GENERATE WORKING AUTOMATION CODE ===
function generateWorkingCode(userDataDir, profileName) {
    console.log('🔍 STEP 6: Generating working automation code...\n');

    const workingCode = `
// WORKING CHROME PROFILE AUTOMATION
// This uses your verified profile settings

const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');

async function createWorkingProfileDriver() {
    console.log('🔑 Creating driver with WORKING profile settings...');
    
    const chromeOptions = new chrome.Options();
    
    // === VERIFIED PROFILE SETTINGS ===
    // chromeOptions.addArguments('--user-data-dir="${userDataDir.replace(/\\/g, '\\\\')}');
    chromeOptions.addArguments('--profile-directory=${profileName}');
    
    // === ESSENTIAL FLAGS ===
    chromeOptions.addArguments('--no-first-run');
    chromeOptions.addArguments('--no-default-browser-check');
    chromeOptions.addArguments('--disable-default-apps');
    
    // === STEALTH SETTINGS ===
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-infobars');
    
    console.log('📁 User Data Dir: ${userDataDir}');
    console.log('📂 Profile: ${profileName}');
    
    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();
    
    // Hide automation
    await driver.executeScript(\`
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
        });
    \`);
    
    console.log('✅ Working profile driver created!');
    return driver;
}

// Test the working driver
async function testWorkingDriver() {
    let driver;
    try {
        driver = await createWorkingProfileDriver();
        await driver.get('https://www.rivosafeguard.com/insight/');
        console.log('🎉 Profile automation is working!');
        await driver.sleep(5000);
    } catch (error) {
        console.error('❌ Error:', error.message);
    } finally {
        if (driver) await driver.quit();
    }
}

// Export for use in your main script
module.exports = { createWorkingProfileDriver, testWorkingDriver };

// Run test if called directly
if (require.main === module) {
    testWorkingDriver();
}
`;

    // Save the working code
    fs.writeFileSync('working-profile-driver.js', workingCode);

    console.log('✅ Working code saved as: working-profile-driver.js');
    console.log('🧪 Test it: node working-profile-driver.js');
    console.log('📝 Use createWorkingProfileDriver() in your main script\n');
}

// === MAIN DIAGNOSTIC FUNCTION ===
async function runCompleteProfileDiagnostic() {
    console.log('🏥 COMPLETE PROFILE DIAGNOSTIC');
    console.log('═══════════════════════════════════════\n');
    console.log('This will fix your Chrome profile issue step by step\n');

    try {
        // Step 1: Check Chrome processes
        const chromeRunning = await checkChromeProcesses();
        if (chromeRunning) {
            console.log('🛑 STOP HERE: Close Chrome first!\n');
            return;
        }

        // Step 2: Find profile path
        const userDataDir = findAndVerifyProfilePath();
        if (!userDataDir) {
            console.log('🛑 STOP HERE: Fix Chrome installation first!\n');
            return;
        }

        // Step 3: Verify Profile 1
        const profile1Check = verifySpecificProfile(userDataDir, 'Profile 1');

        let profileToUse = 'Profile 1';
        if (!profile1Check.exists && profile1Check.actualProfiles) {
            console.log('❌ Profile 1 not found. Available profiles:');
            profile1Check.actualProfiles.forEach(p => console.log(`   - ${p}`));

            // Use first available profile
            profileToUse = profile1Check.actualProfiles[0] || 'Default';
            console.log(`🔄 Using "${profileToUse}" instead\n`);
        }

        // Step 4: Generate manual command
        const manualCommand = generateManualCommand(userDataDir, profileToUse);

        console.log('⏸️  PAUSE HERE: Test the manual command above');
        console.log('❓ Did the manual command work? (y/n)');

        const readline = require('readline');
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        const manualWorked = await new Promise(resolve => {
            rl.question('Enter y if manual command worked, n if it failed: ', resolve);
        });
        rl.close();

        if (manualWorked.toLowerCase() !== 'y') {
            console.log('\n❌ Manual command failed. Profile issues:');
            console.log('1. Profile might not exist');
            console.log('2. Chrome might not be installed correctly');
            console.log('3. Permission issues');
            console.log('\n🔧 Fix the manual command first, then run this diagnostic again\n');
            return;
        }

        // Step 5: Test Selenium
        console.log('\n🧪 Manual command worked! Testing Selenium...\n');
        const seleniumWorked = await testSeleniumWithProfile(userDataDir, profileToUse);

        if (seleniumWorked) {
            // Step 6: Generate working code
            generateWorkingCode(userDataDir, profileToUse);

            console.log('🎉 DIAGNOSTIC COMPLETE - SUCCESS!\n');
            console.log('📋 SUMMARY:');
            console.log(`✅ Chrome profile path: ${userDataDir}`);
            console.log(`✅ Working profile: ${profileToUse}`);
            console.log(`✅ Manual command: Works`);
            console.log(`✅ Selenium test: Works`);
            console.log(`✅ Generated code: working-profile-driver.js`);
            console.log('\n🚀 NEXT STEPS:');
            console.log('1. Test: node working-profile-driver.js');
            console.log('2. Replace your driver creation with the working version');
            console.log('3. Run your automation normally\n');

        } else {
            console.log('\n❌ Selenium test failed even though manual command worked');
            console.log('🔧 This suggests a Selenium-specific issue');
            console.log('💡 Try updating selenium-webdriver: npm update selenium-webdriver\n');
        }

    } catch (error) {
        console.error('💥 Diagnostic failed:', error.message);
        console.log('\n🔧 Try these solutions:');
        console.log('1. Run as administrator/sudo');
        console.log('2. Check Chrome installation');
        console.log('3. Update Node.js and npm');
    }
}

async function testWorkingProfile() {
    let driver;

    try {
        console.log('🧪 Testing your working Profile 1...\n');

        // Create driver with your working profile
        driver = await createWorkingProfileDriver('Profile 1');
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        // Test 1: Verify profile
        console.log('🔍 Step 1: Verifying profile...');
        await driver.get('chrome://version/');
        await driver.sleep(2000);

        const versionInfo = await driver.executeScript('return document.documentElement.innerText;');
        if (versionInfo.includes('Profile 1')) {
            console.log('✅ Profile verification: Using Profile 1 correctly');
        } else {
            console.log('⚠️  Profile verification: May not be using Profile 1');
        }

        // Test 2: Check Rivo Safeguard login
        console.log('\n🌐 Step 2: Testing Rivo Safeguard login...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await driver.sleep(5000);

        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ SUCCESS! Found login elements - You are logged in!');
            console.log('🎉 PROFILE AUTOMATION IS WORKING!');

            // Show current URL and title
            const currentUrl = await driver.getCurrentUrl();
            const title = await driver.getTitle();
            console.log(`📄 Page: ${title}`);
            console.log(`🔗 URL: ${currentUrl}`);

            return true;

        } catch (error) {
            console.log('❌ Login elements not found');
            console.log('💡 This means you need to log into Rivo Safeguard in Profile 1 first');

            const currentUrl = await driver.getCurrentUrl();
            console.log(`📄 Current URL: ${currentUrl}`);

            return false;
        }

    } catch (error) {
        console.error('❌ Test failed:', error.message);
        return false;
    } finally {
        if (driver) {
            console.log('\n👀 Keeping browser open for 10 seconds to verify...');
            await driver.sleep(10000);
            await driver.quit();
            console.log('✅ Test browser closed');
        }
    }
}

async function runCompleteAutomation() {
    let driver;

    try {
        console.log('🚀 Running complete automation with working profile...\n');

        // Create working driver
        driver = await createWorkingProfileDriver('Profile 1');
        await driver.manage().window().setRect({ width: 1920, height: 1080 });

        // Navigate to Rivo Safeguard
        console.log('🌐 Navigating to Rivo Safeguard...');
        await driver.get('https://www.rivosafeguard.com/insight/');
        await driver.sleep(3000);

        // Check login status
        try {
            await driver.wait(until.elementLocated(By.css('.sch-container-left')), 10000);
            console.log('✅ Successfully logged in with Profile 1!');

            // Start your automation here
            console.log('🤖 Starting user creation automation...');

            // Navigate to user creation (your existing code)
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

            console.log('✅ Successfully navigated to user creation page!');
            console.log('🎉 Profile automation is working - you can now add your user creation logic!');

            // Keep browser open for manual verification
            console.log('\n👀 Keeping browser open for 15 seconds to verify navigation...');
            await driver.sleep(15000);

        } catch (error) {
            console.log('❌ Not logged in to Rivo Safeguard');
            console.log('💡 Please log into Rivo Safeguard with Profile 1 first');
        }

    } catch (error) {
        console.error('❌ Automation failed:', error.message);
    } finally {
        if (driver) {
            await driver.quit();
            console.log('✅ Automation browser closed');
        }
    }
}

// === RUN DIAGNOSTIC ===
if (require.main === module) {
    const args = process.argv.slice(2);

    if (args.includes('--test')) {
        console.log('🧪 TESTING MODE\n');
        testWorkingProfile();
    } else if (args.includes('--automation')) {
        console.log('🤖 AUTOMATION MODE\n');
        runCompleteAutomation();
    } else {
        console.log('🔧 WORKING CHROME PROFILE DRIVER\n');
        console.log('📋 Available commands:');
        console.log('  node working-driver.js --test        Test Profile 1');
        console.log('  node working-driver.js --automation  Run full automation');
        console.log('\n🚀 Running test by default...\n');
        testWorkingProfile();
    }
}

module.exports = {
    runCompleteProfileDiagnostic,
    checkChromeProcesses,
    findAndVerifyProfilePath,
    verifySpecificProfile,
    testSeleniumWithProfile
};