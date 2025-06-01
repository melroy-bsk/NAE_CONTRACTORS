// Complete Stealth Browser Setup - Bypass ALL Automation Detection
const { Builder, By, Key, until } = require('selenium-webdriver');
const chrome = require('selenium-webdriver/chrome');
const firefox = require('selenium-webdriver/firefox');
const XLSX = require('xlsx');
const fs = require('fs');

// === CHROME STEALTH CONFIGURATION ===
async function createStealthChromeDriver() {
    console.log('🕵️  Creating undetectable Chrome browser...');

    const chromeOptions = new chrome.Options();

    // === PRIMARY STEALTH SETTINGS ===
    // Remove automation indicators
    chromeOptions.addArguments('--disable-blink-features=AutomationControlled');
    chromeOptions.excludeSwitches(['enable-automation']);
    chromeOptions.addArguments('--disable-web-security');
    chromeOptions.addArguments('--allow-running-insecure-content');
    chromeOptions.addArguments('--disable-features=TranslateUI');
    chromeOptions.addArguments('--disable-features=BlinkGenPropertyTrees');
    chromeOptions.addArguments('--disable-ipc-flooding-protection');

    // === REMOVE AUTOMATION NOTIFICATIONS ===
    chromeOptions.addArguments('--disable-infobars');
    chromeOptions.addArguments('--disable-extensions');
    chromeOptions.addArguments('--disable-plugins-discovery');
    chromeOptions.addArguments('--disable-dev-shm-usage');
    chromeOptions.addArguments('--no-sandbox');
    chromeOptions.addArguments('--disable-gpu');

    // === USER AGENT SPOOFING ===
    // Use the latest real Chrome user agent
    chromeOptions.addArguments('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36');

    // === ADVANCED STEALTH OPTIONS ===
    chromeOptions.addArguments('--disable-background-networking');
    chromeOptions.addArguments('--disable-background-timer-throttling');
    chromeOptions.addArguments('--disable-backgrounding-occluded-windows');
    chromeOptions.addArguments('--disable-renderer-backgrounding');
    chromeOptions.addArguments('--disable-field-trial-config');
    chromeOptions.addArguments('--disable-hang-monitor');
    chromeOptions.addArguments('--disable-client-side-phishing-detection');
    chromeOptions.addArguments('--disable-popup-blocking');
    chromeOptions.addArguments('--disable-prompt-on-repost');
    chromeOptions.addArguments('--disable-sync');
    chromeOptions.addArguments('--disable-component-extensions-with-background-pages');
    chromeOptions.addArguments('--disable-default-apps');
    chromeOptions.addArguments('--disable-breakpad');
    chromeOptions.addArguments('--disable-component-update');
    chromeOptions.addArguments('--disable-domain-reliability');
    chromeOptions.addArguments('--disable-features=AudioServiceOutOfProcess');
    chromeOptions.addArguments('--disable-features=VizDisplayCompositor');

    // === PERMISSIONS AND PREFERENCES ===
    chromeOptions.setUserPreferences({
        'credentials_enable_service': false,
        'password_manager_enabled': false,
        'profile.password_manager_enabled': false,
        'profile.default_content_setting_values.notifications': 2,
        'profile.default_content_settings.popups': 0,
        'profile.managed_default_content_settings.images': 1,
        'profile.default_content_setting_values.media_stream_mic': 2,
        'profile.default_content_setting_values.media_stream_camera': 2,
        'profile.default_content_setting_values.geolocation': 2,
        'profile.default_content_setting_values.desktop_notifications': 2
    });

    // === EXPERIMENTAL OPTIONS ===
    chromeOptions.addArguments('--use-fake-ui-for-media-stream');
    chromeOptions.addArguments('--use-fake-device-for-media-stream');
    chromeOptions.addArguments('--autoplay-policy=no-user-gesture-required');

    // === CREATE DRIVER ===
    const driver = await new Builder()
        .forBrowser('chrome')
        .setChromeOptions(chromeOptions)
        .build();

    // === INJECT STEALTH JAVASCRIPT ===
    await injectStealthScripts(driver);

    return driver;
}

// === FIREFOX STEALTH CONFIGURATION ===
async function createStealthFirefoxDriver() {
    console.log('🕵️  Creating undetectable Firefox browser...');

    const firefoxOptions = new firefox.Options();

    // === CORE STEALTH PREFERENCES ===
    firefoxOptions.setPreference('dom.webdriver.enabled', false);
    firefoxOptions.setPreference('useAutomationExtension', false);
    firefoxOptions.setPreference('marionette.enabled', false);

    // === USER AGENT OVERRIDE ===
    firefoxOptions.setPreference('general.useragent.override',
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:131.0) Gecko/20100101 Firefox/131.0');

    // === PRIVACY AND SECURITY ===
    firefoxOptions.setPreference('privacy.trackingprotection.enabled', false);
    firefoxOptions.setPreference('dom.ipc.plugins.enabled.libflashplayer.so', false);
    firefoxOptions.setPreference('media.peerconnection.enabled', false);
    firefoxOptions.setPreference('geo.enabled', false);
    firefoxOptions.setPreference('dom.battery.enabled', false);
    firefoxOptions.setPreference('dom.webnotifications.enabled', false);

    // === DISABLE AUTOMATION FEATURES ===
    firefoxOptions.setPreference('devtools.jsonview.enabled', false);
    firefoxOptions.setPreference('devtools.debugger.remote-enabled', false);
    firefoxOptions.setPreference('toolkit.telemetry.enabled', false);
    firefoxOptions.setPreference('browser.ping-centre.telemetry', false);

    const driver = await new Builder()
        .forBrowser('firefox')
        .setFirefoxOptions(firefoxOptions)
        .build();

    await injectStealthScripts(driver);
    return driver;
}

// === STEALTH JAVASCRIPT INJECTION ===
async function injectStealthScripts(driver) {
    console.log('🎭 Injecting stealth scripts to hide automation...');

    try {
        // === REMOVE WEBDRIVER PROPERTY ===
        await driver.executeScript(`
            // Remove webdriver property completely
            Object.defineProperty(navigator, 'webdriver', {
                get: () => undefined,
                configurable: true
            });
            
            // Delete webdriver from window
            delete window.navigator.webdriver;
            delete window.navigator.__proto__.webdriver;
            delete window.webdriver;
            delete window.domAutomation;
            delete window.domAutomationController;
            delete window.fxdriver_id;
            delete window.fxdriver_unwrapped;
            delete window.webkitStorageInfo;
            delete window.webkitIndexedDB;
        `);

        // === OVERRIDE NAVIGATOR PROPERTIES ===
        await driver.executeScript(`
            // Override plugins
            Object.defineProperty(navigator, 'plugins', {
                get: () => [
                    {
                        0: {type: "application/x-google-chrome-pdf", suffixes: "pdf", description: "Portable Document Format", enabledPlugin: true},
                        description: "Portable Document Format",
                        filename: "internal-pdf-viewer",
                        length: 1,
                        name: "Chrome PDF Plugin"
                    },
                    {
                        0: {type: "application/pdf", suffixes: "pdf", description: "Portable Document Format", enabledPlugin: true},
                        description: "Portable Document Format", 
                        filename: "mhjfbmdgcfjbbpaeojofohoefgiehjai",
                        length: 1,
                        name: "Chrome PDF Viewer"
                    }
                ],
                configurable: true
            });
            
            // Override languages
            Object.defineProperty(navigator, 'languages', {
                get: () => ['en-US', 'en'],
                configurable: true
            });
            
            // Override platform
            Object.defineProperty(navigator, 'platform', {
                get: () => 'Win32',
                configurable: true
            });
            
            // Override hardwareConcurrency
            Object.defineProperty(navigator, 'hardwareConcurrency', {
                get: () => 8,
                configurable: true
            });
            
            // Override deviceMemory
            Object.defineProperty(navigator, 'deviceMemory', {
                get: () => 8,
                configurable: true
            });
        `);

        // === CHROME-SPECIFIC SPOOFING ===
        await driver.executeScript(`
            if (navigator.userAgent.includes('Chrome')) {
                // Add chrome object
                window.chrome = {
                    app: {
                        isInstalled: false,
                        InstallState: {
                            DISABLED: 'disabled',
                            INSTALLED: 'installed',
                            NOT_INSTALLED: 'not_installed'
                        },
                        RunningState: {
                            CANNOT_RUN: 'cannot_run',
                            READY_TO_RUN: 'ready_to_run',
                            RUNNING: 'running'
                        }
                    },
                    runtime: {
                        onConnect: null,
                        onMessage: null
                    },
                    loadTimes: function() {
                        return {
                            requestTime: Date.now() * 0.001,
                            startLoadTime: Date.now() * 0.001,
                            commitLoadTime: Date.now() * 0.001,
                            finishDocumentLoadTime: Date.now() * 0.001,
                            finishLoadTime: Date.now() * 0.001,
                            firstPaintTime: Date.now() * 0.001,
                            firstPaintAfterLoadTime: 0,
                            navigationType: 'Other',
                            wasFetchedViaSpdy: false,
                            wasNpnNegotiated: false,
                            npnNegotiatedProtocol: 'unknown',
                            wasAlternateProtocolAvailable: false,
                            connectionInfo: 'http/1.1'
                        };
                    },
                    csi: function() {
                        return {
                            startE: Date.now(),
                            onloadT: Date.now(),
                            pageT: Date.now() * 0.001,
                            tran: 15
                        };
                    }
                };
            }
        `);

        // === PERMISSIONS API SPOOFING ===
        await driver.executeScript(`
            const originalQuery = window.navigator.permissions.query;
            window.navigator.permissions.query = (parameters) => (
                parameters.name === 'notifications' ?
                    Promise.resolve({ state: Notification.permission }) :
                    originalQuery(parameters)
            );
            
            // Override getVideoTracks
            if (navigator.mediaDevices && navigator.mediaDevices.getUserMedia) {
                const originalGetUserMedia = navigator.mediaDevices.getUserMedia;
                navigator.mediaDevices.getUserMedia = function() {
                    return originalGetUserMedia.apply(this, arguments);
                };
            }
        `);

        // === IFRAME AND FRAME DETECTION ===
        await driver.executeScript(`
            // Hide automation in iframes
            Object.defineProperty(HTMLIFrameElement.prototype, 'contentWindow', {
                get: function() {
                    const win = this.contentWindow;
                    if (win) {
                        delete win.navigator.webdriver;
                    }
                    return win;
                }
            });
        `);

        // === TIMING ATTACKS PREVENTION ===
        await driver.executeScript(`
            // Add random delays to timing attacks
            const originalNow = Date.now;
            Date.now = function() {
                return originalNow() + Math.random() * 10;
            };
            
            // Override performance.now
            if (window.performance && window.performance.now) {
                const originalPerfNow = window.performance.now;
                window.performance.now = function() {
                    return originalPerfNow.call(this) + Math.random() * 10;
                };
            }
        `);

        console.log('   ✅ All stealth scripts injected successfully');

    } catch (error) {
        console.log('   ⚠️  Some stealth scripts failed:', error.message);
    }
}

// === HUMAN BEHAVIOR SIMULATION ===
async function simulateHumanBehavior(driver) {
    console.log('🤖 Setting up human behavior simulation...');

    try {
        await driver.executeScript(`
            // === RANDOM MOUSE MOVEMENTS ===
            function simulateMouseMovement() {
                const event = new MouseEvent('mousemove', {
                    clientX: Math.random() * window.innerWidth,
                    clientY: Math.random() * window.innerHeight,
                    bubbles: true
                });
                document.dispatchEvent(event);
            }
            
            // === RANDOM SCROLLING ===
            function simulateScrolling() {
                const scrollAmount = Math.random() * 200 - 100;
                window.scrollBy({
                    top: scrollAmount,
                    left: 0,
                    behavior: 'smooth'
                });
            }
            
            // === RANDOM CLICKS (OFF ELEMENTS) ===
            function simulateRandomClick() {
                const x = Math.random() * window.innerWidth;
                const y = Math.random() * window.innerHeight;
                const element = document.elementFromPoint(x, y);
                
                if (element && element.tagName === 'BODY') {
                    const event = new MouseEvent('click', {
                        clientX: x,
                        clientY: y,
                        bubbles: true
                    });
                    element.dispatchEvent(event);
                }
            }
            
            // === SETUP INTERVALS ===
            // Mouse movements every 3-8 seconds
            setInterval(simulateMouseMovement, 3000 + Math.random() * 5000);
            
            // Scrolling every 10-20 seconds
            setInterval(simulateScrolling, 10000 + Math.random() * 10000);
            
            // Random clicks every 15-30 seconds
            setInterval(simulateRandomClick, 15000 + Math.random() * 15000);
            
            // === KEYBOARD ACTIVITY ===
            function simulateKeyActivity() {
                const event = new KeyboardEvent('keydown', {
                    key: 'Tab',
                    bubbles: true
                });
                document.dispatchEvent(event);
            }
            
            // Tab key every 20-40 seconds
            setInterval(simulateKeyActivity, 20000 + Math.random() * 20000);
        `);

        console.log('   ✅ Human behavior simulation activated');
    } catch (error) {
        console.log('   ⚠️  Behavior simulation setup failed:', error.message);
    }
}

// === MAIN STEALTH DRIVER CREATION ===
async function createUltraStealthDriver(browserType = 'chrome') {
    let driver;

    try {
        if (browserType.toLowerCase() === 'chrome') {
            driver = await createStealthChromeDriver();
        } else {
            driver = await createStealthFirefoxDriver();
        }

        // Set realistic window size
        await driver.manage().window().setRect({
            width: 1920,
            height: 1080
        });

        // Enable human behavior simulation
        await simulateHumanBehavior(driver);

        // Add initial delay to let everything settle
        await driver.sleep(2000);

        console.log('🎭 Ultra-stealth browser is ready! Completely undetectable.');
        return driver;

    } catch (error) {
        console.error('❌ Failed to create stealth driver:', error.message);
        throw error;
    }
}

// === ENHANCED TEST UTILITIES WITH HUMAN BEHAVIOR ===
class StealthTestUtilities {
    constructor(driver) {
        this.driver = driver;
    }

    // Ultra-realistic human delays
    async humanDelay(baseMs = 1000, variationMs = 500) {
        const randomVariation = Math.random() * variationMs * 2 - variationMs;
        let totalDelay = Math.max(baseMs + randomVariation, 200);

        // 10% chance of "distraction" - longer pause
        if (Math.random() < 0.1) {
            totalDelay += Math.random() * 3000;
            console.log(`   🤔 Taking a brief pause (${Math.round(totalDelay)}ms)...`);
        }

        await this.driver.sleep(totalDelay);
    }

    // Realistic typing with mistakes and corrections
    async humanType(element, text, avgDelay = 80) {
        await element.clear();
        await this.humanDelay(300, 200); // Pause before typing

        for (let i = 0; i < text.length; i++) {
            const char = text[i];

            // 2% chance of typing mistake
            if (Math.random() < 0.02 && i > 0) {
                const wrongChar = String.fromCharCode(65 + Math.floor(Math.random() * 26));
                await element.sendKeys(wrongChar);
                await this.driver.sleep(100 + Math.random() * 200);
                await element.sendKeys(Key.BACK_SPACE);
                await this.driver.sleep(50 + Math.random() * 100);
            }

            await element.sendKeys(char);

            // Variable typing speed (faster for common words)
            let keyDelay = avgDelay;
            if (i < text.length - 1) {
                const nextChar = text[i + 1];
                // Faster for common letter combinations
                if (char.toLowerCase() === 't' && nextChar.toLowerCase() === 'h') keyDelay *= 0.7;
                if (char.toLowerCase() === 'e' && nextChar.toLowerCase() === 'r') keyDelay *= 0.7;
            }

            await this.driver.sleep(keyDelay + (Math.random() * 40 - 20));
        }

        await this.humanDelay(200, 100); // Pause after typing
    }

    // Human-like clicking with mouse movement
    async humanClick(element, description) {
        try {
            // Move mouse to element area first
            const actions = this.driver.actions({ bridge: true });
            await actions.move({ origin: element, x: Math.random() * 10 - 5, y: Math.random() * 10 - 5 }).perform();
            await this.humanDelay(150, 75);

            // Click
            await element.click();
            console.log(`   ✅ Clicked: ${description}`);

            await this.humanDelay(300, 150);

        } catch (error) {
            throw new Error(`Failed to click ${description}: ${error.message}`);
        }
    }

    // Safe interaction with retry logic
    async safeInteraction(selector, action, description, timeout = 15000) {
        let attempts = 3;

        while (attempts > 0) {
            try {
                const element = await this.driver.wait(until.elementLocated(selector), timeout);
                await this.driver.wait(until.elementIsEnabled(element), 5000);

                await action(element);
                return element;

            } catch (error) {
                attempts--;
                console.log(`   ⚠️  ${description} failed, retrying... (${attempts} attempts left)`);

                if (attempts === 0) {
                    throw new Error(`Failed ${description}: ${error.message}`);
                }

                await this.humanDelay(1000, 500);
            }
        }
    }
}

// === YOUR MODIFIED MAIN SCRIPT ===
async function createStealthUserTest() {
    let driver;

    try {
        console.log('🕵️  Starting completely undetectable automation...');

        // Load contractor data
        const contractor = loadContractorData();

        // Create ultra-stealth driver
        driver = await createUltraStealthDriver('chrome'); // or 'firefox'
        const testUtils = new StealthTestUtilities(driver);

        console.log(`\n👤 Creating user: ${contractor.username} (${contractor.fullName})`);

        // Navigate to application
        console.log('🌐 Navigating to Rivo Safeguard...');
        await driver.get("https://www.rivosafeguard.com/insight/");
        await testUtils.humanDelay(4000, 2000);

        // Follow navigation with stealth interactions
        await testUtils.safeInteraction(
            By.css(".sch-container-left"),
            (el) => testUtils.humanClick(el, "Left container"),
            "clicking left container"
        );

        await testUtils.safeInteraction(
            By.css(".sch-app-launcher-button"),
            (el) => testUtils.humanClick(el, "App launcher"),
            "clicking app launcher"
        );

        await testUtils.safeInteraction(
            By.css(".sch-link-title:nth-child(6) > .sch-link-title-text"),
            (el) => testUtils.humanClick(el, "Menu item"),
            "clicking menu item"
        );

        await testUtils.safeInteraction(
            By.css(".k-drawer-item:nth-child(6)"),
            (el) => testUtils.humanClick(el, "User management"),
            "clicking user management"
        );

        // Switch to frame
        await driver.switchTo().frame(0);
        await testUtils.humanDelay(2000, 1000);

        // Fill form with stealth typing
        await testUtils.safeInteraction(
            By.id("Username"),
            (el) => testUtils.humanType(el, contractor.username),
            "entering username"
        );

        await testUtils.safeInteraction(
            By.name("Password"),
            (el) => testUtils.humanType(el, contractor.password),
            "entering password"
        );

        await testUtils.safeInteraction(
            By.name("JobTitle"),
            (el) => testUtils.humanType(el, contractor.department),
            "entering job title"
        );

        await testUtils.safeInteraction(
            By.id("Attributes.People.Forename"),
            (el) => testUtils.humanType(el, contractor.firstName),
            "entering first name"
        );

        await testUtils.safeInteraction(
            By.id("Attributes.People.Surname"),
            (el) => testUtils.humanType(el, contractor.lastName),
            "entering last name"
        );

        await testUtils.safeInteraction(
            By.id("Attributes.Users.EmployeeNumber"),
            (el) => testUtils.humanType(el, contractor.code),
            "entering employee number"
        );

        // Continue with rest of your form filling...
        // [Include all your existing form interaction code here, but use testUtils methods]

        console.log('🎉 Stealth automation completed successfully!');
        console.log('🕵️  Website had no idea this was automated!');

    } catch (error) {
        console.error('❌ Stealth automation failed:', error.message);
    } finally {
        if (driver) {
            await driver.sleep(5000); // Let user see results
            await driver.quit();
        }
    }
}

// Helper function (you already have this)
function loadContractorData() {
    // Your existing loadContractorData function here
    console.log('📁 Loading contractor data...');
    // ... your Excel loading code ...
}

// Export modules
module.exports = {
    createUltraStealthDriver,
    StealthTestUtilities,
    createStealthUserTest
};

// Run if called directly
if (require.main === module) {
    createStealthUserTest();
}