// Dry Run Version - User Creation Automation with Mock WebDriver
const XLSX = require('xlsx');
const fs = require('fs');

class MockWebDriver {
    constructor() {
        this.actions = [];
        this.currentFrame = 'default';
    }

    async get(url) {
        this.actions.push(`Navigate to: ${url}`);
    }

    async findElement(locator) {
        return new MockWebElement(locator, this);
    }

    async switchTo() {
        return {
            frame: (frameIndex) => {
                this.currentFrame = frameIndex;
                this.actions.push(`Switch to frame: ${frameIndex}`);
            },
            defaultContent: () => {
                this.currentFrame = 'default';
                this.actions.push('Switch to default content');
            }
        };
    }

    manage() {
        return {
            window: () => ({
                setRect: (dimensions) => {
                    this.actions.push(`Set window size: ${dimensions.width}x${dimensions.height}`);
                }
            })
        };
    }

    async sleep(ms) {
        this.actions.push(`Wait: ${ms}ms`);
    }

    async quit() {
        this.actions.push('Close browser');
    }

    getActions() {
        return this.actions;
    }
}

class MockWebElement {
    constructor(locator, driver) {
        this.locator = locator;
        this.driver = driver;
    }

    async click() {
        this.driver.actions.push(`Click element: ${this.locatorToString()}`);
    }

    async clear() {
        this.driver.actions.push(`Clear element: ${this.locatorToString()}`);
    }

    async sendKeys(text) {
        this.driver.actions.push(`Type '${text}' into: ${this.locatorToString()}`);
    }

    async findElement(locator) {
        return new MockWebElement(locator, this.driver);
    }

    locatorToString() {
        if (this.locator.id) return `#${this.locator.id}`;
        if (this.locator.css) return this.locator.css;
        if (this.locator.name) return `[name="${this.locator.name}"]`;
        if (this.locator.xpath) return `xpath: ${this.locator.xpath}`;
        return 'unknown locator';
    }
}

// Mock Builder for Selenium
class MockBuilder {
    forBrowser(browser) {
        return this;
    }

    async build() {
        console.log('🔍 DRY RUN MODE: Using Mock WebDriver (no browser will open)');
        return new MockWebDriver();
    }
}

class UserCreationAutomator {
    constructor(dryRun = false) {
        this.driver = null;
        this.contractors = [];
        this.dryRun = dryRun;
    }

    // Read and parse Excel file using existing First Name and Last Name columns
    async loadContractorsFromExcel(filePath) {
        try {
            console.log(`📊 Reading Excel file: ${filePath}`);

            // Check if file exists
            if (!fs.existsSync(filePath)) {
                throw new Error(`Excel file not found: ${filePath}`);
            }

            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            console.log(`📋 Found ${rawData.length - 1} rows in Excel (excluding header)`);
            console.log(`📑 Sheet name: ${sheetName}`);

            // Show header structure
            if (rawData.length > 0) {
                console.log('📌 Header columns:', rawData[0]);
            }

            // Find column indices for required fields
            const headers = rawData[0] || [];
            const codeIndex = 0; // Assuming first column is always code
            const firstNameIndex = this.findColumnIndex(headers, ['first name', 'firstname', 'forename']);
            const lastNameIndex = this.findColumnIndex(headers, ['last name', 'lastname', 'surname']);
            const departmentIndex = this.findColumnIndex(headers, ['department', 'dept']);

            console.log(`📍 Column mapping:`);
            console.log(`   Code: Column ${codeIndex} (${headers[codeIndex]})`);
            console.log(`   First Name: Column ${firstNameIndex} (${headers[firstNameIndex]})`);
            console.log(`   Last Name: Column ${lastNameIndex} (${headers[lastNameIndex]})`);
            console.log(`   Department: Column ${departmentIndex} (${headers[departmentIndex]})`);

            if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
                throw new Error('Required columns not found. Need: First Name, Last Name, Department');
            }

            // Skip header row and process data
            for (let i = 1; i < rawData.length; i++) {
                const row = rawData[i];
                if (row && row[codeIndex] && row[firstNameIndex] && row[lastNameIndex]) {
                    const code = row[codeIndex].toString().trim();
                    const firstName = row[firstNameIndex].toString().trim();
                    const lastName = row[lastNameIndex].toString().trim();
                    const department = row[departmentIndex] ? row[departmentIndex].toString().trim() : '';
                    const fullName = `${firstName} ${lastName}`;

                    console.log(`🔍 Row ${i}: ${fullName} - Department: ${department}`);

                    if (department.toLowerCase() === 'security') {
                        const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                        this.contractors.push({
                            code: code,
                            fullName: fullName,
                            firstName: firstName,
                            lastName: lastName,
                            department: department || 'Security',
                            username: username,
                            password: username
                        });

                        console.log(`✅ Added Security contractor: ${username}`);
                    } else {
                        console.log(`⏭️  Skipped (not Security department): ${fullName}`);
                    }
                } else {
                    console.log(`⚠️  Row ${i}: Missing required data (code, first name, or last name)`);
                }
            }

            console.log(`\n🎯 Final result: ${this.contractors.length} Security contractors loaded`);
            return this.contractors;
        } catch (error) {
            console.error('❌ Error reading Excel file:', error.message);
            throw error;
        }
    }

    // Helper function to find column index by possible header names
    findColumnIndex(headers, possibleNames) {
        for (let i = 0; i < headers.length; i++) {
            const header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
            if (possibleNames.some(name => header.includes(name))) {
                return i;
            }
        }
        return -1;
    }

    // Initialize WebDriver (with mock support)
    async initializeDriver() {
        if (this.dryRun) {
            this.driver = await new MockBuilder().forBrowser('firefox').build();
        } else {
            const { Builder } = require('selenium-webdriver');
            this.driver = await new Builder().forBrowser('firefox').build();
        }

        await this.driver.manage().window().setRect({ width: 1382, height: 736 });
        console.log('🌐 WebDriver initialized');
    }

    // Navigate to the initial page and setup
    async navigateToUserCreationPage() {
        const { By } = require('selenium-webdriver');

        await this.driver.get("https://www.rivosafeguard.com/insight/");
        await this.driver.findElement(By.css(".sch-container-left")).click();
        await this.driver.findElement(By.css(".sch-app-launcher-button")).click();
        await this.driver.findElement(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text")).click();
        await this.driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();
        await this.driver.switchTo().frame(0);

        console.log('🎯 Navigated to user creation page');
    }

    // Create a single user (with detailed logging)
    async createUser(contractor, isFirstUser = false) {
        const { By } = require('selenium-webdriver');

        try {
            console.log(`\n👤 Creating user: ${contractor.username} (${contractor.fullName})`);
            console.log(`   📝 Details: ${contractor.firstName} ${contractor.lastName}, Dept: ${contractor.department}, Code: ${contractor.code}`);

            if (!isFirstUser) {
                console.log('🔄 Navigating to add new user page...');
                await this.driver.switchTo().defaultContent();
                await this.driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();
                await this.driver.switchTo().frame(0);
                await this.driver.sleep(2000);
            }

            // Fill form fields with detailed logging
            await this.fillFormField("Username", "Username", contractor.username);
            await this.fillFormField("Password", "Password", contractor.password);
            await this.fillFormField("Job Title", "JobTitle", contractor.department);
            await this.fillFormField("First Name", "Attributes.People.Forename", contractor.firstName);
            await this.fillFormField("Last Name", "Attributes.People.Surname", contractor.lastName);
            await this.fillFormField("Employee Number", "Attributes.Users.EmployeeNumber", contractor.code);

            // Hierarchy operations
            console.log('🏢 Setting up hierarchy and groups...');
            await this.handleHierarchySelection();

            // User settings
            console.log('⚙️  Configuring user settings...');
            await this.setUserConfiguration();

            // Save user
            await this.driver.findElement(By.name("save")).click();
            console.log(`✅ Successfully created user: ${contractor.username}`);

            await this.driver.sleep(3000);

        } catch (error) {
            console.error(`❌ Failed to create user ${contractor.username}:`, error.message);
            throw error;
        }
    }

    async fillFormField(label, selector, value) {
        const { By } = require('selenium-webdriver');

        const element = await this.driver.findElement(
            selector.includes('.') ? By.id(selector) : By.name(selector)
        );
        await element.clear();
        await element.sendKeys(value);
        console.log(`   📝 ${label}: ${value}`);
    }

    async handleHierarchySelection() {
        const { By } = require('selenium-webdriver');

        try {
            await this.driver.findElement(By.id("Hierarchies01zzz")).click();
            await this.driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await this.driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();

            // Additional hierarchy steps...
            console.log('   🏢 Hierarchy configuration completed');
        } catch (error) {
            console.log('   ⚠️  Some hierarchy selections failed, continuing...');
        }
    }

    async setUserConfiguration() {
        const { By } = require('selenium-webdriver');

        try {
            // Set User Group to Third Party Staff
            const userGroupDropdown = await this.driver.findElement(By.id("UGroupID"));
            await userGroupDropdown.findElement(By.xpath("//option[. = 'Third Party Staff']")).click();
            console.log('   👥 User Group: Third Party Staff');

            // Set Status of Employment to Current
            const statusDropdown = await this.driver.findElement(By.id("Attributes.Users.StatusOfEmployment"));
            await statusDropdown.findElement(By.xpath("//option[. = 'Current']")).click();
            console.log('   💼 Employment Status: Current');

            // Set User Type to Limited access user
            const userTypeDropdown = await this.driver.findElement(By.id("UserTypeID"));
            await userTypeDropdown.findElement(By.xpath("//option[. = 'Limited access user']")).click();
            console.log('   🔐 User Type: Limited access user');

        } catch (error) {
            console.log('   ⚠️  Some configuration settings failed, continuing...');
        }
    }

    // Show summary of what would be done
    showDryRunSummary() {
        console.log('\n📊 DRY RUN SUMMARY');
        console.log('='.repeat(50));
        console.log(`📈 Total contractors found: ${this.contractors.length}`);

        if (this.contractors.length === 0) {
            console.log('⚠️  No Security department contractors found to create!');
            return;
        }

        console.log('\n👥 Users that would be created:');
        this.contractors.forEach((contractor, index) => {
            console.log(`${index + 1}. ${contractor.fullName}`);
            console.log(`   Username: ${contractor.username}`);
            console.log(`   Password: ${contractor.password}`);
            console.log(`   Department: ${contractor.department}`);
            console.log(`   Employee Code: ${contractor.code}`);
            console.log('');
        });

        if (this.dryRun && this.driver instanceof MockWebDriver) {
            console.log('\n🎬 Browser actions that would be performed:');
            const actions = this.driver.getActions();
            actions.slice(0, 10).forEach((action, index) => {
                console.log(`${index + 1}. ${action}`);
            });
            if (actions.length > 10) {
                console.log(`   ... and ${actions.length - 10} more actions`);
            }
        }
    }

    async cleanup() {
        if (this.driver) {
            await this.driver.quit();
            console.log('🧹 Cleanup completed');
        }
    }
}

// Main function with dry run support
async function main() {
    const args = process.argv.slice(2);
    const dryRun = args.includes('--dry-run') || args.includes('-d');

    if (dryRun) {
        console.log('🔍 DRY RUN MODE ACTIVATED - No browser will open');
        console.log('='.repeat(50));
    }

    const automator = new UserCreationAutomator(dryRun);

    try {
        // Load contractors from Excel file
        await automator.loadContractorsFromExcel('Copy of Contractors staff list.xlsx');

        if (dryRun) {
            // In dry run mode, just show what would happen
            automator.showDryRunSummary();

            // Optionally initialize mock driver to show browser actions
            await automator.initializeDriver();
            await automator.navigateToUserCreationPage();

            if (automator.contractors.length > 0) {
                console.log('\n🎭 Simulating creation of first user...');
                await automator.createUser(automator.contractors[0], true);
            }

            automator.showDryRunSummary();

        } else {
            // Normal execution
            console.log('\n🚀 LIVE MODE - Browser will open');
            console.log('First 5 contractors to be created:');
            automator.contractors.slice(0, 5).forEach((contractor, index) => {
                console.log(`${index + 1}. ${contractor.fullName} -> ${contractor.username}`);
            });

            await automator.initializeDriver();
            await automator.navigateToUserCreationPage();

            const readline = require('readline');
            const rl = readline.createInterface({
                input: process.stdin,
                output: process.stdout
            });

            const choice = await new Promise(resolve => {
                rl.question('\nWhat would you like to do?\n1. Create first user only\n2. Create all users\nEnter choice (1 or 2): ', resolve);
            });
            rl.close();

            if (choice === '1') {
                await automator.createUser(automator.contractors[0], true);
            } else if (choice === '2') {
                for (let i = 0; i < automator.contractors.length; i++) {
                    await automator.createUser(automator.contractors[i], i === 0);
                }
            }
        }

        console.log('✅ Process completed!');

    } catch (error) {
        console.error('❌ Process failed:', error.message);
    } finally {
        await automator.cleanup();
    }
}

if (require.main === module) {
    main().catch(console.error);
}

module.exports = UserCreationAutomator;