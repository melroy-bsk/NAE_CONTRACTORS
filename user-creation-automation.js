// Updated User Creation from Excel using First Name and Last Name columns
const { Builder, By, Key, until } = require('selenium-webdriver');
const XLSX = require('xlsx');
const fs = require('fs');
const assert = require('assert');

class UserCreationAutomator {
    constructor() {
        this.driver = null;
        this.contractors = [];
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

    // Read and parse Excel file using First Name and Last Name columns
    async loadContractorsFromExcel(filePath) {
        try {
            const workbook = XLSX.readFile(filePath);
            const sheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[sheetName];
            const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

            // Find column indices for required fields
            const headers = rawData[0] || [];
            const codeIndex = 0; // Assuming first column is always code
            const firstNameIndex = this.findColumnIndex(headers, ['first name', 'firstname', 'forename']);
            const lastNameIndex = this.findColumnIndex(headers, ['last name', 'lastname', 'surname']);
            const departmentIndex = this.findColumnIndex(headers, ['department', 'dept']);

            console.log('Column mapping:');
            console.log(`  Code: Column ${codeIndex} (${headers[codeIndex]})`);
            console.log(`  First Name: Column ${firstNameIndex} (${headers[firstNameIndex]})`);
            console.log(`  Last Name: Column ${lastNameIndex} (${headers[lastNameIndex]})`);
            console.log(`  Department: Column ${departmentIndex} (${headers[departmentIndex]})`);

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

                    if (department.toLowerCase() === 'security') {
                        // Generate username and password (code_firstname_lastname in lowercase)
                        const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                        this.contractors.push({
                            code: code,
                            fullName: fullName,
                            firstName: firstName,
                            lastName: lastName,
                            department: department || 'Security',
                            username: username,
                            password: username // Using same as username as requested
                        });
                    }
                }
            }

            console.log(`Loaded ${this.contractors.length} contractors from Excel file`);
            return this.contractors;
        } catch (error) {
            console.error('Error reading Excel file:', error);
            throw error;
        }
    }

    // Initialize WebDriver
    async initializeDriver() {
        this.driver = await new Builder().forBrowser('firefox').build();
        await this.driver.manage().window().setRect({ width: 1382, height: 736 });
        console.log('WebDriver initialized');
    }

    // Navigate to the initial page and setup
    async navigateToUserCreationPage() {
        await this.driver.get("https://www.rivosafeguard.com/insight/");

        // Follow the navigation sequence from the original script
        await this.driver.findElement(By.css(".sch-container-left")).click();
        await this.driver.findElement(By.css(".sch-app-launcher-button")).click();
        await this.driver.findElement(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text")).click();
        await this.driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();

        // Switch to frame
        await this.driver.switchTo().frame(0);
        console.log('Navigated to user creation page');
    }

    // Navigate back to add new user (for subsequent users)
    async navigateToAddNewUser() {
        // Switch back to default content first
        await this.driver.switchTo().defaultContent();

        // Click on the drawer item to show user creation page again
        await this.driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();

        // Switch to frame again
        await this.driver.switchTo().frame(0);
        console.log('Navigated back to add new user page');
    }

    // Create a single user
    async createUser(contractor, isFirstUser = false) {
        try {
            console.log(`Creating user: ${contractor.username} (${contractor.fullName})`);

            // If not the first user, navigate to add new user page
            if (!isFirstUser) {
                await this.navigateToAddNewUser();
                // Add a small delay to ensure page loads
                await this.driver.sleep(2000);
            }

            // Fill in Username
            const usernameField = await this.driver.findElement(By.id("Username"));
            await usernameField.clear();
            await usernameField.sendKeys(contractor.username);

            // Fill in Password (first time)
            const passwordField = await this.driver.findElement(By.name("Password"));
            await passwordField.clear();
            await passwordField.sendKeys(contractor.password);

            // Fill in Job Title
            const jobTitleField = await this.driver.findElement(By.name("JobTitle"));
            await jobTitleField.clear();
            await jobTitleField.sendKeys(contractor.department);

            // Fill in First Name
            const forenameField = await this.driver.findElement(By.id("Attributes.People.Forename"));
            await forenameField.clear();
            await forenameField.sendKeys(contractor.firstName);

            // Fill in Last Name
            const surnameField = await this.driver.findElement(By.id("Attributes.People.Surname"));
            await surnameField.clear();
            await surnameField.sendKeys(contractor.lastName);

            // Fill in Employee Number
            const employeeNumberField = await this.driver.findElement(By.id("Attributes.Users.EmployeeNumber"));
            await employeeNumberField.clear();
            await employeeNumberField.sendKeys(contractor.code);

            // Handle Hierarchy selection (following original script pattern)
            await this.driver.findElement(By.id("Hierarchies01zzz")).click();
            await this.driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await this.driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();

            // Select home location hierarchy
            try {
                await this.driver.findElement(By.css("#\\31 657_HomeLocationHierarchy > option:nth-child(2)")).click();
            } catch (error) {
                console.log('Home location hierarchy selector not found, continuing...');
            }

            // Additional hierarchy navigation
            await this.driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await this.driver.findElement(By.css(".HierarchyNode__SelectDiv-sc-1ytna50-2")).click();
            await this.driver.findElement(By.css(".DropdownDisplay__DropdownText-sc-6p7u3y-0 > span:nth-child(1)")).click();

            // Navigate hierarchy tree
            try {
                await this.driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
                await this.driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
                await this.driver.findElement(By.css(".HierarchyNode__NodeText-sc-1ytna50-1")).click();
            } catch (error) {
                console.log('Hierarchy navigation failed, continuing...');
            }

            // Add hierarchy groups/locations
            await this.driver.findElement(By.name("addHierarchyGroupsLocations")).click();
            await this.driver.findElement(By.css("#OverlayContainer > .StandardButton")).click();

            // Click on group tab
            try {
                await this.driver.findElement(By.id("GroupTabDiv_23890")).click();
            } catch (error) {
                console.log('Group tab not found, continuing...');
            }

            // Update password (second time in different field)
            const passwordField2 = await this.driver.findElement(By.name("Password"));
            await passwordField2.clear();
            await passwordField2.sendKeys(contractor.password);

            // Add hierarchy groups/locations again
            await this.driver.findElement(By.name("addHierarchyGroupsLocations")).click();
            await this.driver.findElement(By.css("#ManageHierarchyUGroupLocationsForm > .GroupTab:nth-child(1)")).click();

            // Select User Group - Third Party Staff
            const userGroupDropdown = await this.driver.findElement(By.id("UGroupID"));
            await userGroupDropdown.findElement(By.xpath("//option[. = 'Third Party Staff']")).click();

            // Additional hierarchy selections
            await this.driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();

            try {
                await this.driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
                await this.driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
                await this.driver.findElement(By.css(".HierarchyNode__NodeText-sc-1ytna50-1")).click();
                await this.driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();
            } catch (error) {
                console.log('Additional hierarchy selections failed, continuing...');
            }

            // Set Status of Employment to Current
            const statusDropdown = await this.driver.findElement(By.id("Attributes.Users.StatusOfEmployment"));
            await statusDropdown.findElement(By.xpath("//option[. = 'Current']")).click();

            // Set User Type to Limited access user
            const userTypeDropdown = await this.driver.findElement(By.id("UserTypeID"));
            await userTypeDropdown.findElement(By.xpath("//option[. = 'Limited access user']")).click();

            // Save the user
            await this.driver.findElement(By.name("save")).click();

            console.log(`✓ Successfully created user: ${contractor.username}`);

            // Wait a bit for the save to complete
            await this.driver.sleep(3000);

        } catch (error) {
            console.error(`✗ Failed to create user ${contractor.username}:`, error.message);
            throw error;
        }
    }

    // Create all users
    async createAllUsers() {
        console.log(`Starting to create ${this.contractors.length} users...`);

        for (let i = 0; i < this.contractors.length; i++) {
            const contractor = this.contractors[i];
            const isFirstUser = i === 0;

            try {
                await this.createUser(contractor, isFirstUser);
                console.log(`Progress: ${i + 1}/${this.contractors.length} users created`);

                // Add delay between users to avoid overwhelming the system
                if (i < this.contractors.length - 1) {
                    await this.driver.sleep(2000);
                }

            } catch (error) {
                console.error(`Failed to create user ${i + 1}: ${contractor.username}`);

                // Ask user if they want to continue or stop
                const readline = require('readline');
                const rl = readline.createInterface({
                    input: process.stdin,
                    output: process.stdout
                });

                const answer = await new Promise(resolve => {
                    rl.question('Do you want to continue with the next user? (y/n): ', resolve);
                });
                rl.close();

                if (answer.toLowerCase() !== 'y') {
                    console.log('Stopping automation as requested.');
                    break;
                }
            }
        }
    }

    // Create just the first user (for testing)
    async createFirstUser() {
        if (this.contractors.length === 0) {
            throw new Error('No contractors loaded');
        }

        console.log('Creating first user for testing...');
        await this.createUser(this.contractors[0], true);
    }

    // Cleanup
    async cleanup() {
        if (this.driver) {
            await this.driver.quit();
            console.log('WebDriver closed');
        }
    }
}

// Main execution function
async function main() {
    const automator = new UserCreationAutomator();

    try {
        // Load contractors from Excel file
        await automator.loadContractorsFromExcel('Copy of Contractors staff list.xlsx');

        // Show first few contractors for verification
        console.log('\nFirst 5 contractors to be created:');
        automator.contractors.slice(0, 5).forEach((contractor, index) => {
            console.log(`${index + 1}. ${contractor.fullName} -> ${contractor.username}`);
        });

        // Initialize driver and navigate to page
        await automator.initializeDriver();
        await automator.navigateToUserCreationPage();

        // Ask user what they want to do
        const readline = require('readline');
        const rl = readline.createInterface({
            input: process.stdin,
            output: process.stdout
        });

        const choice = await new Promise(resolve => {
            rl.question('\nWhat would you like to do?\n1. Create first user only (for testing)\n2. Create all users\nEnter choice (1 or 2): ', resolve);
        });
        rl.close();

        if (choice === '1') {
            await automator.createFirstUser();
        } else if (choice === '2') {
            await automator.createAllUsers();
        } else {
            console.log('Invalid choice. Exiting.');
        }

        console.log('Automation completed!');

    } catch (error) {
        console.error('Automation failed:', error);
    } finally {
        await automator.cleanup();
    }
}

// Export for use as module
module.exports = UserCreationAutomator;

// Run if called directly
if (require.main === module) {
    main().catch(console.error);
}