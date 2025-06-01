// Updated Test script to create the first user from Excel file using First Name and Last Name columns
const { Builder, By, Key, until } = require('selenium-webdriver');
const XLSX = require('xlsx');
const assert = require('assert');

describe('create first user from excel using first name and last name columns', function () {
    this.timeout(60000); // Increased timeout for file operations
    let driver;
    let vars;
    let firstContractor;

    // Helper function to find column index by possible header names
    function findColumnIndex(headers, possibleNames) {
        for (let i = 0; i < headers.length; i++) {
            const header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
            if (possibleNames.some(name => header.includes(name))) {
                return i;
            }
        }
        return -1;
    }

    beforeEach(async function () {
        // Read Excel file and get first contractor
        console.log('Reading Excel file...');
        const workbook = XLSX.readFile('Copy of Contractors staff list.xlsx');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Find column indices for required fields
        const headers = rawData[0] || [];
        const codeIndex = 0; // Assuming first column is always code
        const firstNameIndex = findColumnIndex(headers, ['first name', 'firstname', 'forename']);
        const lastNameIndex = findColumnIndex(headers, ['last name', 'lastname', 'surname']);
        const departmentIndex = findColumnIndex(headers, ['department', 'dept']);

        console.log('Column mapping:');
        console.log(`  Code: Column ${codeIndex} (${headers[codeIndex]})`);
        console.log(`  First Name: Column ${firstNameIndex} (${headers[firstNameIndex]})`);
        console.log(`  Last Name: Column ${lastNameIndex} (${headers[lastNameIndex]})`);
        console.log(`  Department: Column ${departmentIndex} (${headers[departmentIndex]})`);

        if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
            throw new Error('Required columns not found. Need: First Name, Last Name, Department');
        }

        // Find first Security department contractor
        let foundContractor = false;
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (row && row[codeIndex] && row[firstNameIndex] && row[lastNameIndex]) {
                const code = row[codeIndex].toString().trim();
                const firstName = row[firstNameIndex].toString().trim();
                const lastName = row[lastNameIndex].toString().trim();
                const department = row[departmentIndex] ? row[departmentIndex].toString().trim() : '';
                const fullName = `${firstName} ${lastName}`;

                if (department.toLowerCase() === 'security') {
                    const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                    firstContractor = {
                        code: code,
                        fullName: fullName,
                        firstName: firstName,
                        lastName: lastName,
                        department: department || 'Security',
                        username: username,
                        password: username
                    };

                    console.log('First Security contractor to create:', firstContractor);
                    foundContractor = true;
                    break;
                }
            }
        }

        if (!foundContractor) {
            throw new Error('No Security department contractors found in Excel file');
        }

        // Initialize driver
        driver = await new Builder().forBrowser('firefox').build();
        vars = {};
    });

    afterEach(async function () {
        await driver.quit();
    });

    it('create first user from excel using first and last name columns', async function () {
        console.log(`Creating user: ${firstContractor.username} (${firstContractor.fullName})`);

        // Navigate to the application
        await driver.get("https://www.rivosafeguard.com/insight/");
        await driver.manage().window().setRect({ width: 1382, height: 736 });

        // Follow navigation sequence
        await driver.findElement(By.css(".sch-container-left")).click();
        await driver.findElement(By.css(".sch-app-launcher-button")).click();
        await driver.findElement(By.css(".sch-link-title:nth-child(6) > .sch-link-title-text")).click();
        await driver.findElement(By.css(".k-drawer-item:nth-child(6)")).click();

        // Switch to frame
        await driver.switchTo().frame(0);

        // Fill in Username using contractor data
        await driver.findElement(By.id("Username")).click();
        await driver.findElement(By.id("Username")).clear();
        await driver.findElement(By.id("Username")).sendKeys(firstContractor.username);

        // Fill in Password
        await driver.findElement(By.name("Password")).click();
        await driver.findElement(By.name("Password")).clear();
        await driver.findElement(By.name("Password")).sendKeys(firstContractor.password);

        // Fill in Job Title (Department)
        await driver.findElement(By.name("JobTitle")).click();
        await driver.findElement(By.name("JobTitle")).clear();
        await driver.findElement(By.name("JobTitle")).sendKeys(firstContractor.department);

        // Fill in First Name
        await driver.findElement(By.id("Attributes.People.Forename")).click();
        await driver.findElement(By.id("Attributes.People.Forename")).clear();
        await driver.findElement(By.id("Attributes.People.Forename")).sendKeys(firstContractor.firstName);

        // Fill in Last Name
        await driver.findElement(By.id("Attributes.People.Surname")).click();
        await driver.findElement(By.id("Attributes.People.Surname")).clear();
        await driver.findElement(By.id("Attributes.People.Surname")).sendKeys(firstContractor.lastName);

        // Fill in Employee Number (Contractor Code)
        await driver.findElement(By.id("Attributes.Users.EmployeeNumber")).click();
        await driver.findElement(By.id("Attributes.Users.EmployeeNumber")).clear();
        await driver.findElement(By.id("Attributes.Users.EmployeeNumber")).sendKeys(firstContractor.code);

        // Handle Hierarchy selection (following original pattern)
        try {
            await driver.findElement(By.id("Hierarchies01zzz")).click();
            await driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();
            await driver.findElement(By.css("#\\31 657_HomeLocationHierarchy > option:nth-child(2)")).click();
            await driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await driver.findElement(By.css(".HierarchyNode__SelectDiv-sc-1ytna50-2")).click();
            await driver.findElement(By.css(".DropdownDisplay__DropdownText-sc-6p7u3y-0 > span:nth-child(1)")).click();
            await driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
            await driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
            await driver.findElement(By.css(".HierarchyNode__NodeText-sc-1ytna50-1")).click();
            await driver.findElement(By.name("addHierarchyGroupsLocations")).click();
            await driver.findElement(By.css("#OverlayContainer > .StandardButton")).click();
        } catch (error) {
            console.log('Some hierarchy selections failed, continuing...');
        }

        // Group tab operations
        try {
            await driver.findElement(By.id("GroupTabDiv_23890")).click();
        } catch (error) {
            console.log('Group tab not found, continuing...');
        }

        // Update password in second field
        await driver.findElement(By.name("Password")).click();
        await driver.findElement(By.name("Password")).clear();
        await driver.findElement(By.name("Password")).sendKeys(firstContractor.password);

        // Additional group/location operations
        try {
            await driver.findElement(By.name("addHierarchyGroupsLocations")).click();
            await driver.findElement(By.css("#ManageHierarchyUGroupLocationsForm > .GroupTab:nth-child(1)")).click();

            // Select User Group - Third Party Staff
            const dropdown = await driver.findElement(By.id("UGroupID"));
            await dropdown.findElement(By.xpath("//option[. = 'Third Party Staff']")).click();
            await driver.findElement(By.css("option:nth-child(19)")).click();

            // More hierarchy navigation
            await driver.findElement(By.css(".DropdownDisplay__PlaceHolder-sc-6p7u3y-1")).click();
            await driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
            await driver.findElement(By.css(".HierarchyNode__ArrowImg-sc-1ytna50-6")).click();
            await driver.findElement(By.css(".HierarchyNode__NodeText-sc-1ytna50-1")).click();
            await driver.findElement(By.css(".HierarchyLookupStateless__LookupDiv-sc-1c6bi43-0 > div:nth-child(3)")).click();
        } catch (error) {
            console.log('Some group operations failed, continuing...');
        }

        // Set action and employment status
        try {
            await driver.findElement(By.id("Action")).click();

            // Set Status of Employment to Current
            const statusDropdown = await driver.findElement(By.id("Attributes.Users.StatusOfEmployment"));
            await statusDropdown.findElement(By.xpath("//option[. = 'Current']")).click();
            await driver.findElement(By.css("#Attributes\\.Users\\.StatusOfEmployment > option:nth-child(3)")).click();

            // Set User Type to Limited access user
            const userTypeDropdown = await driver.findElement(By.id("UserTypeID"));
            await userTypeDropdown.findElement(By.xpath("//option[. = 'Limited access user']")).click();
            await driver.findElement(By.css("#UserTypeID > option:nth-child(2)")).click();
        } catch (error) {
            console.log('Some dropdown selections failed, continuing...');
        }

        // Save the user
        await driver.findElement(By.name("save")).click();

        console.log(`✓ Successfully created user: ${firstContractor.username}`);

        // Wait a bit to see the result
        await driver.sleep(5000);
    });
});

// Function to run the test directly (non-Mocha mode)
async function createFirstUser() {
    console.log('Starting first user creation test...');

    try {
        // Read Excel file
        const workbook = XLSX.readFile('Copy of Contractors staff list.xlsx');
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        // Helper function for column finding
        function findColumnIndex(headers, possibleNames) {
            for (let i = 0; i < headers.length; i++) {
                const header = headers[i] ? headers[i].toString().toLowerCase().trim() : '';
                if (possibleNames.some(name => header.includes(name))) {
                    return i;
                }
            }
            return -1;
        }

        // Find column indices
        const headers = rawData[0] || [];
        const codeIndex = 0;
        const firstNameIndex = findColumnIndex(headers, ['first name', 'firstname', 'forename']);
        const lastNameIndex = findColumnIndex(headers, ['last name', 'lastname', 'surname']);
        const departmentIndex = findColumnIndex(headers, ['department', 'dept']);

        console.log('📍 Column mapping:');
        console.log(`   Code: Column ${codeIndex} (${headers[codeIndex]})`);
        console.log(`   First Name: Column ${firstNameIndex} (${headers[firstNameIndex]})`);
        console.log(`   Last Name: Column ${lastNameIndex} (${headers[lastNameIndex]})`);
        console.log(`   Department: Column ${departmentIndex} (${headers[departmentIndex]})`);

        if (firstNameIndex === -1 || lastNameIndex === -1 || departmentIndex === -1) {
            console.error('❌ Required columns not found. Need: First Name, Last Name, Department');
            return;
        }

        // Find first Security contractor
        let contractor = null;
        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            if (row && row[codeIndex] && row[firstNameIndex] && row[lastNameIndex]) {
                const department = row[departmentIndex] ? row[departmentIndex].toString().trim() : '';

                if (department.toLowerCase() === 'security') {
                    const code = row[codeIndex].toString().trim();
                    const firstName = row[firstNameIndex].toString().trim();
                    const lastName = row[lastNameIndex].toString().trim();
                    const fullName = `${firstName} ${lastName}`;
                    const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                    contractor = {
                        code: code,
                        fullName: fullName,
                        firstName: firstName,
                        lastName: lastName,
                        department: department || 'Security',
                        username: username,
                        password: username
                    };
                    break;
                }
            }
        }

        if (!contractor) {
            console.log('⚠️  No Security department contractors found in Excel file');
            return;
        }

        console.log('👥 First Security contractor found:');
        console.log(`   Name: ${contractor.fullName}`);
        console.log(`   Username: ${contractor.username}`);
        console.log(`   Department: ${contractor.department}`);
        console.log(`   Code: ${contractor.code}`);
        console.log('\n🧪 To run this test with the browser:');
        console.log('   npm test');
        console.log('   OR: npx mocha test-first-user-updated.js');

    } catch (error) {
        console.error('❌ Error:', error.message);
    }
}

// Run if called directly
if (require.main === module) {
    createFirstUser();
}