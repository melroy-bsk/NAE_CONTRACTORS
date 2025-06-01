// Simple Excel Validation Script - No Browser Required
const XLSX = require('xlsx');
const fs = require('fs');

function validateExcelFile(filePath = 'Copy of Contractors staff list.xlsx') {
    console.log('📊 Excel File Validation');
    console.log('='.repeat(40));

    try {
        // Check if file exists
        if (!fs.existsSync(filePath)) {
            console.log(`❌ File not found: ${filePath}`);
            console.log('💡 Make sure the Excel file is in the same directory as this script');
            return false;
        }

        console.log(`✅ File found: ${filePath}`);

        // Read Excel file
        const workbook = XLSX.readFile(filePath);
        console.log(`📑 Sheets in workbook: ${workbook.SheetNames.join(', ')}`);

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const rawData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

        console.log(`📋 Total rows: ${rawData.length} (including header)`);

        if (rawData.length === 0) {
            console.log('❌ Excel file is empty');
            return false;
        }

        // Show header structure
        console.log('\n📌 Column Headers:');
        const headers = rawData[0] || [];
        headers.forEach((header, index) => {
            console.log(`   ${index}: ${header || '(empty)'}`);
        });

        // Find column indices for required fields
        const codeIndex = 0; // Assuming first column is always code
        const firstNameIndex = findColumnIndex(headers, ['first name', 'firstname', 'forename']);
        const lastNameIndex = findColumnIndex(headers, ['last name', 'lastname', 'surname']);
        const departmentIndex = findColumnIndex(headers, ['department', 'dept']);

        // Validate expected columns
        console.log('\n🔍 Validating required columns...');
        console.log(`📍 Column mapping:`);

        const hasCode = codeIndex >= 0 && headers[codeIndex];
        const hasFirstName = firstNameIndex >= 0;
        const hasLastName = lastNameIndex >= 0;
        const hasDept = departmentIndex >= 0;

        console.log(`   Code: ${hasCode ? '✅' : '❌'} Column ${codeIndex} "${headers[codeIndex] || 'missing'}"`);
        console.log(`   First Name: ${hasFirstName ? '✅' : '❌'} Column ${firstNameIndex >= 0 ? firstNameIndex : 'not found'} "${headers[firstNameIndex] || 'missing'}"`);
        console.log(`   Last Name: ${hasLastName ? '✅' : '❌'} Column ${lastNameIndex >= 0 ? lastNameIndex : 'not found'} "${headers[lastNameIndex] || 'missing'}"`);
        console.log(`   Department: ${hasDept ? '✅' : '❌'} Column ${departmentIndex >= 0 ? departmentIndex : 'not found'} "${headers[departmentIndex] || 'missing'}"`);

        // Process data and count Security contractors
        console.log('\n📊 Data Analysis:');
        let totalRows = 0;
        let validRows = 0;
        let securityContractors = 0;
        let sampleUsers = [];

        for (let i = 1; i < rawData.length; i++) {
            const row = rawData[i];
            totalRows++;

            if (row && row[codeIndex] &&
                (firstNameIndex >= 0 ? row[firstNameIndex] : false) &&
                (lastNameIndex >= 0 ? row[lastNameIndex] : false)) {

                validRows++;
                const department = departmentIndex >= 0 && row[departmentIndex] ?
                    row[departmentIndex].toString().trim() : '';

                if (department.toLowerCase() === 'security') {
                    securityContractors++;

                    // Create sample user data
                    const code = row[codeIndex].toString().trim();
                    const firstName = row[firstNameIndex].toString().trim();
                    const lastName = row[lastNameIndex].toString().trim();
                    const fullName = `${firstName} ${lastName}`;
                    const username = `${code}_${firstName}_${lastName}`.toLowerCase().replace(/\s+/g, '_');

                    if (sampleUsers.length < 3) {
                        sampleUsers.push({
                            code: code,
                            fullName: fullName,
                            firstName: firstName,
                            lastName: lastName,
                            department: department,
                            username: username
                        });
                    }
                }
            }
        }

        console.log(`   📈 Total data rows: ${totalRows}`);
        console.log(`   ✅ Valid rows (have code + first name + last name): ${validRows}`);
        console.log(`   🛡️  Security department contractors: ${securityContractors}`);

        if (securityContractors === 0) {
            console.log('\n⚠️  WARNING: No Security department contractors found!');
            console.log('   Make sure the Department column contains "Security" for the contractors you want to create.');
        }

        // Show sample users
        if (sampleUsers.length > 0) {
            console.log('\n👥 Sample users that would be created:');
            sampleUsers.forEach((user, index) => {
                console.log(`   ${index + 1}. ${user.fullName}`);
                console.log(`      Username: ${user.username}`);
                console.log(`      Department: ${user.department}`);
                console.log(`      Code: ${user.code}`);
            });
        }

        // Show any issues
        console.log('\n🔧 Potential Issues:');
        let issueCount = 0;

        if (!hasCode || !hasFirstName || !hasLastName || !hasDept) {
            console.log(`   ❌ Missing required columns. Need: Code, First Name, Last Name, Department`);
            issueCount++;
        }

        if (validRows < totalRows) {
            console.log(`   ⚠️  ${totalRows - validRows} rows missing required data`);
            issueCount++;
        }

        if (securityContractors === 0) {
            console.log(`   ⚠️  No Security department contractors found`);
            issueCount++;
        }

        if (issueCount === 0) {
            console.log(`   ✅ No issues found! Ready to process ${securityContractors} users.`);
        }

        return true;

    } catch (error) {
        console.log(`❌ Error reading Excel file: ${error.message}`);
        return false;
    }
}

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

// Check for different possible file names
function findExcelFile() {
    const possibleNames = [
        'Copy of Contractors staff list.xlsx',
        'Contractors staff list.xlsx',
        'contractors.xlsx',
        'staff.xlsx'
    ];

    for (const name of possibleNames) {
        if (fs.existsSync(name)) {
            return name;
        }
    }

    console.log('🔍 Looking for Excel files in current directory...');
    const files = fs.readdirSync('.').filter(f => f.endsWith('.xlsx') || f.endsWith('.xls'));

    if (files.length > 0) {
        console.log('📁 Found Excel files:');
        files.forEach(f => console.log(`   - ${f}`));
        return files[0];
    }

    return null;
}

// Main execution
if (require.main === module) {
    const args = process.argv.slice(2);
    const fileName = args[0] || findExcelFile();

    if (!fileName) {
        console.log('❌ No Excel file specified or found');
        console.log('Usage: node validate-excel.js [filename.xlsx]');
        process.exit(1);
    }

    validateExcelFile(fileName);
}

module.exports = { validateExcelFile, findExcelFile };