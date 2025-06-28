/**
 * Setup Helper for Organization Chart Generator
 * 
 * This file contains helper functions to make setup and testing easier.
 * Run these functions to validate your setup before generating the full org chart.
 */

/**
 * Interactive setup function - prompts user for Sheet ID
 * Run this function first to set up your configuration
 */
function interactiveSetup() {
  const ui = SpreadsheetApp.getUi();
  
  // Get Sheet ID from user
  const sheetIdResponse = ui.prompt(
    'Setup Organization Chart Generator',
    'Please enter your Google Sheet ID (the long string from the sheet URL):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (sheetIdResponse.getSelectedButton() === ui.Button.OK) {
    const sheetId = sheetIdResponse.getResponseText().trim();
    
    if (sheetId) {
      // Test the sheet ID
      try {
        const sheet = SpreadsheetApp.openById(sheetId);
        const sheetName = sheet.getName();
        
        ui.alert(
          'Success!',
          `Connected to sheet: "${sheetName}"\n\nNext steps:\n1. Update the SHEET_ID constant in your main script\n2. Run testDataReading() to validate your data\n3. Run generateOrgChart() to create your org chart`,
          ui.ButtonSet.OK
        );
        
        console.log(`Sheet ID validated: ${sheetId}`);
        console.log(`Sheet name: ${sheetName}`);
        console.log('Update your SHEET_ID constant with this value');
        
      } catch (error) {
        ui.alert(
          'Error',
          `Could not access sheet with ID: ${sheetId}\n\nPlease check:\n1. The Sheet ID is correct\n2. You have access to the sheet\n3. The sheet exists`,
          ui.ButtonSet.OK
        );
        console.error('Sheet access error:', error);
      }
    }
  }
}

/**
 * Validate sheet structure and show sample data
 */
function validateSheetStructure() {
  if (SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
    console.log('Please update SHEET_ID first. Run interactiveSetup() or manually update the constant.');
    return;
  }
  
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    const data = sheet.getDataRange().getValues();
    
    if (data.length < 2) {
      console.log('⚠️  Sheet appears to have no data rows (only headers or empty)');
      return;
    }
    
    const headers = data[0];
    console.log('📋 Sheet Structure Analysis:');
    console.log(`   Total rows: ${data.length} (including header)`);
    console.log(`   Total columns: ${headers.length}`);
    console.log('');
    
    // Check for required columns
    const requiredColumns = ['Name', 'Job Title', 'Organization Name', 'Location', 'Email', 'Manager UID', 'User ID'];
    const foundColumns = {};
    
    console.log('🔍 Column Analysis:');
    headers.forEach((header, index) => {
      const isRequired = requiredColumns.includes(header);
      console.log(`   ${index + 1}. "${header}" ${isRequired ? '✅ (required)' : ''}`);
      foundColumns[header] = index;
    });
    
    console.log('');
    
    // Check for missing required columns
    const missingColumns = requiredColumns.filter(col => !(col in foundColumns));
    if (missingColumns.length > 0) {
      console.log('❌ Missing Required Columns:');
      missingColumns.forEach(col => {
        console.log(`   - ${col}`);
      });
      console.log('');
    } else {
      console.log('✅ All required columns found!');
      console.log('');
    }
    
    // Show sample data
    console.log('📊 Sample Data (first 3 rows):');
    for (let i = 1; i <= Math.min(3, data.length - 1); i++) {
      const row = data[i];
      console.log(`   Row ${i}:`);
      console.log(`     Name: ${row[foundColumns['Name']] || 'N/A'}`);
      console.log(`     Job Title: ${row[foundColumns['Job Title']] || 'N/A'}`);
      console.log(`     Organization: ${row[foundColumns['Organization Name']] || 'N/A'}`);
      console.log(`     Manager UID: ${row[foundColumns['Manager UID']] || 'N/A'}`);
      console.log(`     Status: ${row[foundColumns['Status']] || 'N/A'}`);
      console.log('');
    }
    
    // Count current employees
    if (foundColumns['Status']) {
      const currentEmployees = data.slice(1).filter(row => 
        row[foundColumns['Status']] === 'Current Employee'
      ).length;
      console.log(`👥 Current Employees: ${currentEmployees} out of ${data.length - 1} total records`);
    }
    
    if (missingColumns.length === 0) {
      console.log('');
      console.log('🎉 Your sheet structure looks good! You can now run generateOrgChart()');
    }
    
  } catch (error) {
    console.error('❌ Error validating sheet:', error);
    console.log('Please check your SHEET_ID and ensure you have access to the sheet.');
  }
}

/**
 * Quick test to check manager-employee relationships
 */
function validateHierarchy() {
  if (SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
    console.log('Please update SHEET_ID first.');
    return;
  }
  
  try {
    const employees = readEmployeeData();
    console.log(`📈 Hierarchy Analysis for ${employees.length} employees:`);
    console.log('');
    
    // Create user ID map
    const userIds = new Set(employees.map(emp => emp.userId));
    
    // Check manager relationships
    let validManagers = 0;
    let invalidManagers = 0;
    let topLevel = 0;
    
    employees.forEach(emp => {
      if (!emp.managerUID || emp.managerUID.trim() === '') {
        topLevel++;
      } else if (userIds.has(emp.managerUID)) {
        validManagers++;
      } else {
        invalidManagers++;
        console.log(`⚠️  ${emp.name} has invalid Manager UID: ${emp.managerUID}`);
      }
    });
    
    console.log(`👑 Top-level employees (no manager): ${topLevel}`);
    console.log(`✅ Valid manager relationships: ${validManagers}`);
    console.log(`❌ Invalid manager relationships: ${invalidManagers}`);
    console.log('');
    
    if (invalidManagers === 0) {
      console.log('🎉 All manager relationships are valid!');
    } else {
      console.log('⚠️  Some manager relationships need fixing. Check the warnings above.');
    }
    
    // Show top-level managers
    const topLevelManagers = employees.filter(emp => 
      !emp.managerUID || emp.managerUID.trim() === '' || !userIds.has(emp.managerUID)
    );
    
    console.log('');
    console.log('👑 Top-Level Managers:');
    topLevelManagers.forEach(manager => {
      console.log(`   - ${manager.name} (${manager.jobTitle})`);
    });
    
  } catch (error) {
    console.error('❌ Error validating hierarchy:', error);
  }
}

/**
 * Create a test document to verify formatting
 */
function createTestDocument() {
  try {
    // Create a small test hierarchy
    const testDoc = DocumentApp.create('Organization Chart - Test Document');
    const body = testDoc.getBody();
    
    // Add title
    const title = body.appendParagraph('Organization Chart - Test');
    title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Add sample hierarchy
    body.appendParagraph('');
    body.appendParagraph('Sample Organization Structure:');
    body.appendParagraph('');
    
    // Sample CEO (name in bold, all info on one line)
    const ceoText = 'John Doe - Title: Chief Executive Officer | Organization: Executive | Location: New York, NY | Email: john.doe@company.com';
    const ceo = body.appendListItem(ceoText);
    ceo.setNestingLevel(0);
    ceo.setGlyphType(DocumentApp.GlyphType.BULLET);
    ceo.editAsText().setBold(0, 7, true).setFontSize(14); // Bold "John Doe"
    
    // Sample Director (name in bold, all info on one line)
    const directorText = 'Jane Smith - Title: Director of Engineering | Organization: Engineering | Location: Boston, MA | Email: jane.smith@company.com';
    const director = body.appendListItem(directorText);
    director.setNestingLevel(1);
    director.setGlyphType(DocumentApp.GlyphType.BULLET);
    director.editAsText().setBold(0, 9, true).setFontSize(12); // Bold "Jane Smith"
    
    // Sample Individual Contributor (name in bold, all info on one line)
    const employeeText = 'Bob Johnson - Title: Senior Software Engineer | Organization: Engineering | Location: Boston, MA | Email: bob.johnson@company.com';
    const employee = body.appendListItem(employeeText);
    employee.setNestingLevel(2);
    employee.setGlyphType(DocumentApp.GlyphType.BULLET);
    employee.editAsText().setBold(0, 10, true).setFontSize(11); // Bold "Bob Johnson"
    
    console.log(`✅ Test document created: ${testDoc.getUrl()}`);
    console.log('Review the formatting and styling, then run generateOrgChart() for your actual data.');
    
    return testDoc.getUrl();
    
  } catch (error) {
    console.error('❌ Error creating test document:', error);
  }
}

/**
 * Complete setup wizard - runs all validation steps
 */
function runSetupWizard() {
  console.log('🚀 Organization Chart Generator - Setup Wizard');
  console.log('===============================================');
  console.log('');
  
  // Step 1: Validate configuration
  console.log('Step 1: Validating Configuration...');
  if (SHEET_ID === 'YOUR_GOOGLE_SHEET_ID_HERE') {
    console.log('❌ SHEET_ID not configured. Please run interactiveSetup() first.');
    return;
  }
  console.log('✅ SHEET_ID configured');
  console.log('');
  
  // Step 2: Validate sheet structure
  console.log('Step 2: Validating Sheet Structure...');
  validateSheetStructure();
  console.log('');
  
  // Step 3: Validate hierarchy
  console.log('Step 3: Validating Hierarchy Relationships...');
  validateHierarchy();
  console.log('');
  
  // Step 4: Create test document
  console.log('Step 4: Creating Test Document...');
  createTestDocument();
  console.log('');
  
  console.log('🎉 Setup wizard complete!');
  console.log('Next step: Run generateOrgChart() to create your organization chart.');
} 