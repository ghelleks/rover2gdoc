/**
 * Google Apps Script to generate an Organization Chart from Google Sheets data
 * 
 * This script reads employee data from a Google Sheet and creates a hierarchical
 * organization chart in a Google Doc format.
 * 
 * Configuration is loaded from config.gs - see config.gs for setup instructions.
 */

/**
 * Main function to generate the organization chart
 */
function generateOrgChart() {
  try {
    // Read data from Google Sheet
    const employeeData = readEmployeeData();
    
    // Process and build hierarchy
    const hierarchy = buildHierarchy(employeeData);
    
    // Create or update Google Doc
    createOrgChartDoc(hierarchy);
    
    console.log('Organization chart generated successfully!');
  } catch (error) {
    console.error('Error generating organization chart:', error);
    throw error;
  }
}

/**
 * Read employee data from Google Sheet
 */
function readEmployeeData() {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  if (data.length < 2) {
    throw new Error('Sheet appears to be empty or has no data rows');
  }
  
  const headers = data[0];
  const employees = [];
  
  // Find column indices
  const columnMap = {};
  headers.forEach((header, index) => {
    columnMap[header] = index;
  });
  
  // Required columns
  const requiredColumns = ['Name', 'Job Title', 'Organization Name', 'Location', 'Email', 'Manager UID', 'User ID'];
  const missingColumns = requiredColumns.filter(col => !(col in columnMap));
  if (missingColumns.length > 0) {
    throw new Error(`Missing required columns: ${missingColumns.join(', ')}`);
  }
  
  // Process each employee row
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    
    const employee = {
      name: row[columnMap['Name']] || '',
      userId: row[columnMap['User ID']] || '',
      jobTitle: row[columnMap['Job Title']] || row[columnMap['Business Card Title']] || '',
      organization: row[columnMap['Organization Name']] || '',
      location: row[columnMap['Location']] || '',
      email: row[columnMap['Email']] || '',
      telephone: row[columnMap['Telephone']] || row[columnMap['Mobile']] || row[columnMap['Home Phone']] || '',
      managerUID: row[columnMap['Manager UID']] || '',
      region: row[columnMap['Region']] || '',
      status: row[columnMap['Status']] || '',
      employeeType: row[columnMap['Employee Type']] || ''
    };
    
    // Only include current employees
    if (employee.status === 'Current Employee' && employee.name.trim() !== '') {
      employees.push(employee);
    }
  }
  
  console.log(`Loaded ${employees.length} current employees`);
  return employees;
}

/**
 * Build hierarchical structure from employee data
 */
function buildHierarchy(employees) {
  // Create lookup maps
  const employeeMap = new Map();
  const managerMap = new Map();
  
  // Index employees by User ID
  employees.forEach(emp => {
    employeeMap.set(emp.userId, emp);
  });
  
  // Build manager-subordinate relationships
  employees.forEach(emp => {
    if (emp.managerUID && emp.managerUID.trim() !== '') {
      if (!managerMap.has(emp.managerUID)) {
        managerMap.set(emp.managerUID, []);
      }
      managerMap.get(emp.managerUID).push(emp);
    }
  });
  
  // Find top-level managers (those without managers or whose managers are not in the dataset)
  const topLevelManagers = employees.filter(emp => {
    return !emp.managerUID || 
           emp.managerUID.trim() === '' || 
           !employeeMap.has(emp.managerUID);
  });
  
  // Build hierarchy tree
  function buildTree(manager) {
    const node = {
      employee: manager,
      subordinates: []
    };
    
    const directReports = managerMap.get(manager.userId) || [];
    directReports.forEach(subordinate => {
      node.subordinates.push(buildTree(subordinate));
    });
    
    // Sort subordinates by job title hierarchy and name
    node.subordinates.sort((a, b) => {
      const titleRankA = getTitleRank(a.employee.jobTitle);
      const titleRankB = getTitleRank(b.employee.jobTitle);
      
      if (titleRankA !== titleRankB) {
        return titleRankA - titleRankB;
      }
      return a.employee.name.localeCompare(b.employee.name);
    });
    
    return node;
  }
  
  // Build trees for all top-level managers
  const hierarchyTrees = topLevelManagers.map(manager => buildTree(manager));
  
  // Sort top-level managers by seniority
  hierarchyTrees.sort((a, b) => {
    const titleRankA = getTitleRank(a.employee.jobTitle);
    const titleRankB = getTitleRank(b.employee.jobTitle);
    
    if (titleRankA !== titleRankB) {
      return titleRankA - titleRankB;
    }
    return a.employee.name.localeCompare(b.employee.name);
  });
  
  return hierarchyTrees;
}

/**
 * Get numerical rank for job titles (lower number = higher rank)
 */
function getTitleRank(title) {
  const titleLower = title.toLowerCase();
  
  if (titleLower.includes('senior director')) return 1;
  if (titleLower.includes('director')) return 2;
  if (titleLower.includes('senior manager')) return 3;
  if (titleLower.includes('manager')) return 4;
  if (titleLower.includes('senior principal')) return 5;
  if (titleLower.includes('principal')) return 6;
  if (titleLower.includes('senior')) return 7;
  if (titleLower.includes('lead')) return 8;
  if (titleLower.includes('intern')) return 10;
  if (titleLower.includes('collaborative partner')) return 9;
  
  return 8; // Default rank for other titles
}

/**
 * Create or update Google Doc with organization chart
 */
function createOrgChartDoc(hierarchy) {
  let doc;
  
  if (DOC_ID && DOC_ID !== 'YOUR_GOOGLE_DOC_ID_HERE') {
    // Open existing document
    try {
      doc = DocumentApp.openById(DOC_ID);
      doc.getBody().clear();
    } catch (error) {
      console.log('Could not open existing doc, creating new one');
      doc = DocumentApp.create('Organization Chart - Generated ' + new Date().toLocaleDateString());
    }
  } else {
    // Create new document
    doc = DocumentApp.create('Organization Chart - Generated ' + new Date().toLocaleDateString());
  }
  
  const body = doc.getBody();
  
  // Add title
  const title = body.appendParagraph('Organization Chart');
  title.setHeading(DocumentApp.ParagraphHeading.HEADING1);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  
  // Add generation timestamp
  const timestamp = body.appendParagraph(`Generated on: ${new Date().toLocaleString()}`);
  timestamp.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  timestamp.editAsText().setFontSize(10).setItalic(true);
  
  body.appendParagraph(''); // Empty line
  
  // Add organization hierarchy
  hierarchy.forEach(tree => {
    renderHierarchyTree(body, tree, 0);
    body.appendParagraph(''); // Empty line between top-level hierarchies
  });
  
  // Add summary statistics
  addSummaryStatistics(body, hierarchy);
  
  console.log(`Organization chart created: ${doc.getUrl()}`);
  return doc.getUrl();
}

/**
 * Render hierarchy tree in the document using unordered lists
 */
function renderHierarchyTree(body, node, level) {
  const employee = node.employee;
  const isManager = node.subordinates.length > 0;
  
  // Build the complete employee information on one line
  const details = [];
  if (employee.jobTitle) details.push(`Title: ${employee.jobTitle}`);
  if (employee.organization) details.push(`Organization: ${employee.organization}`);
  if (employee.location) details.push(`Location: ${employee.location}`);
  if (employee.email) details.push(`Email: ${employee.email}`);
  if (employee.telephone) details.push(`Phone: ${employee.telephone}`);
  
  // Create the full text with name and details
  const fullText = details.length > 0 
    ? `${employee.name} - ${details.join(' | ')}`
    : employee.name;
    
  // Create employee entry as list item
  const employeeItem = body.appendListItem(fullText);
  
  // Set list nesting level and bullet type
  employeeItem.setNestingLevel(level);
  employeeItem.setGlyphType(DocumentApp.GlyphType.BULLET);
  
  // Style the text - make name bold, keep details regular
  const textRange = employeeItem.editAsText();
  
  // Make the name portion bold (from start to the dash or end of name)
  const nameEndIndex = employee.name.length;
  textRange.setBold(0, nameEndIndex - 1, true);
  
  // Set font size based on level
  if (level === 0) {
    textRange.setFontSize(14);
  } else if (level === 1) {
    textRange.setFontSize(12);
  } else {
    textRange.setFontSize(11);
  }
  
  // Recursively render subordinates
  node.subordinates.forEach(subordinate => {
    renderHierarchyTree(body, subordinate, level + 1);
  });
}

/**
 * Add summary statistics to the document
 */
function addSummaryStatistics(body, hierarchy) {
  body.appendPageBreak();
  
  const summaryTitle = body.appendParagraph('Summary Statistics');
  summaryTitle.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  
  // Count employees by level and organization
  const stats = {
    totalEmployees: 0,
    byOrganization: new Map(),
    byLocation: new Map(),
    byTitle: new Map()
  };
  
  function collectStats(node) {
    const emp = node.employee;
    stats.totalEmployees++;
    
    // By organization
    const org = emp.organization || 'Unknown';
    stats.byOrganization.set(org, (stats.byOrganization.get(org) || 0) + 1);
    
    // By location
    const loc = emp.location || 'Unknown';
    stats.byLocation.set(loc, (stats.byLocation.get(loc) || 0) + 1);
    
    // By title level
    const titleLevel = getTitleLevel(emp.jobTitle);
    stats.byTitle.set(titleLevel, (stats.byTitle.get(titleLevel) || 0) + 1);
    
    // Process subordinates
    node.subordinates.forEach(collectStats);
  }
  
  hierarchy.forEach(collectStats);
  
  // Display statistics
  body.appendParagraph(`Total Employees: ${stats.totalEmployees}`);
  
  body.appendParagraph('\nBy Organization:');
  Array.from(stats.byOrganization.entries())
    .sort((a, b) => b[1] - a[1])
    .forEach(([org, count]) => {
      body.appendParagraph(`  ${org}: ${count}`);
    });
  
  body.appendParagraph('\nBy Location:');
  Array.from(stats.byLocation.entries())
    .sort((a, b) => b[1] - a[1])
    .forEach(([loc, count]) => {
      body.appendParagraph(`  ${loc}: ${count}`);
    });
  
  body.appendParagraph('\nBy Title Level:');
  Array.from(stats.byTitle.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .forEach(([title, count]) => {
      body.appendParagraph(`  ${title}: ${count}`);
    });
}

/**
 * Get title level for statistics
 */
function getTitleLevel(title) {
  const titleLower = title.toLowerCase();
  
  if (titleLower.includes('director')) return 'Director Level';
  if (titleLower.includes('manager')) return 'Manager Level';
  if (titleLower.includes('principal')) return 'Principal Level';
  if (titleLower.includes('senior')) return 'Senior Level';
  if (titleLower.includes('intern')) return 'Intern Level';
  if (titleLower.includes('collaborative partner')) return 'Partner Level';
  
  return 'Individual Contributor';
}

/**
 * Helper function to create a new document and get shareable link
 */
function createNewOrgChart() {
  const url = generateOrgChart();
  console.log('New organization chart created at:', url);
  return url;
}

/**
 * Test function to validate data reading
 */
function testDataReading() {
  try {
    const data = readEmployeeData();
    console.log('Sample employee data:');
    console.log(data.slice(0, 3)); // Show first 3 employees
    console.log(`Total employees loaded: ${data.length}`);
  } catch (error) {
    console.error('Error reading data:', error);
  }
} 