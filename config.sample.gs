/**
 * Sample Configuration file for Organization Chart Generator
 * 
 * SETUP INSTRUCTIONS:
 * 1. Copy this file and rename it to "config.gs" 
 * 2. Update the IDs below with your actual Google Sheet and Document IDs
 * 3. The real config.gs file will be ignored by git (as specified in .gitignore)
 * 
 * DO NOT commit the real config.gs file with actual IDs to source control!
 */

// Replace with your actual Google Sheet ID
// To find your Sheet ID: Open your Google Sheet and look at the URL
// https://docs.google.com/spreadsheets/d/[SHEET_ID]/edit
// Copy the long string between /d/ and /edit
const SHEET_ID = 'YOUR_GOOGLE_SHEET_ID_HERE';

// Replace with your target Google Doc ID (optional)
// To find your Doc ID: Create or open a Google Doc and look at the URL  
// https://docs.google.com/document/d/[DOC_ID]/edit
// Copy the long string between /d/ and /edit
// Leave as default to create a new document each time
const DOC_ID = 'YOUR_GOOGLE_DOC_ID_HERE'; 