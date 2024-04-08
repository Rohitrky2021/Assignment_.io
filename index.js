// Import required modules
const fs = require('fs'); // File system module for interacting with files
const XLSX = require('xlsx'); // XLSX module for Excel file manipulation

// Function to convert nested JSON data to Excel format
/**
 * Converts a nested JSON object to an Excel file.
 * @param {Object} json - The nested JSON object to convert.
 * @param {string} outputFileName - The name of the output Excel file.
 */
function convertNestedJSONToExcel(json, outputFileName) {

  // Create a new workbook
  const workbook = XLSX.utils.book_new();


// =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

try{
  // Sheet 1: Basic Information with Test Columns
  const sheet1Data = [
    {
      "Name": json.name, // Name of the store
      "Location": json.location, // Location of the store
      "Is Open": json.isOpen, // Whether the store is open or not
      "Number of Sections": json.numberOfSections, // Number of sections in the store
      "Contact": json.contact, // Contact information (null in this example)
      "Popular Genres": json.popularGenres.join(', '), // Popular genres in the store, joined into a string
    }
  ];

  const sheet1 = XLSX.utils.json_to_sheet(sheet1Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet1, 'Sheet1');


  // =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  // Sheet 2: Empty Sheet
  const sheet2 = XLSX.utils.aoa_to_sheet([]); // Empty sheet
  XLSX.utils.book_append_sheet(workbook, sheet2, 'Sheet2');


  // =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  // Sheet 3: Test Object
  const sheet3Data = [];
  if (json.test) { // If test data exists in the JSON
    for (const key in json.test) {
      const row = { "Test Key": key, "Test Value": json.test[key] }; // Extract test data key-value pairs
      sheet3Data.push(row); // Add to sheet3Data array
    }
  }
  const sheet3 = XLSX.utils.json_to_sheet(sheet3Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet3, 'Sheet3');


  // =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  // Sheet 4: Section 1 with Title, Author, Price, Is Available Columns
  const sheet4Data = [];
  if (json.sections && json.sections[0] && json.sections[0].books) {
    json.sections[0].books.forEach((book, index) => { // Iterate through books in Section 1
      const row = { "Title": book.title, "Author": book.author, "Price": book.price, "Is Available": book.isAvailable, "Section": "Section 1" }; // Extract book details and section name
      sheet4Data.push(row); // Add to sheet4Data array
    });
  }
  const sheet4 = XLSX.utils.json_to_sheet(sheet4Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet4, 'Sheet4');


  // =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

  // Sheet 5: Section 2 with Title, Author, Price, Is Available Columns
  const sheet5Data = [];
  if (json.sections && json.sections[1] && json.sections[1].books) {
    json.sections[1].books.forEach((book, index) => { // Iterate through books in Section 2
      const row = { "Title": book.title, "Author": book.author, "Price": book.price, "Is Available": book.isAvailable, "Section": "Section 2" }; // Extract book details and section name
      sheet5Data.push(row); // Add to sheet5Data array
    });
  }
  const sheet5 = XLSX.utils.json_to_sheet(sheet5Data, { skipHeader: true });
  XLSX.utils.book_append_sheet(workbook, sheet5, 'Sheet5');

  // Write the workbook to an Excel file
  XLSX.writeFile(workbook, outputFileName); // Save the workbook to the specified file

  console.log(`Conversion successful. Check ${outputFileName}`); // Print success message
} catch (error) {
  console.error('Error occurred while converting JSON to Excel:', error.message);
}
}


// =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>


// Example nested JSON data
const nestedJSON = {
  "name": "The Reading Nook", // Store name
  "location": "123 Book St, Bibliopolis", // Store location
  "isOpen": true, // Whether the store is open
  "numberOfSections": 2, // Number of sections in the store
  "contact": null, // Contact information (null in this example)
  "popularGenres": ["Fiction", "Mystery", "Sci-Fi", "Non-Fiction"], // Popular genres in the store
  "test": { // Test data
    "test1": "Test 1",
    "test2": {
      "test3": "Test 3"
    }
  },
  "sections": [ // Store sections
    {
      "sectionName": "Section 1", // Name of Section 1
      "books": [ // Books in Section 1
        {
          "title": "Journey to the Unknown",
          "author": "Alice Wonder",
          "price": 12.99,
          "isAvailable": true
        },
        {
          "title": "Mystery of the Ancient Map",
          "author": "Clive Cussler",
          "price": 15.50,
          "isAvailable": false
        }
      ]
    },
    {
      "sectionName": "Section 2", // Name of Section 2
      "books": [ // Books in Section 2
        {
          "title": "The Reality of Myths",
          "author": "Helen Troy",
          "price": 18.25,
          "isAvailable": true
        }
      ]
    }
  ]
};


// =============================================================================>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>

// Specify the output file name
const outputFileName = 'output.xlsx';


// Convert nested JSON to Excel
convertNestedJSONToExcel(nestedJSON, outputFileName); // Call the function to convert JSON to Excel

// console.log(`Conversion successful. Check ${outputFileName}`); // Print success message
