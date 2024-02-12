const fs = require('fs');

// File path for the key-value store
const storeFilePath = 'C:\\Users\\PKDev02\\file-api\\keyValueStore.json';

// Function to save docId and EmailID to the file-based key-value store
function saveAssociationToFile(docId, emailID) {
    // Read existing data from the file, or initialize an empty object if the file doesn't exist
    let data = {};
    try {
        data = JSON.parse(fs.readFileSync(storeFilePath));
    } catch (error) {
        // File doesn't exist or is empty, create an empty object
    }

    // Overwrite the existing emailID with the new one
    data[docId] = emailID;

    // Write the updated data back to the file with each key-value pair on a new line
    fs.writeFileSync(storeFilePath, JSON.stringify(data, null, 2));
}

// Function to query docId and get EmailID(s) from the file-based key-value store
function getEmailIDFromFile(docId) {
    // Read data from the file, or return null if the file doesn't exist
    try {
        const data = JSON.parse(fs.readFileSync(storeFilePath));
        return data[docId] || null;
    } catch (error) {
        return null;
    }
}

// Create the JSON file if it doesn't exist
if (!fs.existsSync(storeFilePath)) {
    fs.writeFileSync(storeFilePath, '{}');
}

// Export functions
module.exports = {
    saveAssociationToFile,
    getEmailIDFromFile
};

// Example usage
// saveAssociationToFile('doc123', 'email1@example.com');

// const emailID = getEmailIDFromFile('doc123');
// console.log(emailID); // Output: 'email2@example.com' (Updated email ID)
