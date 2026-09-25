/**
 * Scans all project files for hardcoded column indices 
 * referencing rows from the "Posts" sheet.
 */
function auditHardcodedColumns() {
  const scriptId = ScriptApp.getScriptId();
  const token = ScriptApp.getOAuthToken();
  const url = "https://script.googleapis.com/v1/projects/" + scriptId + "/content";
  
  const options = {
    headers: { Authorization: "Bearer " + token }
  };
  
  try {
    const response = UrlFetchApp.fetch(url, options);
    const project = JSON.parse(response.getContentText());
    
    Logger.log("--- HARDCODED COLUMN AUDIT RESULTS ---");
    let foundCount = 0;

    project.files.forEach(file => {
      const lines = file.source.split('\n');
      lines.forEach((line, index) => {
        
        // Pattern 1: row[number] - Finding array indices
        const arrayMatch = line.match(/\[\d+\]/);
        
        // Pattern 2: getRange(row, number) - Finding static column ranges
        const rangeMatch = line.match(/getRange\([^,]+,\s*\d+/);

        if (arrayMatch || rangeMatch) {
          // Filter: Only log if it looks like it's processing sheet data
          if (line.includes("row") || line.includes("data") || line.includes("posts")) {
            Logger.log(`FILE: ${file.name} | LINE ${index + 1}: ${line.trim()}`);
            foundCount++;
          }
        }
      });
    });

    if (foundCount === 0) {
      Logger.log("No hardcoded indices found. All systems appear dynamic!");
    } else {
      Logger.log(`Scan Complete. Found ${foundCount} potential hardcoded references.`);
    }

 } catch (e) {
    Logger.log("API Error: Ensure Apps Script API is ON at script.google.com/home/usersettings");
    Logger.log(e.toString());
  }
}