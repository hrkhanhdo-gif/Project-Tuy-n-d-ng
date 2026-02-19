// CODE.GS - SERVER SIDE SCRIPT

// USER CONFIGURATION
const SPREADSHEET_ID = '1znYwNYb0zXMucDuoq8cD2WiyWbTLjnmW3vEP-JYfmoA';
const CV_FOLDER_ID = '1xhqxLSXYQLQ0qSYhnQjAXkwI-EbUtv_M';
const CANDIDATE_SHEET_NAME = 'DATA ỨNG VIÊN';
const SETTINGS_SHEET_NAME = 'CẤU HÌNH HỆ THỐNG';

// 1. SETUP & ROUTING
function doGet(e) {
  // Check if DB exists to decide initial UI state
  const isDbSetup = !!SPREADSHEET_ID;
  const t = HtmlService.createTemplateFromFile('Index');
  t.isDbSetup = isDbSetup;
  return t.evaluate()
      .setTitle('Recruitment ATS System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

// 2. AUTHENTICATION
function apiLogin(username, password) {
  try {
      // Simple check for config existence
      if (!SPREADSHEET_ID) {
        return { success: false, message: 'Chưa cấu hình Spreadsheet ID.' };
      }

      // Normalize inputs
      const cleanUser = username ? username.toString().trim() : '';
      const cleanPass = password ? password.toString().trim() : '';

      const users = apiGetTableData('Users');
      
      // 1. Check DB Users (Case insensitive for username)
      const user = users.find(u => {
          return u.Username && u.Username.toString().toLowerCase() === cleanUser.toLowerCase() && u.Password == cleanPass;
      });
      
      if (user) {
        return { success: true, user: { username: user.Username, role: user.Role, name: user.Full_Name } };
      }
      
      // 2. Fallback Default Admin (if not found in DB)
      // Allow 'admin' (case insensitive) with specific passwords
      if (cleanUser.toLowerCase() === 'admin' && (cleanPass === 'admin' || cleanPass === '123456')) {
           return { success: true, message: 'Login via Default Admin', user: { username: 'admin', role: 'Admin', name: 'System Admin' } };
      }

      return { success: false, message: 'Sai tên đăng nhập hoặc mật khẩu. (Input: ' + cleanUser + ')' };
  } catch (e) {
      return { success: false, message: 'Server Error: ' + e.toString() };
  }
}

// 3. API: GET INITIAL DATA
function apiGetInitialData() {
  Logger.log('=== apiGetInitialData Called ===');
  Logger.log('SPREADSHEET_ID: ' + SPREADSHEET_ID);
  Logger.log('CANDIDATE_SHEET_NAME: ' + CANDIDATE_SHEET_NAME);
  
  if (!SPREADSHEET_ID) {
    Logger.log('ERROR: No SPREADSHEET_ID configured');
    return { candidates: [], stages: [] };
  }

  try {
    const candidates = apiGetTableData(CANDIDATE_SHEET_NAME);
    Logger.log('Candidates loaded: ' + candidates.length);
    
    // Clean candidates data - convert dates to strings, ensure all fields are serializable
    const cleanCandidates = candidates.map(function(c) {
      return {
        ID: String(c.ID || ''),
        Name: String(c.Name || ''),
        Phone: String(c.Phone || ''),
        Email: String(c.Email || ''),
        Position: String(c.Position || ''),
        Source: String(c.Source || ''),
        Stage: String(c.Stage || ''),
        Status: String(c.Status || ''),
        CV_Link: String(c.CV_Link || ''),
        Applied_Date: c.Applied_Date ? (c.Applied_Date instanceof Date ? c.Applied_Date.toISOString() : String(c.Applied_Date)) : '',
        Department: String(c.Department || ''),
        Contact_Status: String(c.Contact_Status || ''),
        Recruiter: String(c.Recruiter || ''),
        Experience: String(c.Experience || ''),
        Education: String(c.Education || ''),
        Expected_Salary: String(c.Expected_Salary || ''),
        Notes: String(c.Notes || '')
      };
    });
    
    // Stages
    let stages = apiGetTableData('Stages');
    Logger.log('Stages loaded: ' + stages.length);
    
    if(stages.length === 0) {
        Logger.log('Using default stages');
        stages = [
             {ID: 'S1', Stage_Name: 'Apply', Order: 1, Color: '#0d6efd'},
             {ID: 'S2', Stage_Name: 'Interview', Order: 2, Color: '#fd7e14'},
             {ID: 'S3', Stage_Name: 'Offer', Order: 3, Color: '#198754'},
             {ID: 'S4', Stage_Name: 'Rejected', Order: 4, Color: '#dc3545'}
        ];
    }
    
    const result = {
      candidates: cleanCandidates,
      stages: stages
    };
    
    Logger.log('Returning data - Candidates: ' + cleanCandidates.length + ', Stages: ' + stages.length);
    return result;
    
  } catch(e) {
    Logger.log('ERROR in apiGetInitialData: ' + e.toString());
    return { candidates: [], stages: [] };
  }
}


// 4. DATABASE HELPERS
function getSheetByName(name) {
  if (!SPREADSHEET_ID) return null;
  try {
    return SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(name);
  } catch (e) {
    return null;
  }
}

function apiGetTableData(sheetName) {
  Logger.log('--- apiGetTableData called for: ' + sheetName);
  const sheet = getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('ERROR: Sheet not found: ' + sheetName);
    return [];
  }
  
  const startRow = 2;
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  Logger.log('Sheet found. lastRow: ' + lastRow + ', lastCol: ' + lastCol);
  
  if (lastRow < startRow) {
    Logger.log('No data rows (lastRow < 2)');
    return [];
  }
  
  // Get Headers
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  Logger.log('Headers: ' + JSON.stringify(headers));
  
  // Get Data
  const data = sheet.getRange(startRow, 1, lastRow - 1, lastCol).getValues();
  Logger.log('Data rows retrieved: ' + data.length);
  
  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      // Clean header slightly (trim)
      const key = header.toString().trim();
      if(key) obj[key] = row[index];
    });
    return obj;
  });
}

// 5. INITIAL SETUP (Legacy support, mostly manual now)
function apiSetupDatabase() {
    // User already has DB, so we just return success to skip UI prompt
    return { success: true, url: 'https://docs.google.com/spreadsheets/d/' + SPREADSHEET_ID };
}

// 6. API: CANDIDATE MANAGEMENT
function apiCreateCandidate(formObject, fileData) {
  try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Không tìm thấy Sheet: ' + CANDIDATE_SHEET_NAME };

    const newId = 'C' + new Date().getTime(); 
    const appliedDate = new Date().toISOString().slice(0, 10); 
    
    // 1. Handle File Upload
    let cvLink = '';
    if(fileData && fileData.data && fileData.name) {
        try {
            const folder = DriveApp.getFolderById(CV_FOLDER_ID);
            const contentType = fileData.type || 'application/pdf';
            const blob = Utilities.newBlob(Utilities.base64Decode(fileData.data), contentType, fileData.name);
            
            // Rename to Phone Number if exists
            if(formObject.phone) {
                // Get extension
                const ext = fileData.name.split('.').pop();
                blob.setName(formObject.phone + '.' + ext);
            }
            
            const file = folder.createFile(blob);
            file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
            cvLink = file.getUrl();
        } catch(err) {
            Logger.log('Upload Error: ' + err);
            // Continue without file if upload fails, but log it
            cvLink = 'Error Uploading: ' + err.toString(); 
        }
    }

    // 2. Append to Sheet
    const lastCol = sheet.getLastColumn();
    let headers = [];
    
    // Handle empty sheet case
    if (lastCol === 0) {
        // Init default headers if sheet is empty
        headers = ['ID', 'Name', 'Phone', 'Email', 'Position', 'Source', 'Stage', 'CV_Link', 'Applied_Date', 'Status'];
        sheet.appendRow(headers);
    } else {
        headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    const headerMap = {};
    headers.forEach((h, i) => headerMap[h.toString().trim()] = i);
    
    const row = new Array(headers.length).fill('');
    
    // Mapping: Form Field -> Header Name (approximate)
    const map = {
        'ID': newId,
        'Name': formObject.name,
        'Phone': formObject.phone,
        'Email': formObject.email,
        'Position': formObject.position,
        'Applied_Date': appliedDate,
        'Status': 'Apply',
        'Source': 'Website',
        'CV_Link': cvLink,
        'Stage': 'Apply',
        'Họ và Tên': formObject.name,
        'Số điện thoại': formObject.phone,
        'Vị trí': formObject.position,
        'Nguồn': 'Website',
        'Link CV': cvLink,
        'Ngày ứng tuyển': appliedDate,
        'Trạng thái': 'Apply'
    };
    
    // Helper to find index case-insensitively
    function setCol(key, val) {
        // Try exact match first
        if (headerMap.hasOwnProperty(key)) {
            row[headerMap[key]] = val;
            return;
        }
        // Try fuzzy match
        for(let h in headerMap) {
            if(h.toLowerCase().includes(key.toLowerCase())) {
                row[headerMap[h]] = val;
                return;
            }
        }
    }
    
    Object.keys(map).forEach(k => setCol(k, map[k]));

    sheet.appendRow(row);
    
    return { success: true, message: 'Thêm hồ sơ thành công!', data: apiGetInitialData() };
  } catch (e) {
    return { success: false, message: 'Lỗi server: ' + e.toString() };
  }
}

function apiUpdateCandidateStatus(candidateId, newStatus) {
   try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) return { success: false, message: 'Database not found' };
    
    // More robust header search for update as well
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let statusIndex = -1;
    let idIndex = -1;
    
    headers.forEach((h, i) => {
        const lower = h.toString().toLowerCase();
        if(lower === 'id') idIndex = i;
        if(lower.includes('status') || lower === 'trạng thái') statusIndex = i;
    });
    
    if (idIndex === -1) return { success: false, message: 'Cannot find ID column' };
    
    for (let i = 1; i < data.length; i++) {
        if (data[i][idIndex] == candidateId) {
             if(statusIndex > -1) sheet.getRange(i + 1, statusIndex + 1).setValue(newStatus);
            // Also update Stage if exists? (Optional, kept simple for now)
            return { success: true };
        }
    }
    return { success: false, message: 'Candidate ID not found' };
   } catch (e) {
       return { success: false, message: e.toString() };
   }
}

// UPDATE FULL CANDIDATE DETAILS
function apiUpdateCandidate(candidateData) {
  Logger.log('=== UPDATE CANDIDATE (FULL) ===');
  Logger.log('Data: ' + JSON.stringify(candidateData));
  
  try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    let headers = data[0];
    const idColIndex = headers.findIndex(h => h.toString().toLowerCase() === 'id');
    
    if (idColIndex === -1) {
      return { success: false, message: 'ID column not found' };
    }
    
    // AUTO-CREATE MISSING COLUMNS
    const missingColumns = [];
    Object.keys(candidateData).forEach(key => {
      if (key !== 'ID' && !headers.includes(key)) {
        missingColumns.push(key);
      }
    });
    
    if (missingColumns.length > 0) {
      Logger.log('Creating missing columns: ' + missingColumns.join(', '));
      
      // Add new columns to header row
      const lastCol = sheet.getLastColumn();
      missingColumns.forEach((colName, index) => {
        sheet.getRange(1, lastCol + index + 1).setValue(colName);
        headers.push(colName);
      });
      
      Logger.log('Added ' + missingColumns.length + ' new columns');
    }
    
    // Find row with matching ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] == candidateData.ID) {
        Logger.log('Found candidate at row ' + (i + 1));
        
        // Update each field
        Object.keys(candidateData).forEach(key => {
          if (key !== 'ID') {
            const colIndex = headers.indexOf(key);
            if (colIndex !== -1) {
              const value = candidateData[key] || '';
              sheet.getRange(i + 1, colIndex + 1).setValue(value);
              Logger.log('Updated ' + key + ' = ' + value);
            }
          }
        });
        
        Logger.log('Successfully updated candidate at row ' + (i + 1));
        
        // Return updated data
        return {
          success: true,
          message: 'Cập nhật thành công!',
          data: apiGetInitialData() // Refresh all data
        };
      }
    }
    
    return { success: false, message: 'Candidate ID not found: ' + candidateData.ID };
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// DELETE CANDIDATE
function apiDeleteCandidate(candidateId) {
  Logger.log('=== DELETE CANDIDATE ===');
  Logger.log('ID: ' + candidateId);
  
  try {
    const sheet = getSheetByName(CANDIDATE_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const idColIndex = headers.findIndex(h => h.toString().toLowerCase() === 'id');
    
    if (idColIndex === -1) {
      return { success: false, message: 'ID column not found' };
    }
    
    // Find and delete row with matching ID
    for (let i = 1; i < data.length; i++) {
      if (data[i][idColIndex] == candidateId) {
        Logger.log('Found candidate at row ' + (i + 1) + ', deleting...');
        sheet.deleteRow(i + 1);
        Logger.log('Successfully deleted candidate');
        
        return {
          success: true,
          message: 'Đã xóa ứng viên thành công!'
        };
      }
    }
    
    return { success: false, message: 'Candidate ID not found: ' + candidateId };
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// 7. API: JOB MANAGEMENT
function apiGetJobs() {
    if (!SPREADSHEET_ID) return [];
    try {
        const jobs = apiGetTableData('Jobs');
        return jobs.reverse(); // Newest first
    } catch (e) {
        return [];
    }
}

function apiCreateJob(formObject) {
    try {
        let sheet = getSheetByName('Jobs');
        if (!sheet) {
             const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
             sheet = ss.insertSheet('Jobs');
             sheet.appendRow(['ID', 'Title', 'Department', 'Location', 'Type', 'Status', 'Created_Date', 'Description']);
        }
        
        // Handle empty sheet case for Jobs
        if (sheet.getLastColumn() === 0) {
             sheet.appendRow(['ID', 'Title', 'Department', 'Location', 'Type', 'Status', 'Created_Date', 'Description']);
        }

        const newId = 'J' + new Date().getTime();
        const createdDate = new Date().toISOString().slice(0, 10);

        const row = [
            newId,
            formObject.title,
            formObject.department,
            formObject.location,
            formObject.type,
            'Open',
            createdDate,
            formObject.description
        ];
        
        sheet.appendRow(row);
        
        return { success: true, message: 'Đã tạo tin tuyển dụng thành công!' };
    } catch (e) {
        return { success: false, message: e.toString() };
    }
}

// 8. API: SETTINGS & ADMINISTRATION
function apiGetSettings() {
    let stages = apiGetTableData('Stages');
    if(stages.length === 0) {
      stages = [
           {Stage_Name: 'Apply', Order: 1, Color: '#0d6efd'},
           {Stage_Name: 'Interview', Order: 2, Color: '#fd7e14'},
           {Stage_Name: 'Offer', Order: 3, Color: '#198754'},
           {Stage_Name: 'Rejected', Order: 4, Color: '#dc3545'}
      ];
    }

    return {
        users: apiGetTableData('Users'),
        stages: stages,
        departments: apiGetTableData('Departments')
    };
}

function apiSaveStages(stagesArray) {
    try {
        let sheet = getSheetByName('Stages');
        if (!sheet) {
            // Create Stages sheet if it doesn't exist
            const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
            sheet = ss.insertSheet('Stages');
        }
        
        sheet.clearContents();
        sheet.appendRow(['ID', 'Stage_Name', 'Order', 'Color']); // Header
        
        stagesArray.forEach(s => {
            sheet.appendRow([s.ID, s.Stage_Name, s.Order, s.Color]);
        });
        
        return { success: true, message: 'Đã lưu cấu hình quy trình!' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiCreateUser(user) {
    try {
        let sheet = getSheetByName('Users');
        if (!sheet) {
             const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
             sheet = ss.insertSheet('Users');
             sheet.appendRow(['Username', 'Password', 'Full_Name', 'Role', 'Email', 'Department']);
        }
        
        // Check duplicate
        const users = apiGetTableData('Users');
        if(users.find(u => u.Username === user.username)) {
            return { success: false, message: 'Tên đăng nhập đã tồn tại' };
        }
        
        // Handle empty sheet (if exists but cleared)
        if (sheet.getLastColumn() === 0) {
            sheet.appendRow(['Username', 'Password', 'Full_Name', 'Role', 'Email', 'Department']);
        }
        
        sheet.appendRow([user.username, user.password, user.fullname, user.role, user.email, user.department]);
        return { success: true, message: 'Đã thêm người dùng' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiDeleteUser(username) {
    try {
        const sheet = getSheetByName('Users');
        if(!sheet) return { success: false, message: 'Users sheet not found' };
        
        const data = sheet.getDataRange().getValues();
        for(let i=1; i<data.length; i++) {
            if(data[i][0] == username) {
                sheet.deleteRow(i+1);
                return { success: true };
            }
        }
        return { success: false, message: 'User not found' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

function apiGetEmailTemplates() {
    try {
        let sheet = getSheetByName('Email_Templates');
        if (!sheet) {
            const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
            sheet = ss.insertSheet('Email_Templates');
            sheet.appendRow(['ID', 'Name', 'Subject', 'Body']);
            // Add defaults
            sheet.appendRow(['1', 'Offer Email', 'Mời nhận việc - [Candidate Name]', 'Chào [Name],\n\nChúc mừng bạn đã trúng tuyển...']);
            sheet.appendRow(['2', 'Reject Email', 'Thông báo kết quả phỏng vấn', 'Chào [Name],\n\nCảm ơn bạn đã quan tâm...']);
            sheet.appendRow(['3', 'Interview Invite', 'Mời phỏng vấn', 'Chào [Name],\n\nChúng tôi muốn mời bạn tham gia phỏng vấn...']);
        }
        
        // Handle empty/new sheet
        if (sheet.getLastColumn() === 0) {
            sheet.appendRow(['ID', 'Name', 'Subject', 'Body']);
            // Add defaults
            sheet.appendRow(['1', 'Offer Email', 'Mời nhận việc - [Candidate Name]', 'Chào [Name],\n\nChúc mừng bạn đã trúng tuyển...']);
            sheet.appendRow(['2', 'Reject Email', 'Thông báo kết quả phỏng vấn', 'Chào [Name],\n\nCảm ơn bạn đã quan tâm...']);
            sheet.appendRow(['3', 'Interview Invite', 'Mời phỏng vấn', 'Chào [Name],\n\nChúng tôi muốn mời bạn tham gia phỏng vấn...']);
        }
        
        return apiGetTableData('Email_Templates');
    } catch (e) {
        return [];
    }
}

function apiSaveEmailTemplate(template) {
    try {
        const sheet = getSheetByName('Email_Templates');
        if (!sheet) return { success: false, message: 'Database not found' };
        
        const data = sheet.getDataRange().getValues();
        // Simple Update by ID (Name in this case or ID)
        for(let i=1; i<data.length; i++) {
            if(data[i][0] == template.id) {
                sheet.getRange(i+1, 3).setValue(template.subject);
                sheet.getRange(i+1, 4).setValue(template.body);
                return { success: true, message: 'Đã lưu mẫu email' };
            }
        }
        return { success: false, message: 'Template not found' };
    } catch(e) {
        return { success: false, message: e.toString() };
    }
}

// ============================================
// DEPARTMENT & POSITION MANAGEMENT APIs
// ============================================

// Get all departments and their positions
function apiGetDepartments() {
  Logger.log('=== GET DEPARTMENTS ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    
    // If sheet doesn't exist, create it with sample data
    if (!sheet) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const newSheet = ss.insertSheet(SETTINGS_SHEET_NAME);
      
      // Set headers
      newSheet.getRange('A1').setValue('Phòng ban');
      newSheet.getRange('B1').setValue('Vị trí 1');
      newSheet.getRange('C1').setValue('Vị trí 2');
      newSheet.getRange('D1').setValue('Vị trí 3');
      
      // Add sample data
      newSheet.getRange('A2').setValue('Phòng nhân sự');
      newSheet.getRange('B2').setValue('Tuyển dụng');
      newSheet.getRange('C2').setValue('Đào tạo');
      
      newSheet.getRange('A3').setValue('Phòng kế toán');
      newSheet.getRange('B3').setValue('Kế toán trưởng');
      newSheet.getRange('C3').setValue('Kế toán thuế');
      
      Logger.log('Created new settings sheet with sample data');
      
      return {
        success: true,
        departments: [
          { name: 'Phòng nhân sự', positions: ['Tuyển dụng', 'Đào tạo'] },
          { name: 'Phòng kế toán', positions: ['Kế toán trưởng', 'Kế toán thuế'] }
        ]
      };
    }
    
    const data = sheet.getDataRange().getValues();
    const departments = [];
    
    // Skip header row (row 0)
    for (let i = 1; i < data.length; i++) {
      const deptName = data[i][0];
      if (!deptName) continue; // Skip empty rows
      
      const positions = [];
      // Collect all non-empty positions from columns B onwards
      for (let j = 1; j < data[i].length; j++) {
        const position = data[i][j];
        if (position && position.toString().trim()) {
          positions.push(position.toString().trim());
        }
      }
      
      departments.push({
        name: deptName.toString().trim(),
        positions: positions
      });
    }
    
    Logger.log('Found ' + departments.length + ' departments');
    return { success: true, departments: departments };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString(), departments: [] };
  }
}

// Add new department
function apiAddDepartment(deptName) {
  Logger.log('=== ADD DEPARTMENT: ' + deptName + ' ===');
  
  try {
    let sheet = getSheetByName(SETTINGS_SHEET_NAME);
    
    // Create sheet if doesn't exist
    if (!sheet) {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
      sheet.getRange('A1').setValue('Phòng ban');
    }
    
    // Check if department already exists
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === deptName.toLowerCase()) {
        return { success: false, message: 'Phòng ban đã tồn tại' };
      }
    }
    
    // Add to next empty row
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1).setValue(deptName);
    
    Logger.log('Added department at row ' + (lastRow + 1));
    return { success: true, message: 'Đã thêm phòng ban' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Add position to department
function apiAddPosition(deptName, position) {
  Logger.log('=== ADD POSITION: ' + position + ' to ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        // Find next empty column in this row
        let emptyCol = -1;
        for (let j = 1; j < data[i].length + 5; j++) { // Check a few extra columns
          if (!data[i][j] || !data[i][j].toString().trim()) {
            emptyCol = j + 1; // +1 for 1-indexed
            break;
          }
        }
        
        if (emptyCol === -1) {
          // No empty column found in existing data, add to next column
          emptyCol = data[i].length + 1;
        }
        
        sheet.getRange(i + 1, emptyCol).setValue(position);
        Logger.log('Added position at row ' + (i + 1) + ', col ' + emptyCol);
        return { success: true, message: 'Đã thêm vị trí' };
      }
    }
    
    return { success: false, message: 'Phòng ban không tồn tại' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Delete department
function apiDeleteDepartment(deptName) {
  Logger.log('=== DELETE DEPARTMENT: ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find and delete department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        sheet.deleteRow(i + 1);
        Logger.log('Deleted department at row ' + (i + 1));
        return { success: true, message: 'Đã xóa phòng ban' };
      }
    }
    
    return { success: false, message: 'Phòng ban không tồn tại' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Delete position from department
function apiDeletePosition(deptName, position) {
  Logger.log('=== DELETE POSITION: ' + position + ' from ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        // Find position column
        for (let j = 1; j < data[i].length; j++) {
          if (data[i][j] && data[i][j].toString().trim() === position) {
            sheet.getRange(i + 1, j + 1).setValue('');
            Logger.log('Deleted position at row ' + (i + 1) + ', col ' + (j + 1));
            return { success: true, message: 'Đã xóa vị trí' };
          }
        }
        return { success: false, message: 'Vị trí không tồn tại' };
      }
    }
    
    return { success: false, message: 'Phòng ban không tồn tại' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// EDIT DEPARTMENT NAME
function apiEditDepartment(oldName, newName) {
  Logger.log('=== EDIT DEPARTMENT: ' + oldName + ' to ' + newName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    // Validate new name
    if (!newName || !newName.trim()) {
      return { success: false, message: 'Tên phòng ban không được để trống' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Check if new name already exists (and it's not the same department)
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim().toLowerCase() === newName.toLowerCase() && data[i][0].toString().trim() !== oldName) {
        return { success: false, message: 'Tên phòng ban đã tồn tại' };
      }
    }
    
    // Find and update department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === oldName) {
        sheet.getRange(i + 1, 1).setValue(newName);
        Logger.log('Updated department at row ' + (i + 1));
        return { success: true, message: 'Đã cập nhật tên phòng ban' };
      }
    }
    
    return { success: false, message: 'Phòng ban không tồn tại' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}

// EDIT POSITION NAME
function apiEditPosition(deptName, oldPosition, newPosition) {
  Logger.log('=== EDIT POSITION: ' + oldPosition + ' to ' + newPosition + ' in ' + deptName + ' ===');
  
  try {
    const sheet = getSheetByName(SETTINGS_SHEET_NAME);
    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }
    
    // Validate new position
    if (!newPosition || !newPosition.trim()) {
      return { success: false, message: 'Tên vị trí không được để trống' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    // Find department row
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] && data[i][0].toString().trim() === deptName) {
        // Find old position column
        for (let j = 1; j < data[i].length; j++) {
          if (data[i][j] && data[i][j].toString().trim() === oldPosition) {
            sheet.getRange(i + 1, j + 1).setValue(newPosition);
            Logger.log('Updated position at row ' + (i + 1) + ', col ' + (j + 1));
            return { success: true, message: 'Đã cập nhật tên vị trí' };
          }
        }
        return { success: false, message: 'Vị trí không tồn tại' };
      }
    }
    
    return { success: false, message: 'Phòng ban không tồn tại' };
    
  } catch (e) {
    Logger.log('ERROR: ' + e.toString());
    return { success: false, message: e.toString() };
  }
}
