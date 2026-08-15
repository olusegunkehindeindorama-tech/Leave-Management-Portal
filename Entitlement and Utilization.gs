
/**
 * LEAVE ENTITLEMENT RULES ENGINE & UTILIZATION CALCULATOR
 * This script tests the policy engine by matching every employee against the Sys_LeavePolicies 
 * and generates a matrix of their entitlements. It also calculates utilized leave days.
 */

function generateEntitlementMatrix() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get Data from Sheets
  var empSheet = ss.getSheetByName('tblEmployee');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  
  if (!empSheet || !policySheet) {
    SpreadsheetApp.getUi().alert("Error: Ensure 'tblEmployee' and 'Sys_LeavePolicies' sheets exist.");
    return;
  }

  var empData = empSheet.getDataRange().getValues();
  var policyData = policySheet.getDataRange().getValues();
  
  // Extract headers
  var empHeaders = empData.shift();
  var policyHeaders = policyData.shift();
  
  // Identify Policy Columns (0-indexed)
  var pCol = {
    bu: policyHeaders.indexOf("Business Unit"),
    cat: policyHeaders.indexOf("Category"),
    status: policyHeaders.indexOf("Status"),
    gender: policyHeaders.indexOf("Gender"),
    lifecycle: policyHeaders.indexOf("Who is entitled (Lifecycle)"),
    leaveType: policyHeaders.indexOf("Leave Type"),
    entitlement: policyHeaders.indexOf("Annual Entitlements") // Using exact header from your CSV
  };

  // Extract all unique leave types for the columns of our new sheet
  var uniqueLeaveTypes = [];
  for (var i = 0; i < policyData.length; i++) {
    var type = policyData[i][pCol.leaveType];
    if (type && uniqueLeaveTypes.indexOf(type) === -1) {
      uniqueLeaveTypes.push(type);
    }
  }
  
  // Prepare Output Data Array [ [Emp ID, Emp Name, LeaveType1, LeaveType2, ...] ]
  var outputHeaders = ["Emp ID", "Emp Name", "BU", "Category", "Status"].concat(uniqueLeaveTypes);
  var outputData = [outputHeaders];
  
  var today = new Date(); // Evaluate as of right now
  
  // LOOP THROUGH EVERY EMPLOYEE
  for (var e = 0; e < empData.length; e++) {
    var emp = empData[e];
    if (!emp[1]) continue; // Skip if no Emp ID
    
    var empId = emp[1]; // Assuming Col B is Emp ID
    var empName = emp[5]; // Assuming Col F is Emp Name
    var bu = String(emp[0]).trim().toUpperCase();
    var catFull = String(emp[2]).trim();
    var status = String(emp[3]).trim();
    var gender = String(emp[4]).trim();
    
    // Parse Date of Join (Assuming DD/MM/YYYY format based on CSV)
    var dojRaw = String(emp[7]).trim(); // Col H is Date of Join
    var doj = parseDDMMYYYY(dojRaw);
    
    var catInitials = mapCategoryToInitials(catFull);
    
    var rowOut = [empId, empName, bu, catFull, status];
    
    // Check against every unique leave type
    for (var l = 0; l < uniqueLeaveTypes.length; l++) {
      var currentLeaveType = uniqueLeaveTypes[l];
      
      var bestScore = -1;
      var determinedEntitlement = "Not Entitled"; // Default state
      
      // Filter policy rows for this specific leave type
      for (var p = 0; p < policyData.length; p++) {
        var pol = policyData[p];
        if (pol[pCol.leaveType] !== currentLeaveType) continue;
        
        // 1. Evaluate A-D Fields using Weighted Scores
        var scoreBU = checkMatch(bu, String(pol[pCol.bu]).trim().toUpperCase(), 100);
        var scoreCat = checkMatch(catInitials, String(pol[pCol.cat]).trim(), 10);
        var scoreStatus = checkMatch(status, String(pol[pCol.status]).trim(), 5);
        var scoreGender = checkMatch(gender, String(pol[pCol.gender]).trim(), 1);
        
        // If ANY field is a hard mismatch (-1), this policy row is entirely disqualified
        if (scoreBU === -1 || scoreCat === -1 || scoreStatus === -1 || scoreGender === -1) {
          continue; 
        }
        
        // 2. Evaluate Lifecycle (DOJ vs Today)
        var lifecycleRule = String(pol[pCol.lifecycle]).trim();
        var isLifecycleValid = evaluateLifecycle(doj, lifecycleRule, today);
        if (!isLifecycleValid) {
          continue; // Disqualified by time
        }
        
        // 3. Calculate Total Score for this valid policy
        var totalScore = scoreBU + scoreCat + scoreStatus + scoreGender;
        
        // 4. Check if this is the strongest matching rule so far
        if (totalScore > bestScore) {
          bestScore = totalScore;
          var entValue = String(pol[pCol.entitlement]).trim();
          determinedEntitlement = (entValue === "") ? "Unlimited" : entValue;
        }
      }
      
      rowOut.push(determinedEntitlement);
    }
    
    outputData.push(rowOut);
  }
  
  // WRITE TO SHEET
  var outSheetName = "Entitlement_Test";
  var outSheet = ss.getSheetByName(outSheetName);
  if (outSheet) {
    outSheet.clear(); // Clear old test if it exists
  } else {
    outSheet = ss.insertSheet(outSheetName);
  }
  
  outSheet.getRange(1, 1, outputData.length, outputData[0].length).setValues(outputData);
  
  // Formatting for readability
  outSheet.getRange(1, 1, 1, outputData[0].length).setFontWeight("bold").setBackground("#d9ead3");
  outSheet.autoResizeColumns(1, outputData[0].length);
  
  SpreadsheetApp.getUi().alert("Matrix generation complete! Check the 'Entitlement_Test' sheet.");
}

/**
 * CALCULATE LEAVE UTILIZED & ENTITLEMENT YEAR
 * Evaluates records in tblLeave, computes utilized days, checks carry-over balances,
 * automatically merges previously split rows, and safely splits transactions 
 * crossing two entitlement years using -a and -b suffixes.
 */
function calculateLeaveUtilized() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var leaveSheet = ss.getSheetByName('tblLeave');
  var shiftSheet = ss.getSheetByName('tblShift');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  var startBalSheet = ss.getSheetByName('StartingBal');
  
  if (!leaveSheet || !shiftSheet || !policySheet || !startBalSheet) {
    SpreadsheetApp.getUi().alert("Error: Ensure 'tblLeave', 'tblShift', 'Sys_LeavePolicies', and 'StartingBal' sheets exist.");
    return;
  }
  
  // --- DYNAMIC YEAR CALCULATION ---
  var currentYear = new Date().getFullYear(); // e.g., 2026
  var prevYear = currentYear - 1;             // e.g., 2025
  var prevYearBalColumn = prevYear + " Balance"; // e.g., "2025 Balance"
  
  // 1. Load Policy Maps (Calculation Methods & Carry Forward Deadline)
  var policyData = policySheet.getDataRange().getValues();
  var pHeaders = policyData.shift();
  var pTypeIdx = pHeaders.indexOf("Leave Type");
  var pMethodIdx = pHeaders.indexOf("Calculation Method");
  var pCarryIdx = pHeaders.indexOf("Carry Forward Deadline");
  
  var methodMap = {};
  var carryForwardMap = {}; 
  
  for (var p = 0; p < policyData.length; p++) {
    var type = String(policyData[p][pTypeIdx]).trim();
    var method = String(policyData[p][pMethodIdx]).trim();
    var carryDeadline = String(policyData[p][pCarryIdx]).trim();
    
    if (type) {
      if (!methodMap[type]) methodMap[type] = method;
      if (carryDeadline && carryDeadline.toLowerCase() !== "no") {
          var parsedDate = parseDDMMYYYY(carryDeadline);
          parsedDate.setFullYear(currentYear); 
          carryForwardMap[type] = parsedDate;
        }
    }
  }
  
  // 2. Load Shift Roaster into an Object Dictionary
  var shiftData = shiftSheet.getDataRange().getValues();
  var shiftHeaders = shiftData.shift(); 
  
  var shiftMap = {};
  for (var s = 0; s < shiftData.length; s++) {
    var eId = String(shiftData[s][0]).trim().toUpperCase();
    var rawDate = shiftData[s][1];
    var shiftCode = String(shiftData[s][2]).trim().toUpperCase();
    
    var dateObj = new Date(rawDate);
    if (isNaN(dateObj.getTime())) continue;
    var dateStr = formatDateKey(dateObj); 
    
    if (!shiftMap[eId]) shiftMap[eId] = {};
    shiftMap[eId][dateStr] = shiftCode;
  }
  
  // 3. Load Starting Balances (Previous Year Carry Over) dynamically
  var sbData = startBalSheet.getDataRange().getValues();
  var carryOverMap = {};
  if (sbData.length > 0) {
    var sbHeaders = sbData.shift();
    var sbEmpIdx = sbHeaders.indexOf("Emp No") !== -1 ? sbHeaders.indexOf("Emp No") : sbHeaders.indexOf("Emp ID");
    var sbBalIdx = sbHeaders.indexOf(prevYearBalColumn); // Dynamically finds "2025 Balance"
    
    if (sbEmpIdx !== -1 && sbBalIdx !== -1) {
      for (var b = 0; b < sbData.length; b++) {
        var eId = String(sbData[b][sbEmpIdx]).trim().toUpperCase();
        var bal = Number(sbData[b][sbBalIdx]) || 0;
        carryOverMap[eId] = bal;
      }
    } else {
      SpreadsheetApp.getUi().alert("Warning: Could not find '" + prevYearBalColumn + "' column in StartingBal sheet. Carry over balances will be 0.");
    }
  }
  
  // 4. Process Leaves in tblLeave (with Pre-Merge Logic)
  var rawLeaveData = leaveSheet.getDataRange().getValues();
  var lHeaders = rawLeaveData[0];
  var lEntryIdx = lHeaders.indexOf("Entry Code");
  var lEmpIdx = lHeaders.indexOf("Emp ID");
  var lTypeIdx = lHeaders.indexOf("Leave Type");
  var lStartIdx = lHeaders.indexOf("Start Date");
  var lEndIdx = lHeaders.indexOf("End Date");
  var lUtilIdx = lHeaders.indexOf("Leave Utilized");
  var lYearIdx = lHeaders.indexOf("Entitlement Year");
  
  if (lUtilIdx === -1 || lYearIdx === -1) {
    SpreadsheetApp.getUi().alert("Error: 'Leave Utilized' or 'Entitlement Year' column not found in 'tblLeave'.");
    return;
  }
  
  // PRE-MERGE: Remove -b rows and strip -a from Entry Codes to restore to original state
  var baseLeaveData = [];
  for (var r = 1; r < rawLeaveData.length; r++) {
    var row = rawLeaveData[r];
    var entryCode = String(row[lEntryIdx]).trim();
    
    if (entryCode.match(/-b$/)) {
      continue; // Skip the 'b' half entirely, deleting it from memory
    }
    if (entryCode.match(/-a$/)) {
      row[lEntryIdx] = entryCode.replace(/-a$/, ""); // Strip the 'a' to restore original code
    }
    baseLeaveData.push(row);
  }
  
  var newLeaveData = [];
  newLeaveData.push(lHeaders); // Push headers
  
  // PROCESS EACH UNIFIED ROW
  for (var i = 0; i < baseLeaveData.length; i++) {
    var row = baseLeaveData[i];
    var empId = String(row[lEmpIdx]).trim().toUpperCase();
    var leaveType = String(row[lTypeIdx]).trim();
    var startDateRaw = row[lStartIdx];
    var endDateRaw = row[lEndIdx];
    
    var startDate = new Date(startDateRaw);
    var endDate = new Date(endDateRaw);
    
    // Safety check for empty or invalid rows
    if (!empId || isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      newLeaveData.push(row);
      continue;
    }
    
    var calcMethod = methodMap[leaveType] || "ActualDays";
    var isAnnual = leaveType.toLowerCase().indexOf("annual") > -1; //Comment out to make all 1.5
    var totalDays = 0;
    
    // --- Step A: Calculate Total Utilized Days ---
    if (calcMethod === "ActualDays") {
      var diffTime = endDate.getTime() - startDate.getTime();
      totalDays = Math.round(diffTime / (1000 * 3600 * 24)) + 1; 
      
    } else if (calcMethod === "ShiftRoaster") {
      var currDate = new Date(startDate.getTime());
      
      while (currDate <= endDate) {
        var dateStr = formatDateKey(currDate);
        var shiftCode = (shiftMap[empId] && shiftMap[empId][dateStr]) ? shiftMap[empId][dateStr] : null;
        
        if (shiftCode) {
          if (["A", "B", "D", "N", "M"].indexOf(shiftCode) > -1) {
            totalDays += isAnnual ? 1.5 : 1; //Change to only 1.5 if apply to all
          } else if (shiftCode === "G") {
            totalDays += 1;
          } else if (shiftCode === "O") {
            totalDays += 0;
          } else {
            var day = currDate.getDay();
            totalDays += (day === 0 || day === 6) ? 0 : 1;
          }
        } else {
          var day = currDate.getDay();
          totalDays += (day === 0 || day === 6) ? 0 : 1;
        }
        currDate.setDate(currDate.getDate() + 1);
      }
    }
    
    // --- Step B: Determine Entitlement Year & Split Logic ---
    if (totalDays > 0 && carryForwardMap[leaveType]) {
      var availableCarryOver = carryOverMap[empId] || 0;
      var expiryDate = carryForwardMap[leaveType]; 
      
      // If leave starts after the dynamic deadline, carryover is strictly 0 (expired)
      if (expiryDate && startDate > expiryDate) {
        availableCarryOver = 0;
      }
      
      if (availableCarryOver >= totalDays) {
        // Completely absorbed by previous year
        row[lUtilIdx] = totalDays;
        row[lYearIdx] = prevYear;
        carryOverMap[empId] -= totalDays; // Deduct from memory
        newLeaveData.push(row);
        
      } else if (availableCarryOver > 0) {
        // SPLIT REQUIRED: Leave uses remaining carry over, and spills into current year
        
        // 1. First Row (Previous Year: Suffix -a)
        var rowPrevYear = row.slice(); // Duplicate the row
        if (lEntryIdx !== -1 && rowPrevYear[lEntryIdx]) rowPrevYear[lEntryIdx] = String(rowPrevYear[lEntryIdx]) + "-a"; 
        rowPrevYear[lUtilIdx] = availableCarryOver;
        rowPrevYear[lYearIdx] = prevYear;
        newLeaveData.push(rowPrevYear);
        
        // 2. Second Row (Current Year: Suffix -b)
        var remainder = totalDays - availableCarryOver;
        if (lEntryIdx !== -1 && row[lEntryIdx]) row[lEntryIdx] = String(row[lEntryIdx]) + "-b"; 
        row[lUtilIdx] = remainder;
        row[lYearIdx] = currentYear;
        carryOverMap[empId] = 0; // Carry over completely depleted
        newLeaveData.push(row);
        
      } else {
        // Standard: No carry over left, completely absorbed by current year
        row[lUtilIdx] = totalDays;
        row[lYearIdx] = currentYear;
        newLeaveData.push(row);
      }
      
    } else {
      // Non-carry forward leaves strictly use current year
      row[lUtilIdx] = totalDays;
      row[lYearIdx] = currentYear;
      newLeaveData.push(row);
    }
  }
  
  // 5. Write Complete Updates back to sheet
  leaveSheet.clearContents();
  leaveSheet.getRange(1, 1, newLeaveData.length, newLeaveData[0].length).setValues(newLeaveData);
  
}

/** * HELPER: Evaluates exact match vs DEFAULT vs Mismatch. 
 * Returns assigned weight if exact, 0 if DEFAULT, -1 if mismatch.
 */
function checkMatch(empValue, policyValue, weight) {
  if (policyValue === "DEFAULT") return 0;
  if (empValue.toLowerCase() === policyValue.toLowerCase()) return weight;
  return -1; // Mismatch
}

/** * HELPER: Converts full category names to the initials used in the policy
 */
function mapCategoryToInitials(fullCategory) {
  var cat = fullCategory.toLowerCase();
  if (cat.indexOf("junior staff") > -1) return "JS";
  if (cat.indexOf("non union mgt") > -1) return "NUMS";
  if (cat.indexOf("non union senior") > -1) return "NUSS";
  if (cat.indexOf("senior staff") > -1) return "SS";
  if (cat.indexOf("trainee") > -1) return "T";
  return fullCategory; // Fallback
}

/** * HELPER: Parses DD/MM/YYYY dates from the CSV to JS Date objects
 */
function parseDDMMYYYY(dateString) {
  if (!dateString) return new Date();
  var parts = dateString.split("/");
  if (parts.length === 3) {
    // Note: JS Date is YYYY, MM (0-11), DD
    return new Date(parts[2], parts[1] - 1, parts[0]);
  }
  return new Date(dateString); // Fallback for standard formats
}

/** * HELPER: Evaluates strings like ">= 1 year", "< 6 months" against DOJ
 */
function evaluateLifecycle(doj, ruleString, todayDate) {
  if (!ruleString || ruleString === "DEFAULT" || ruleString === "") return true;
  
  var regex = /([><=]+)\s*(\d+)\s*(year|years|month|months|day|days)/i;
  var match = ruleString.match(regex);
  
  if (!match) return true; // If we can't parse the rule, assume no restriction
  
  var operator = match[1];
  var value = parseFloat(match[2]);
  var unit = match[3].toLowerCase();
  
  // Calculate difference in days
  var diffTime = todayDate.getTime() - doj.getTime();
  var diffDays = diffTime / (1000 * 3600 * 24);
  
  var targetDays = 0;
  if (unit.indexOf("year") > -1) targetDays = value * 365.25;
  if (unit.indexOf("month") > -1) targetDays = value * 30.4375;
  if (unit.indexOf("day") > -1) targetDays = value;
  
  switch(operator) {
    case ">=": return diffDays >= targetDays;
    case ">": return diffDays > targetDays;
    case "<=": return diffDays <= targetDays;
    case "<": return diffDays < targetDays;
    case "==": 
    case "=": return Math.round(diffDays) === Math.round(targetDays);
    default: return true;
  }
}

/** * HELPER: Formats a JS Date object into a YYYY-MM-DD string for dictionary mapping
 */
function formatDateKey(dateObj) {
  var y = dateObj.getFullYear();
  var m = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  var d = ("0" + dateObj.getDate()).slice(-2);
  return y + "-" + m + "-" + d;
}