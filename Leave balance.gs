
/**
 * IN-MEMORY LEAVE BALANCE CALCULATOR
 * Dynamically calculates entitlements, aggregates usage, and computes final balances.
 * 
 * @param {string} optEmpId - (Optional) Pass an Emp ID for fast, single-user HTML UI lookup.
 * @param {string} optBU - (Optional) Filter by Business Unit for reporting.
 * @param {string} optDept - (Optional) Filter by Department for reporting.
 * @returns {Object} A structured object containing employee details, entitlements, usage, and balances.
 */
function getLeaveBalancesData(optEmpId, optBU, optDept) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Fetch Sheets
  var empSheet = ss.getSheetByName('tblEmployee');
  var policySheet = ss.getSheetByName('Sys_LeavePolicies');
  var leaveSheet = ss.getSheetByName('tblLeave');
  var startBalSheet = ss.getSheetByName('StartingBal');
  
  if (!empSheet || !policySheet || !leaveSheet) {
    throw new Error("Missing required sheets for balance calculation.");
  }

  var today = new Date();
  var currentYear = today.getFullYear();
  var prevYear = currentYear - 1;

  // ==========================================
  // 2. PARSE POLICIES & VISIBILITY RULES
  // ==========================================
  var policyData = policySheet.getDataRange().getValues();
  var pHeaders = policyData.shift().map(function(h) { return String(h).trim(); });
  
  var pCol = {
    bu: pHeaders.indexOf("Business Unit"),
    cat: pHeaders.indexOf("Category"),
    status: pHeaders.indexOf("Status"),
    gender: pHeaders.indexOf("Gender"),
    lifecycle: pHeaders.indexOf("Who is entitled (Lifecycle)"),
    leaveType: pHeaders.indexOf("Leave Type"),
    entitlement: pHeaders.indexOf("Annual Entitlements"),
    showBalance: pHeaders.indexOf("Show in Balance Page"), // Assuming this column exists
    carryDeadline: pHeaders.indexOf("Carry Forward Deadline")
  };

  var visibleLeaveTypes = []; // Leave types to return in the payload
  var carryDeadlines = {};    // Deadlines per leave type
  
  for (var p = 0; p < policyData.length; p++) {
    var lType = String(policyData[p][pCol.leaveType]).trim();
    var show = String(policyData[p][pCol.showBalance]).trim().toLowerCase();
    var deadlineRaw = String(policyData[p][pCol.carryDeadline]).trim();
    
    if (lType && visibleLeaveTypes.indexOf(lType) === -1 && (show === "yes" || show === "true")) {
      visibleLeaveTypes.push(lType);
    }
    
    // Parse deadline date (e.g., "30/06" or standard date)
    if (deadlineRaw && deadlineRaw.toLowerCase() !== "no") {
      var deadlineDate = parseDDMMYYYY(deadlineRaw); // Utilizing your existing helper
      // If no year specified in policy, assume current year for the deadline
      deadlineDate.setFullYear(currentYear); 
      carryDeadlines[lType] = deadlineDate;
    }
  }

  // ==========================================
  // 3. AGGREGATE LEAVE USAGE (FROM tblLeave)
  // ==========================================
  var leaveData = leaveSheet.getDataRange().getValues();
  var lHeaders = leaveData.shift().map(function(h) { return String(h).trim(); });
  
  var lEmpIdx = lHeaders.indexOf("Emp ID");
  var lTypeIdx = lHeaders.indexOf("Leave Type");
  var lUtilIdx = lHeaders.indexOf("Leave Utilized");
  var lYearIdx = lHeaders.indexOf("Entitlement Year");
  var lStatusIdx = lHeaders.indexOf("Status"); // To exclude rejected/cancelled
  
  var usageMap = {}; 
  // Structure: usageMap[empId][leaveType][year] = totalDaysUtilized
  
  for (var l = 0; l < leaveData.length; l++) {
    var row = leaveData[l];
    var status = lStatusIdx !== -1 ? String(row[lStatusIdx]).trim().toLowerCase() : "approved";
    
    if (status === "rejected" || status === "cancelled") continue;
    
    var eId = String(row[lEmpIdx]).trim().toUpperCase();
    var lType = String(row[lTypeIdx]).trim();
    var utilized = Number(row[lUtilIdx]) || 0;
    var eYear = Number(row[lYearIdx]) || currentYear;
    
    if (!eId) continue;
    if (!usageMap[eId]) usageMap[eId] = {};
    if (!usageMap[eId][lType]) usageMap[eId][lType] = {};
    if (!usageMap[eId][lType][eYear]) usageMap[eId][lType][eYear] = 0;
    
    usageMap[eId][lType][eYear] += utilized;
  }

  // ==========================================
  // 4. LOAD STARTING BALANCES (PREVIOUS YEAR)
  // ==========================================
  var carryOverMap = {};
  if (startBalSheet) {
    var sbData = startBalSheet.getDataRange().getValues();
    if (sbData.length > 1) {
      var sbHeaders = sbData.shift().map(function(h) { return String(h).trim(); });
      var sbEmpIdx = sbHeaders.indexOf("Emp ID") > -1 ? sbHeaders.indexOf("Emp ID") : sbHeaders.indexOf("Emp No");
      var sbBalIdx = sbHeaders.indexOf(prevYear + " Balance"); 
      
      if (sbEmpIdx !== -1 && sbBalIdx !== -1) {
        for (var b = 0; b < sbData.length; b++) {
          var eId = String(sbData[b][sbEmpIdx]).trim().toUpperCase();
          carryOverMap[eId] = Number(sbData[b][sbBalIdx]) || 0;
        }
      }
    }
  }

  // ==========================================
  // 5. CALCULATE ENTITLEMENTS & FINAL BALANCES
  // ==========================================
  var empData = empSheet.getDataRange().getValues();
  var eHeaders = empData.shift().map(function(h) { return String(h).trim(); });
  
  var outputData = {}; 
  
  for (var e = 0; e < empData.length; e++) {
    var emp = empData[e];
    var empId = String(emp[eHeaders.indexOf("Emp ID")]).trim().toUpperCase();
    
    if (!empId) continue;
    
    // UI/Report Filters
    if (optEmpId && empId !== String(optEmpId).trim().toUpperCase()) continue;
    
    var bu = String(emp[eHeaders.indexOf("Business Unit")]).trim();
    var dept = String(emp[eHeaders.indexOf("Department")]).trim();
    var catFull = String(emp[eHeaders.indexOf("Category")]).trim();
    var empStatus = String(emp[eHeaders.indexOf("Status")]).trim();
    var gender = String(emp[eHeaders.indexOf("Gender")]).trim();
    var empName = String(emp[eHeaders.indexOf("Emp Name")]).trim();
    
    if (optBU && bu.toUpperCase() !== String(optBU).toUpperCase()) continue;
    if (optDept && dept.toUpperCase() !== String(optDept).toUpperCase()) continue;
    
    var doj = parseDDMMYYYY(String(emp[eHeaders.indexOf("Date of Join")]).trim());
    var catInitials = mapCategoryToInitials(catFull); // Existing helper
    
    // Initialize Employee Profile
    outputData[empId] = {
      profile: { id: empId, name: empName, bu: bu, dept: dept, category: catFull },
      balances: {},     // Final calculated available days
      entitlements: {}, // Base entitlements as per policy
      usage: {}         // Consumed days
    };
    
    // Evaluate only leaves marked to show on the Balance Page
    for (var i = 0; i < visibleLeaveTypes.length; i++) {
      var currentLeaveType = visibleLeaveTypes[i];
      var bestScore = -1;
      var determinedEntitlement = 0; // Default state
      var isUnlimited = false;
      
      // Calculate Entitlement (Reusing your tested Engine Logic)
      for (var p = 0; p < policyData.length; p++) {
        var pol = policyData[p];
        if (String(pol[pCol.leaveType]).trim() !== currentLeaveType) continue;
        
        var scoreBU = checkMatch(bu.toUpperCase(), String(pol[pCol.bu]).trim().toUpperCase(), 100);
        var scoreCat = checkMatch(catInitials, String(pol[pCol.cat]).trim(), 10);
        var scoreStatus = checkMatch(empStatus, String(pol[pCol.status]).trim(), 5);
        var scoreGender = checkMatch(gender, String(pol[pCol.gender]).trim(), 1);
        
        if (scoreBU === -1 || scoreCat === -1 || scoreStatus === -1 || scoreGender === -1) continue; 
        
        var lifecycleRule = String(pol[pCol.lifecycle]).trim();
        if (!evaluateLifecycle(doj, lifecycleRule, today)) continue;
        
        var totalScore = scoreBU + scoreCat + scoreStatus + scoreGender;
        
        if (totalScore > bestScore) {
          bestScore = totalScore;
          var entValue = String(pol[pCol.entitlement]).trim();
          if (entValue === "" || entValue.toLowerCase() === "unlimited") {
            isUnlimited = true;
          } else {
            determinedEntitlement = Number(entValue) || 0;
          }
        }
      }
      
      // Calculate Year Balances
      var currYearUsage = (usageMap[empId] && usageMap[empId][currentLeaveType] && usageMap[empId][currentLeaveType][currentYear]) ? usageMap[empId][currentLeaveType][currentYear] : 0;
      var prevYearUsage = (usageMap[empId] && usageMap[empId][currentLeaveType] && usageMap[empId][currentLeaveType][prevYear]) ? usageMap[empId][currentLeaveType][prevYear] : 0;
      
      var prevYearAvailable = 0;
      var currentYearAvailable = 0;
      var totalAvailable = 0;
      
      if (isUnlimited) {
        totalAvailable = "Unlimited";
      } else {
        // Carry Forward Logic
        var deadline = carryDeadlines[currentLeaveType];
        
        if (deadline && today <= deadline) {
          // It's before the deadline: calculate Prev Year Remainder
          // Assuming StartingBal sheet tracks the gross carry forward available for Annual Leaves
          var baseCarryForward = (currentLeaveType.toLowerCase().indexOf("annual") > -1) ? (carryOverMap[empId] || 0) : 0;
          
          prevYearAvailable = Math.max(0, baseCarryForward - prevYearUsage);
        }
        
        currentYearAvailable = Math.max(0, determinedEntitlement - currYearUsage);
        totalAvailable = currentYearAvailable + prevYearAvailable;
      }
      
      // Store to payload
      outputData[empId].entitlements[currentLeaveType] = isUnlimited ? "Unlimited" : determinedEntitlement;
      outputData[empId].usage[currentLeaveType] = currYearUsage + prevYearUsage; // Total usage spanning active years
      outputData[empId].balances[currentLeaveType] = totalAvailable;
    }
  }
  
  return outputData;
}

function apiGetEmployeeBalance(empId) {
  var data = getLeaveBalancesData(empId); // Passes the ID to filter instantly
  return data[empId] ? data[empId] : { error: "Employee Not Found" };
}

function generateReportArray(buFilter, deptFilter) {
  var rawData = getLeaveBalancesData(null, buFilter, deptFilter);
  var report = [["Emp ID", "Name", "Department", "BU", "Annual Leave Balance", "Sick Leave Balance"]]; // Dynamically build headers based on visible leaves

  for (var empId in rawData) {
    var emp = rawData[empId];
    report.push([
      emp.profile.id, 
      emp.profile.name, 
      emp.profile.dept, 
      emp.profile.bu, 
      emp.balances["Annual Leave"] || 0,
      emp.balances["Sick Leave"] || 0
    ]);
  }
  return report;
}


