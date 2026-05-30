/*************************************************
 * CUSTOMER MANAGEMENT SYSTEM
 * CODE.GS
 *************************************************/

const SHEET_NAME = "CUSTOMER_DATA";
const USERS_SHEET = "USERS";

/*************************************************
 * WEB APP
 *************************************************/
function doGet() {
  return HtmlService.createTemplateFromFile("Login")
    .evaluate()
    .setTitle("Customer Management System")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/*************************************************
 * LOGIN
 *************************************************/
function loginUser(username, password) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(USERS_SHEET);

  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {

    if (
      data[i][0].toString().trim() === username &&
      data[i][1].toString().trim() === password
    ) {

      return {
        success: true,
        role: data[i][2],
        staffName: data[i][0]
      };
    }
  }

  return {
    success: false
  };
}

/*************************************************
 * SAVE CUSTOMER
 *************************************************/
function saveCustomer(obj) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const serviceCharge = Number(obj.serviceCharge) || 0;
  const bankCharge = Number(obj.bankCharge) || 0;

  const totalAmount = serviceCharge + bankCharge;

  sheet.appendRow([
    Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    ),
    obj.staffName,
    obj.customerName,
    obj.serviceType,
    serviceCharge,
    bankCharge,
    totalAmount,
    obj.paymentMode,
    "Pending",
    new Date()
  ]);

  return {
    success: true
  };
}

/*************************************************
 * GET STAFF TODAY DATA
 *************************************************/
function getStaffTodayData(staffName) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("CUSTOMER_DATA");

  const values = sheet.getDataRange().getValues();

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  let rows = [];
  let totalService = 0;
  let totalBank = 0;
  let grandTotal = 0;
  let edistrictCount = 0;

  for (let i = 1; i < values.length; i++) {

    let rowDate = Utilities.formatDate(
      new Date(values[i][0]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    let rowStaff = String(values[i][1]).toLowerCase();

    if (
      rowDate === today &&
      rowStaff === String(staffName).toLowerCase()
    ) {

      let serviceCharge = Number(values[i][4]) || 0;
      let bankCharge = Number(values[i][5]) || 0;
      let totalAmount = Number(values[i][6]) || 0;

      totalService += serviceCharge;
      totalBank += bankCharge;
      grandTotal += totalAmount;

      if (
        String(values[i][3]).toLowerCase() === "e-district"
      ) {
        edistrictCount++;
      }

      rows.push({
        rowNumber: i + 1,
        slNo: rows.length + 1,
        customerName: values[i][2],
        serviceType: values[i][3],
        serviceCharge: serviceCharge,
        bankCharge: bankCharge,
        totalAmount: totalAmount,
        paymentMode: values[i][7],
        status: values[i][8]
      });

    }
  }

  return {
    rows: rows,
    totals: {
      totalService: totalService,
      totalBank: totalBank,
      grandTotal: grandTotal,
      edistrictCount: edistrictCount
    }
  };
}

/*************************************************
 * LAST CUSTOMER NAME
 *************************************************/
function getLastCustomerName(staffName) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const values = sheet.getDataRange().getValues();

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  let lastName = "";

  for (let i = values.length - 1; i >= 1; i--) {

    if (
      values[i][0] == today &&
      values[i][1] == staffName
    ) {
      lastName = values[i][2];
      break;
    }
  }

  return lastName;
}

/*************************************************
 * ADMIN DATA
 *************************************************/
function getAdminData(selectedDate) {

  const sheet = SpreadsheetApp.getActiveSpreadsheet()
    .getSheetByName("CUSTOMER_DATA");

  const values = sheet.getDataRange().getValues();

  let prathibha = [];
  let rejitha = [];

  let summary = {
    prathibha: {
      service: 0,
      bank: 0,
      total: 0
    },
    rejitha: {
      service: 0,
      bank: 0,
      total: 0
    }
  };

  for (let i = 1; i < values.length; i++) {

    const rowDate = Utilities.formatDate(
      new Date(values[i][0]),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    if (rowDate !== selectedDate) continue;

    const staff = String(values[i][1]).toLowerCase();

    const row = {
      rowNumber: i + 1,
      slNo: 0,
      customerName: values[i][2],
      serviceType: values[i][3],
      serviceCharge: values[i][4],
      bankCharge: values[i][5],
      totalAmount: values[i][6],
      paymentMode: values[i][7],
      status: values[i][8]
    };

    if (staff === "prathibha") {

      row.slNo = prathibha.length + 1;
      prathibha.push(row);

      summary.prathibha.service += Number(values[i][4]) || 0;
      summary.prathibha.bank += Number(values[i][5]) || 0;
      summary.prathibha.total += Number(values[i][6]) || 0;

    }

    if (staff === "rejitha") {

      row.slNo = rejitha.length + 1;
      rejitha.push(row);

      summary.rejitha.service += Number(values[i][4]) || 0;
      summary.rejitha.bank += Number(values[i][5]) || 0;
      summary.rejitha.total += Number(values[i][6]) || 0;

    }
  }

  return {
    prathibha: prathibha,
    rejitha: rejitha,
    summary: summary
  };
}

/*************************************************
 * UPDATE STATUS
 *************************************************/
function updateStatus(rowNumber, status) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  sheet.getRange(rowNumber, 9).setValue(status);

  return true;
}

/*************************************************
 * UPDATE CUSTOMER
 *************************************************/
function updateCustomer(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  const row = data.rowNumber;

  const serviceCharge = Number(data.serviceCharge) || 0;
  const bankCharge = Number(data.bankCharge) || 0;

  const total = serviceCharge + bankCharge;

  sheet.getRange(row, 3).setValue(data.customerName);
  sheet.getRange(row, 4).setValue(data.serviceType);
  sheet.getRange(row, 5).setValue(serviceCharge);
  sheet.getRange(row, 6).setValue(bankCharge);
  sheet.getRange(row, 7).setValue(total);
  sheet.getRange(row, 8).setValue(data.paymentMode);

  return {
    success: true
  };
}

/*************************************************
 * DELETE CUSTOMER
 *************************************************/
function deleteCustomer(rowNumber) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);

  sheet.deleteRow(rowNumber);

  return true;
}

/*************************************************
 * TODAY DATE
 *************************************************/
function getTodayDate() {

  return Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );
}

function getAdminPage() {
  return HtmlService.createHtmlOutputFromFile("Admin").getContent();
}

function getStaffPage() {
  return HtmlService.createHtmlOutputFromFile("Staff").getContent();
}
