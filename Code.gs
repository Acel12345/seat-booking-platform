// https://docs.google.com/spreadsheets/d/1FlOmAuWk4eQ2vF8oGC1E277CVB8IMhPTCaZY-lDLiZo/edit?usp=sharing

const CONFIG = {
  SPREADSHEET_ID: '1FlOmAuWk4eQ2vF8oGC1E277CVB8IMhPTCaZY-lDLiZo',

  SHEETS: {
    SEATS: 'SEATS',
    BOOKINGS: 'BOOKINGS'
  },

  HOLD_DAYS: 3
};

const STATUS = {
  AVAILABLE: 'AVAILABLE',
  ON_HOLD: 'ON_HOLD',
  BOOKED: 'BOOKED',
  BLOCKED: 'BLOCKED'
};
/* above codes are step to ADD THE MAIN CONFIGURATION*/

function getDatabase_() {
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}


function getTable_(sheetName) {
  const spreadsheet = getDatabase_();
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" was not found.`);
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  if (values.length === 0) {
    throw new Error(`Sheet "${sheetName}" is empty.`);
  }

  const headers = values[0].map(header =>
    String(header).trim()
  );

  const columns = {};

  headers.forEach((header, index) => {
    columns[header.toUpperCase()] = index;
  });

  return {
    sheet,
    values,
    headers,
    columns
  };
}


function column_(table, columnName) {
  const index = table.columns[columnName.toUpperCase()];

  if (index === undefined) {
    throw new Error(
      `Column "${columnName}" was not found in sheet "${table.sheet.getName()}".`
    );
  }

  return index;
}
/* above codes are step to ADD THE SPREADSHEET HELPER FUNCTIONS*/

function getSeats() {

  // First release any hold that already expired.
  releaseExpiredHolds();

  const table = getTable_(CONFIG.SHEETS.SEATS);

  const seatIdCol = column_(table, 'Seat ID');
  const rowCol = column_(table, 'Row');
  const numberCol = column_(table, 'Number');
  const zoneCol = column_(table, 'Zone');
  const priceCol = column_(table, 'Price');
  const xCol = column_(table, 'X');
  const yCol = column_(table, 'Y');
  const statusCol = column_(table, 'Status');
  const holdUntilCol = column_(table, 'Hold Until');

  return table.values
    .slice(1)
    .filter(row => row[seatIdCol] !== '')
    .map(row => {

      const holdUntil = row[holdUntilCol];

      return {
        seatId: String(row[seatIdCol]),
        row: String(row[rowCol]),
        number: row[numberCol],
        zone: row[zoneCol],
        price: Number(row[priceCol]) || 0,
        x: Number(row[xCol]) || 1,
        y: Number(row[yCol]) || 1,
        status: String(row[statusCol]),
        holdUntil:
          holdUntil instanceof Date
            ? holdUntil.toISOString()
            : ''
      };

    });
}
/* above codes are step to CREATE THE FUNCTION THAT READS ALL SEATS*/

function releaseExpiredHolds() {

  const lock = LockService.getScriptLock();

  try {

    lock.waitLock(10000);

    const table = getTable_(CONFIG.SHEETS.SEATS);

    const statusCol = column_(table, 'Status');
    const holdStartCol = column_(table, 'Hold Start');
    const holdUntilCol = column_(table, 'Hold Until');
    const bookingIdCol = column_(table, 'Booking ID');

    const now = new Date();

    const expiredBookingIds = [];

    for (let i = 1; i < table.values.length; i++) {

      const row = table.values[i];

      const status = String(row[statusCol]).trim();
      const holdUntil = row[holdUntilCol];

      if (status !== STATUS.ON_HOLD) {
        continue;
      }

      if (!(holdUntil instanceof Date)) {
        continue;
      }

      if (holdUntil.getTime() > now.getTime()) {
        continue;
      }

      const bookingId = String(row[bookingIdCol]).trim();

      if (bookingId) {
        expiredBookingIds.push(bookingId);
      }

      const spreadsheetRow = i + 1;

      table.sheet
        .getRange(spreadsheetRow, statusCol + 1)
        .setValue(STATUS.AVAILABLE);

      table.sheet
        .getRange(spreadsheetRow, holdStartCol + 1)
        .clearContent();

      table.sheet
        .getRange(spreadsheetRow, holdUntilCol + 1)
        .clearContent();

      table.sheet
        .getRange(spreadsheetRow, bookingIdCol + 1)
        .clearContent();
    }

    updateExpiredBookings_(expiredBookingIds);

    SpreadsheetApp.flush();

  } finally {

    if (lock.hasLock()) {
      lock.releaseLock();
    }

  }
}
function updateExpiredBookings_(bookingIds) {

  if (!bookingIds.length) {
    return;
  }

  const uniqueIds = new Set(bookingIds);

  const table = getTable_(CONFIG.SHEETS.BOOKINGS);

  const bookingIdCol = column_(table, 'Booking ID');
  const statusCol = column_(table, 'Status');

  for (let i = 1; i < table.values.length; i++) {

    const bookingId =
      String(table.values[i][bookingIdCol]).trim();

    if (!uniqueIds.has(bookingId)) {
      continue;
    }

    table.sheet
      .getRange(i + 1, statusCol + 1)
      .setValue('Expired');
  }
}
/* above codes are step to CREATE THE EXPIRED HOLD FUNCTION*/

function generateBookingId_() {

  const uuid = Utilities.getUuid()
    .replace(/-/g, '')
    .substring(0, 8)
    .toUpperCase();

  return `BK-${uuid}`;
}
/* above codes are step to CREATE THE BOOKING ID GENERATOR*/

function holdSeats(buyerName, contact, seatIds) {

  buyerName = String(buyerName || '').trim();
  contact = String(contact || '').trim();

  if (!buyerName) {
    return {
      success: false,
      message: 'Buyer name is required.'
    };
  }

  if (!contact) {
    return {
      success: false,
      message: 'Contact information is required.'
    };
  }

  if (!Array.isArray(seatIds) || seatIds.length === 0) {
    return {
      success: false,
      message: 'Please select at least one seat.'
    };
  }

  seatIds = [...new Set(
    seatIds.map(id => String(id).trim())
  )];


  const lock = LockService.getScriptLock();

  try {

    /*
     * Only one booking modification can happen
     * at a time inside this section.
     */
    lock.waitLock(10000);


    /*
     * Release old holds first.
     */
    releaseExpiredHoldsUnsafe_();


    const table = getTable_(CONFIG.SHEETS.SEATS);

    const seatIdCol = column_(table, 'Seat ID');
    const priceCol = column_(table, 'Price');
    const statusCol = column_(table, 'Status');
    const holdStartCol = column_(table, 'Hold Start');
    const holdUntilCol = column_(table, 'Hold Until');
    const bookingIdCol = column_(table, 'Booking ID');


    /*
     * Create Seat ID → row mapping
     */
    const seatRows = {};

    for (let i = 1; i < table.values.length; i++) {

      const seatId =
        String(table.values[i][seatIdCol])
          .trim()
          .toUpperCase();

      if (seatId) {
        seatRows[seatId] = i;
      }
    }


    /*
     * Check ALL seats first.
     *
     * We do this before updating anything.
     */
    for (const requestedSeatId of seatIds) {

      const key = requestedSeatId.toUpperCase();

      if (seatRows[key] === undefined) {

        return {
          success: false,
          message:
            `Seat ${requestedSeatId} does not exist.`
        };

      }

      const dataRow = seatRows[key];

      const currentStatus =
        String(
          table.values[dataRow][statusCol]
        ).trim();

      if (currentStatus !== STATUS.AVAILABLE) {

        return {
          success: false,
          message:
            `Seat ${requestedSeatId} is no longer available.`
        };

      }
    }


    /*
     * Everything is available.
     * Now create the hold.
     */

    const bookingId = generateBookingId_();

    const holdStart = new Date();

    const holdUntil = new Date(
      holdStart.getTime() +
      CONFIG.HOLD_DAYS * 24 * 60 * 60 * 1000
    );


    let totalPrice = 0;


    for (const requestedSeatId of seatIds) {

      const key = requestedSeatId.toUpperCase();

      const dataRow = seatRows[key];

      /*
       * +1 because spreadsheet rows begin at 1.
       */
      const spreadsheetRow = dataRow + 1;


      totalPrice +=
        Number(
          table.values[dataRow][priceCol]
        ) || 0;


      table.sheet
        .getRange(
          spreadsheetRow,
          statusCol + 1
        )
        .setValue(STATUS.ON_HOLD);


      table.sheet
        .getRange(
          spreadsheetRow,
          holdStartCol + 1
        )
        .setValue(holdStart);


      table.sheet
        .getRange(
          spreadsheetRow,
          holdUntilCol + 1
        )
        .setValue(holdUntil);


      table.sheet
        .getRange(
          spreadsheetRow,
          bookingIdCol + 1
        )
        .setValue(bookingId);
    }


    /*
     * Store booking information.
     */

    createBookingRecord_({
      bookingId,
      buyerName,
      contact,
      seatIds,
      totalPrice,
      holdStart,
      holdUntil
    });


    SpreadsheetApp.flush();


    return {
      success: true,
      bookingId: bookingId,
      seats: seatIds,
      total: totalPrice,
      holdUntil: holdUntil.toISOString(),
      message: 'Seats successfully placed on hold.'
    };


  } catch (error) {

    return {
      success: false,
      message: error.message
    };


  } finally {

    if (lock.hasLock()) {
      lock.releaseLock();
    }

  }
}
/* above codes are step to CREATE HOLD SEATS, CUSTOMER ABLE TO SELECT ONE OR MORE SEATS*/

function releaseExpiredHoldsUnsafe_() {

  const table = getTable_(CONFIG.SHEETS.SEATS);

  const statusCol = column_(table, 'Status');
  const holdStartCol = column_(table, 'Hold Start');
  const holdUntilCol = column_(table, 'Hold Until');
  const bookingIdCol = column_(table, 'Booking ID');

  const now = new Date();

  const expiredBookingIds = [];

  for (let i = 1; i < table.values.length; i++) {

    const row = table.values[i];

    const status =
      String(row[statusCol]).trim();

    const holdUntil =
      row[holdUntilCol];

    if (
      status === STATUS.ON_HOLD &&
      holdUntil instanceof Date &&
      holdUntil.getTime() <= now.getTime()
    ) {

      const bookingId =
        String(row[bookingIdCol]).trim();

      if (bookingId) {
        expiredBookingIds.push(bookingId);
      }


      const spreadsheetRow = i + 1;


      table.sheet
        .getRange(
          spreadsheetRow,
          statusCol + 1
        )
        .setValue(STATUS.AVAILABLE);


      table.sheet
        .getRange(
          spreadsheetRow,
          holdStartCol + 1
        )
        .clearContent();


      table.sheet
        .getRange(
          spreadsheetRow,
          holdUntilCol + 1
        )
        .clearContent();


      table.sheet
        .getRange(
          spreadsheetRow,
          bookingIdCol + 1
        )
        .clearContent();
    }
  }


  updateExpiredBookings_(expiredBookingIds);
}
/* above codes are step to ADD THE INTERNAL EXPIRY FUNCTION*/

function createBookingRecord_(booking) {

  const table =
    getTable_(CONFIG.SHEETS.BOOKINGS);

  const rowValues = {
    'BOOKING ID': booking.bookingId,
    'BUYER NAME': booking.buyerName,
    'CONTACT': booking.contact,
    'SEATS': booking.seatIds.join(', '),
    'QUANTITY': booking.seatIds.length,
    'TOTAL': booking.totalPrice,
    'STATUS': STATUS.ON_HOLD,
    'CREATED AT': booking.holdStart,
    'HOLD UNTIL': booking.holdUntil,
    'CONFIRMED AT': ''
  };


  const newRow = table.headers.map(header => {

    const key =
      String(header).trim().toUpperCase();

    return rowValues[key] ?? '';

  });


  table.sheet.appendRow(newRow);
}
/* above codes are step to write buyer information to Bookings*/

function checkA01Status() {

  const table = getTable_(CONFIG.SHEETS.SEATS);

  const seatIdCol = column_(table, 'Seat ID');
  const statusCol = column_(table, 'Status');
  const holdStartCol = column_(table, 'Hold Start');
  const holdUntilCol = column_(table, 'Hold Until');
  const bookingIdCol = column_(table, 'Booking ID');

  for (let i = 1; i < table.values.length; i++) {

    const seatId =
      String(table.values[i][seatIdCol])
        .trim()
        .toUpperCase();

    if (seatId === 'A01') {

      console.log('A01 found on spreadsheet row: ' + (i + 1));

      console.log(
        'Status = [' +
        table.values[i][statusCol] +
        ']'
      );

      console.log(
        'Hold Start = [' +
        table.values[i][holdStartCol] +
        ']'
      );

      console.log(
        'Hold Until = [' +
        table.values[i][holdUntilCol] +
        ']'
      );

      console.log(
        'Booking ID = [' +
        table.values[i][bookingIdCol] +
        ']'
      );

    }
  }
}

function testBooking() {

  const result = holdSeats(
    'Test Buyer',
    '081234567890',
    ['A01']
  );

  console.log(JSON.stringify(result, null, 2));
}
/* above codes are step to test the backend before building the webpage*/
