// ============================================================
// Database.gs — Google Sheets Read/Write Utilities
// ============================================================

var Database = (function() {

  function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getSheet(sheetName) {
    var ss    = getSpreadsheet();
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) sheet = ss.insertSheet(sheetName);
    return sheet;
  }

  function getAll(sheetName) {
    var sheet   = getSheet(sheetName);
    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    if (lastRow < 2 || lastCol < 1) return [];
    return sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  }

  function findByColumn(sheetName, colIndex, value) {
    var rows = getAll(sheetName);
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][colIndex]) === String(value)) {
        return { rowIndex: i + 2, data: rows[i] };
      }
    }
    return null;
  }

  function findAllByColumn(sheetName, colIndex, value) {
    var rows    = getAll(sheetName);
    var results = [];
    for (var i = 0; i < rows.length; i++) {
      if (String(rows[i][colIndex]) === String(value)) {
        results.push({ rowIndex: i + 2, data: rows[i] });
      }
    }
    return results;
  }

  function insert(sheetName, rowData) {
    getSheet(sheetName).appendRow(rowData);
    return true;
  }

  function updateRow(sheetName, rowIndex, rowData) {
    getSheet(sheetName).getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    return true;
  }

  function updateCell(sheetName, rowIndex, colIndex, value) {
    getSheet(sheetName).getRange(rowIndex, colIndex + 1).setValue(value);
    return true;
  }

  function deleteRow(sheetName, rowIndex) {
    getSheet(sheetName).deleteRow(rowIndex);
    return true;
  }

  function generateId(prefix) {
    return prefix + '_' + Date.now() + '_' + Math.floor(Math.random() * 9000 + 1000);
  }

  function initializeSheets() {
    var ss = getSpreadsheet();

    var defs = [
      {
        name: 'Units',
        headers: ['unit_id','tower','floor','unit_number','status','created_at','occupancy','registration','corpus_paid','corpus_amount','owner_is_tenant','lpg_mode']
      },
      {
        name: 'Owners',
        // Order MUST match OwnersService.gs C indices exactly (42 columns)
        headers: [
          'owner_id','unit_id',
          'name','email','phone','address',
          'name2','name3','phone2','phone3','email2',
          'exec_member','exec_member_title',
          'car_parking_slot','living_status','residing_country',
          'address_india','address_abroad',
          'employer','work_designation','work_address',
          'tc_number','kseb_consumer_no',
          'notes','profile_picture',
          'vehicles_4w','vehicles_2w','bank_names',
          'corpus_fund_pending','num_4w','num_2w',
          'pay_jan','pay_feb','pay_mar','pay_apr','pay_may','pay_jun',
          'pay_jul','pay_aug','pay_sep','pay_oct','pay_nov','pay_dec',
          'created_at','updated_at'
        ]
      },
      {
        name: 'Tenants',
        // Order MUST match TenantsService.gs C indices exactly (41 columns)
        headers: [
          'tenant_id','unit_id','name','name2',
          'email','phone','phone2','address','work',
          'move_in_date','move_out_date','adults','kids',
          'caution_deposit','caution_amount','id_proof','whatsapp_group',
          'wa_phone1','wa_phone2',
          'bank_name','vehicle_4w','vehicle_2w',
          'profile_picture','comment',
          'credai','num_4w','num_2w',
          'pay_jan','pay_feb','pay_mar','pay_apr','pay_may','pay_jun',
          'pay_jul','pay_aug','pay_sep','pay_oct','pay_nov','pay_dec',
          'created_at','updated_at','owner_is_tenant'
        ]
      },
      {
        name: 'Payments',
        headers: [
          'payment_id','unit_id','tenant_id','payment_type','amount',
          'month','screenshot_url','status','notes',
          'submitted_at','reviewed_at','reviewed_by'
        ]
      },
      {
        name: 'Settings',
        headers: ['key','value','updated_at']
      },
      {
        name: 'LPGReadings',
        // One row per unit per month. Amount = consumed × conv_factor × price_per_kg.
        headers: ['reading_id','unit_id','year','month','previous_reading','current_reading',
                  'reading_date','consumed','conv_factor','price_per_kg','amount',
                  'notes','recorded_by','recorded_at']
      },
      {
        name: 'LPGRates',
        // The rate in effect starting a given month — looked up by the
        // LPG service so a reading's calculation never changes later even
        // if the current market rate is updated.
        headers: ['rate_id','effective_year','effective_month','conv_factor','price_per_kg','set_by','set_at']
      },
      {
        name: 'Categories',
        headers: ['category_id','kind','name','sort_order','active','created_at']
      }
    ];

    defs.forEach(function(def) {
      var sheet = ss.getSheetByName(def.name) || ss.insertSheet(def.name);
      if (sheet.getLastRow() === 0) {
        sheet.appendRow(def.headers);
        sheet.getRange(1, 1, 1, def.headers.length)
          .setFontWeight('bold')
          .setBackground('#0f2744')
          .setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      } else {
        // Sheet exists — refresh the header row if the schema gained columns
        // OR if the header row is blank / doesn't match (e.g. after a manual clear).
        var currentCols = sheet.getLastColumn();
        var firstHeader = currentCols > 0 ? String(sheet.getRange(1, 1).getValue()) : '';
        if (currentCols < def.headers.length || firstHeader !== def.headers[0]) {
          if (sheet.getMaxColumns() < def.headers.length) {
            sheet.insertColumnsAfter(sheet.getMaxColumns(), def.headers.length - sheet.getMaxColumns());
          }
          sheet.getRange(1, 1, 1, def.headers.length).setValues([def.headers]);
          sheet.getRange(1, 1, 1, def.headers.length)
            .setFontWeight('bold')
            .setBackground('#0f2744')
            .setFontColor('#ffffff');
        }
      }
    });

    Logger.log('All sheets initialized.');
  }

  return {
    getSpreadsheet:  getSpreadsheet,
    getSheet:        getSheet,
    getAll:          getAll,
    findByColumn:    findByColumn,
    findAllByColumn: findAllByColumn,
    insert:          insert,
    updateRow:       updateRow,
    updateCell:      updateCell,
    deleteRow:       deleteRow,
    generateId:      generateId,
    initializeSheets:initializeSheets
  };
})();
