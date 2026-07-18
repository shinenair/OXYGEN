// ============================================================
// Config.gs — Application Constants & Configuration
// ============================================================

const CONFIG = {
  APP_TITLE: 'Confident Daffodils — Property Management',
  
  // Sheet names used as "tables"
  SHEETS: {
    UNITS:    'Units',
    OWNERS:   'Owners',
    TENANTS:  'Tenants',
    PAYMENTS: 'Payments',
    SETTINGS: 'Settings'
  },

  // Tower definitions
  TOWERS: {
    A: { name: 'Tower One (A)', floors: 8, unitsPerFloor: 8, prefix: 'A' },
    B: { name: 'Tower Two (B)', floors: 9, unitsPerFloor: 8, prefix: 'B' }
  },

  // Total units: 64 (Tower A) + 72 (Tower B) = 136
  TOTAL_UNITS: 136,

  // Unit statuses
  UNIT_STATUS: {
    VACANT:   'Vacant',
    OCCUPIED: 'Occupied'
  },

  // Tenant living statuses
  LIVING_STATUS: {
    RESIDENT:     'Resident',
    NON_RESIDENT: 'Non-Resident'
  },

  // Payment types and default amounts (₹)
  PAYMENT_TYPES: {
    MAINTENANCE:      { label: 'Maintenance',      defaultAmount: 2000 },
    WASTE_MANAGEMENT: { label: 'Waste Management', defaultAmount: 170  },
    LPG:              { label: 'LPG',              defaultAmount: 0    }  // variable
  },

  // Payment statuses
  PAYMENT_STATUS: {
    PENDING:  'Pending',
    VERIFIED: 'Verified',
    REJECTED: 'Rejected'
  },


  // Column definitions for each sheet (order matters — matches spreadsheet columns)
  COLUMNS: {
    UNITS: {
      UNIT_ID:    0,   // A101, B908, etc.
      TOWER:      1,   // A or B
      FLOOR:      2,   // 1–9
      UNIT_NUMBER:3,   // 01–08
      STATUS:     4,   // Vacant / Occupied
      CREATED_AT: 5
    },
    OWNERS: {
      OWNER_ID:         0,
      UNIT_ID:          1,
      NAME:             2,
      EMAIL:            3,
      PHONE:            4,
      ADDRESS:          5,
      CAR_PARKING_SLOT: 6,
      LIVING_STATUS:    7,  // Resident / Non-Resident
      TC_NUMBER:        8,
      KSEB_CONSUMER_NO: 9,
      PROFILE_PICTURE:  10, // URL or Drive file ID
      CREATED_AT:       11,
      UPDATED_AT:       12
    },
    TENANTS: {
      TENANT_ID:   0,
      UNIT_ID:     1,
      NAME:        2,
      EMAIL:       3,
      PHONE:       4,
      ADDRESS:     5,
      MOVE_IN_DATE:6,
      CREATED_AT:  7,
      UPDATED_AT:  8
    },
    PAYMENTS: {
      PAYMENT_ID:    0,
      UNIT_ID:       1,
      TENANT_ID:     2,
      PAYMENT_TYPE:  3,  // Maintenance / Waste Management / LPG
      AMOUNT:        4,
      MONTH:         5,  // YYYY-MM
      SCREENSHOT_URL:6,
      STATUS:        7,  // Pending / Verified / Rejected
      NOTES:         8,
      SUBMITTED_AT:  9,
      REVIEWED_AT:   10,
      REVIEWED_BY:   11
    }
  }
};
