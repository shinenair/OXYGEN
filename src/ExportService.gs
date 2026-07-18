// ============================================================
// ExportService.gs — Data Export & Import Utilities
// Exports sheets as CSV; imports bulk data from CSV strings
// ============================================================

const ExportService = (() => {

  /**
   * Export all units with their owners and tenants to a CSV string.
   */
  function exportUnitsReport() {
    const units = UnitsService.getAllUnits();
    const rows  = [
      ['Unit ID', 'Tower', 'Floor', 'Status', 'Owner Names', 'Owner Phones',
       'Tenant Names', 'Tenant Phones', 'Move-in Dates']
    ];

    units.forEach(u => {
      const owners  = OwnersService.getOwnersByUnit(u.unit_id);
      const tenants = TenantsService.getTenantsByUnit(u.unit_id);

      rows.push([
        u.unit_id,
        u.tower === 'A' ? 'Tower One (A)' : 'Tower Two (B)',
        u.floor,
        u.status,
        owners.map(o => o.name).join(' | '),
        owners.map(o => o.phone).join(' | '),
        tenants.map(t => t.name).join(' | '),
        tenants.map(t => t.phone).join(' | '),
        tenants.map(t => t.move_in_date).join(' | ')
      ]);
    });

    return _arrayToCsv(rows);
  }

  /**
   * Export all owners to CSV.
   */
  function exportOwners() {
    const owners = OwnersService.getAllOwners();
    const rows   = [
      ['Unit ID', 'Name', 'Phone', 'Email', 'Address',
       'Parking Slot', 'Living Status', 'TC Number', 'KSEB Consumer No', 'Added On']
    ];

    owners.forEach(o => {
      rows.push([
        o.unit_id, o.name, o.phone, o.email, o.address,
        o.car_parking_slot, o.living_status, o.tc_number,
        o.kseb_consumer_no, _formatDateSimple(o.created_at)
      ]);
    });

    return _arrayToCsv(rows);
  }

  /**
   * Export all tenants to CSV.
   */
  function exportTenants() {
    const tenants = TenantsService.getAllTenants();
    const rows    = [
      ['Unit ID', 'Name', 'Phone', 'Email', 'Address', 'Move-in Date', 'Added On']
    ];

    tenants.forEach(t => {
      rows.push([
        t.unit_id, t.name, t.phone, t.email, t.address,
        _formatDateSimple(t.move_in_date), _formatDateSimple(t.created_at)
      ]);
    });

    return _arrayToCsv(rows);
  }

  /**
   * Export payments — optionally filtered by month and/or status.
   */
  function exportPayments(filters) {
    const payments = PaymentsService.getAllPayments(filters || null);
    const rows     = [
      ['Payment ID', 'Unit ID', 'Tenant ID', 'Type', 'Amount (₹)',
       'Month', 'Status', 'Submitted On', 'Reviewed On', 'Reviewed By', 'Notes']
    ];

    payments.forEach(p => {
      rows.push([
        p.payment_id, p.unit_id, p.tenant_id, p.payment_type,
        p.amount, p.month, p.status,
        _formatDateSimple(p.submitted_at), _formatDateSimple(p.reviewed_at),
        p.reviewed_by, p.notes
      ]);
    });

    return _arrayToCsv(rows);
  }

  /**
   * Export defaulter list for a given month to CSV.
   */
  function exportDefaulters(month) {
    const defaulters = TenantsService.getDefaulters(month);
    const rows       = [
      ['Unit ID', 'Tenant Name', 'Phone', 'Email', 'Missing Payment Types', 'Est. Outstanding (₹)']
    ];

    defaulters.forEach(d => {
      const outstanding = d.missing_types.reduce((sum, t) => {
        if (t === 'Maintenance')      return sum + 2000;
        if (t === 'Waste Management') return sum + 170;
        return sum;
      }, 0);

      rows.push([
        d.unit_id,
        d.tenant ? d.tenant.name  : '',
        d.tenant ? d.tenant.phone : '',
        d.tenant ? d.tenant.email : '',
        d.missing_types.join(', '),
        outstanding
      ]);
    });

    return _arrayToCsv(rows);
  }

  /**
   * Export monthly collection summary (all verified payments grouped by month).
   */
  function exportMonthlySummary() {
    const payments = PaymentsService.getAllPayments({ status: 'Verified' });
    const grouped  = {};

    payments.forEach(p => {
      if (!grouped[p.month]) {
        grouped[p.month] = { maintenance: 0, waste: 0, lpg: 0, total: 0, count: 0 };
      }
      const amt = Number(p.amount);
      if (p.payment_type === 'Maintenance')      grouped[p.month].maintenance += amt;
      else if (p.payment_type === 'Waste Management') grouped[p.month].waste  += amt;
      else if (p.payment_type === 'LPG')         grouped[p.month].lpg         += amt;
      grouped[p.month].total += amt;
      grouped[p.month].count++;
    });

    const rows = [
      ['Month', 'Maintenance (₹)', 'Waste Management (₹)', 'LPG (₹)', 'Total Collected (₹)', 'No. of Payments']
    ];

    Object.keys(grouped).sort().reverse().forEach(month => {
      const g = grouped[month];
      rows.push([month, g.maintenance, g.waste, g.lpg, g.total, g.count]);
    });

    return _arrayToCsv(rows);
  }

  /**
   * Save a CSV string to a Google Drive file and return its URL.
   */
  function saveCsvToDrive(csvContent, filename) {
    const folder = _getOrCreateExportsFolder();
    const blob   = Utilities.newBlob(csvContent, 'text/csv', filename);
    const file   = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  }

  /**
   * Bulk import owners from a CSV string.
   * Expected columns: unit_id, name, phone, email, living_status, car_parking_slot, tc_number, kseb_consumer_no, address
   * First row is treated as header and skipped.
   */
  function importOwnersCsv(csvString) {
    const rows    = _parseCsv(csvString);
    if (rows.length < 2) return { imported: 0, errors: ['No data rows found.'] };

    const headers = rows[0].map(h => h.trim().toLowerCase());
    const errors  = [];
    let   imported = 0;

    for (let i = 1; i < rows.length; i++) {
      const row  = rows[i];
      const data = {};
      headers.forEach((h, idx) => { data[h] = (row[idx] || '').trim(); });

      if (!data.unit_id || !data.name) {
        errors.push(`Row ${i + 1}: unit_id and name are required. Skipped.`);
        continue;
      }

      try {
        OwnersService.addOwner({
          unit_id:          data.unit_id,
          name:             data.name,
          phone:            data.phone             || '',
          email:            data.email             || '',
          living_status:    data.living_status     || 'Non-Resident',
          car_parking_slot: data.car_parking_slot  || '',
          tc_number:        data.tc_number         || '',
          kseb_consumer_no: data.kseb_consumer_no  || '',
          address:          data.address           || ''
        });
        imported++;
      } catch (err) {
        errors.push(`Row ${i + 1} (${data.unit_id}): ${err.message}`);
      }
    }

    return { imported, errors };
  }

  /**
   * Bulk import tenants from a CSV string.
   * Expected columns: unit_id, name, phone, email, move_in_date, address
   */
  function importTenantsCsv(csvString) {
    const rows    = _parseCsv(csvString);
    if (rows.length < 2) return { imported: 0, errors: ['No data rows found.'] };

    const headers = rows[0].map(h => h.trim().toLowerCase());
    const errors  = [];
    let   imported = 0;

    for (let i = 1; i < rows.length; i++) {
      const row  = rows[i];
      const data = {};
      headers.forEach((h, idx) => { data[h] = (row[idx] || '').trim(); });

      if (!data.unit_id || !data.name) {
        errors.push(`Row ${i + 1}: unit_id and name are required. Skipped.`);
        continue;
      }

      try {
        TenantsService.addTenant({
          unit_id:      data.unit_id,
          name:         data.name,
          phone:        data.phone         || '',
          email:        data.email         || '',
          move_in_date: data.move_in_date  || '',
          address:      data.address       || ''
        });
        imported++;
      } catch (err) {
        errors.push(`Row ${i + 1} (${data.unit_id}): ${err.message}`);
      }
    }

    return { imported, errors };
  }

  // ── Private helpers ──────────────────────────────────────────

  function _arrayToCsv(rows) {
    return rows.map(row =>
      row.map(cell => {
        const s = String(cell === null || cell === undefined ? '' : cell);
        // Escape cells containing commas, quotes, or newlines
        if (s.includes(',') || s.includes('"') || s.includes('\n')) {
          return '"' + s.replace(/"/g, '""') + '"';
        }
        return s;
      }).join(',')
    ).join('\n');
  }

  function _parseCsv(csvString) {
    const lines  = csvString.split('\n');
    return lines
      .filter(line => line.trim() !== '')
      .map(line => {
        const cells = [];
        let   inQuotes = false;
        let   cell     = '';
        for (let i = 0; i < line.length; i++) {
          const ch = line[i];
          if (ch === '"') {
            if (inQuotes && line[i + 1] === '"') { cell += '"'; i++; }
            else { inQuotes = !inQuotes; }
          } else if (ch === ',' && !inQuotes) {
            cells.push(cell); cell = '';
          } else {
            cell += ch;
          }
        }
        cells.push(cell);
        return cells;
      });
  }

  function _formatDateSimple(dateStr) {
    if (!dateStr) return '';
    try { return new Date(dateStr).toLocaleDateString('en-IN'); }
    catch { return String(dateStr); }
  }

  function _getOrCreateExportsFolder() {
    const folderName = 'Confident Daffodils Exports';
    const folders    = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) return folders.next();
    return DriveApp.createFolder(folderName);
  }

  // Public API
  return {
    exportUnitsReport,
    exportOwners,
    exportTenants,
    exportPayments,
    exportDefaulters,
    exportMonthlySummary,
    saveCsvToDrive,
    importOwnersCsv,
    importTenantsCsv
  };

})();

