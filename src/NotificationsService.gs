// ============================================================
// NotificationsService.gs — WhatsApp / SMS Payment Reminders
// Uses CallMeBot (free WhatsApp API) or a Twilio-compatible SMS API
// Configure your credentials in the Settings sheet or Script Properties
// ============================================================

const NotificationsService = (() => {

  // ── Configuration ─────────────────────────────────────────

  /**
   * Get a setting from Script Properties (set via Project Settings in Apps Script).
   * Go to: Apps Script → Project Settings → Script Properties to add these:
   *   CALLMEBOT_API_KEY  — your CallMeBot API key (free, WhatsApp)
   *   TWILIO_SID         — Twilio Account SID (SMS, optional)
   *   TWILIO_TOKEN       — Twilio Auth Token
   *   TWILIO_FROM        — Twilio sender phone number
   */
  function _getProp(key) {
    return PropertiesService.getScriptProperties().getProperty(key) || '';
  }

  // ── WhatsApp via CallMeBot ─────────────────────────────────

  /**
   * Send a WhatsApp message via CallMeBot API (free tier).
   * The recipient must have added the CallMeBot contact first.
   * Phone format: international without + (e.g., 919876543210 for India).
   */
  function sendWhatsApp(phone, message) {
    const apiKey = _getProp('CALLMEBOT_API_KEY');
    if (!apiKey) {
      Logger.log('CALLMEBOT_API_KEY not configured. Skipping WhatsApp.');
      return { success: false, error: 'API key not configured.' };
    }

    const cleanPhone = String(phone).replace(/\D/g, '');
    const url = `https://api.callmebot.com/whatsapp.php?phone=${cleanPhone}&text=${encodeURIComponent(message)}&apikey=${apiKey}`;

    try {
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const code     = response.getResponseCode();
      Logger.log(`WhatsApp to ${cleanPhone}: HTTP ${code}`);
      return { success: code === 200, statusCode: code };
    } catch (err) {
      Logger.log('WhatsApp error: ' + err.message);
      return { success: false, error: err.message };
    }
  }

  // ── SMS via Twilio ─────────────────────────────────────────

  /**
   * Send an SMS via Twilio REST API.
   * Requires TWILIO_SID, TWILIO_TOKEN, TWILIO_FROM in Script Properties.
   */
  function sendSms(phone, message) {
    const sid    = _getProp('TWILIO_SID');
    const token  = _getProp('TWILIO_TOKEN');
    const from   = _getProp('TWILIO_FROM');

    if (!sid || !token || !from) {
      Logger.log('Twilio credentials not configured. Skipping SMS.');
      return { success: false, error: 'Twilio credentials not configured.' };
    }

    const cleanPhone = String(phone).replace(/[^\d+]/g, '');
    const to         = cleanPhone.startsWith('+') ? cleanPhone : '+91' + cleanPhone;
    const url        = `https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`;

    try {
      const response = UrlFetchApp.fetch(url, {
        method:             'post',
        headers:            { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token) },
        payload:            { To: to, From: from, Body: message },
        muteHttpExceptions: true
      });

      const code = response.getResponseCode();
      Logger.log(`SMS to ${to}: HTTP ${code}`);
      return { success: code === 201, statusCode: code };
    } catch (err) {
      Logger.log('SMS error: ' + err.message);
      return { success: false, error: err.message };
    }
  }

  // ── Message Templates ──────────────────────────────────────

  function _maintenanceReminderMsg(tenantName, unitId, month, amount) {
    return `*Confident Daffodils* 🌼\n\nDear ${tenantName},\n\nThis is a friendly reminder that your *Maintenance payment of ₹${amount}* for *${month}* (Unit ${unitId}) is due.\n\nKindly arrange the payment through your usual method.\n\nThank you!`;
  }

  function _wasteReminderMsg(tenantName, unitId, month) {
    return `*Confident Daffodils* 🌼\n\nDear ${tenantName},\n\nYour *Waste Management payment of ₹170* for *${month}* (Unit ${unitId}) is pending.\n\nKindly arrange the payment through your usual method.\n\nThank you!`;
  }

  function _paymentVerifiedMsg(tenantName, unitId, paymentType, amount, month) {
    return `*Confident Daffodils* 🌼\n\nDear ${tenantName},\n\nYour *${paymentType} payment of ₹${amount}* for ${month} (Unit ${unitId}) has been *verified* ✅.\n\nThank you for your payment!`;
  }

  function _paymentRejectedMsg(tenantName, unitId, paymentType, amount, month, reason) {
    return `*Confident Daffodils* 🌼\n\nDear ${tenantName},\n\nYour *${paymentType} payment of ₹${amount}* for ${month} (Unit ${unitId}) was *rejected* ❌.\n\nReason: ${reason || 'Please contact the management.'}\n\nKindly re-submit your payment.`;
  }

  // ── Notification Actions ───────────────────────────────────

  /**
   * Send payment reminder to a single tenant.
   * channel: 'whatsapp' (default) or 'sms'
   */
  function sendReminderToTenant(tenantId, paymentType, channel) {
    const tenant = TenantsService.getTenantById(tenantId);
    if (!tenant || !tenant.phone) {
      return { success: false, error: 'Tenant not found or no phone number.' };
    }

    const now    = new Date();
    const month  = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;
    const amount = paymentType === 'Maintenance' ? 2000 : paymentType === 'Waste Management' ? 170 : '(variable)';
    const msg    = paymentType === 'Maintenance'
      ? _maintenanceReminderMsg(tenant.name, tenant.unit_id, month, amount)
      : _wasteReminderMsg(tenant.name, tenant.unit_id, month);

    if (channel === 'sms') {
      return sendSms(tenant.phone, msg);
    }
    return sendWhatsApp(tenant.phone, msg);
  }

  /**
   * Send bulk reminders to all defaulters for the given month.
   * Returns { sent, failed, skipped } counts.
   */
  function sendBulkReminders(month, paymentType, channel) {
    const defaulters = TenantsService.getDefaulters(month);
    let sent = 0, failed = 0, skipped = 0;

    defaulters.forEach(d => {
      if (!d.missing_types.includes(paymentType)) { skipped++; return; }
      const tenant = d.tenant;
      if (!tenant || !tenant.phone) { skipped++; return; }

      const amount = paymentType === 'Maintenance' ? 2000 : paymentType === 'Waste Management' ? 170 : '(variable)';
      const msg    = paymentType === 'Maintenance'
        ? _maintenanceReminderMsg(tenant.name, tenant.unit_id, month, amount)
        : _wasteReminderMsg(tenant.name, tenant.unit_id, month);

      const result = channel === 'sms' ? sendSms(tenant.phone, msg) : sendWhatsApp(tenant.phone, msg);
      if (result.success) sent++;
      else failed++;

      // Respect rate limits — pause 1 second between messages
      Utilities.sleep(1000);
    });

    return { sent, failed, skipped };
  }

  /**
   * Notify a tenant that their payment was verified.
   */
  function notifyPaymentVerified(paymentId, channel) {
    const payment = PaymentsService.getPaymentById(paymentId);
    if (!payment) return { success: false, error: 'Payment not found.' };

    const tenant = TenantsService.getTenantById(payment.tenant_id);
    if (!tenant || !tenant.phone) return { success: false, error: 'Tenant has no phone.' };

    const msg = _paymentVerifiedMsg(
      tenant.name, payment.unit_id, payment.payment_type,
      payment.amount, payment.month
    );
    return channel === 'sms' ? sendSms(tenant.phone, msg) : sendWhatsApp(tenant.phone, msg);
  }

  /**
   * Notify a tenant that their payment was rejected.
   */
  function notifyPaymentRejected(paymentId, reason, channel) {
    const payment = PaymentsService.getPaymentById(paymentId);
    if (!payment) return { success: false, error: 'Payment not found.' };

    const tenant = TenantsService.getTenantById(payment.tenant_id);
    if (!tenant || !tenant.phone) return { success: false, error: 'Tenant has no phone.' };

    const msg = _paymentRejectedMsg(
      tenant.name, payment.unit_id, payment.payment_type,
      payment.amount, payment.month, reason
    );
    return channel === 'sms' ? sendSms(tenant.phone, msg) : sendWhatsApp(tenant.phone, msg);
  }

  /**
   * Set up a monthly trigger to send automatic reminders on the 5th of each month.
   * Run this once from the Apps Script editor.
   */
  function setupMonthlyReminders() {
    // Remove existing triggers to avoid duplicates
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'runMonthlyReminders') {
        ScriptApp.deleteTrigger(t);
      }
    });

    ScriptApp.newTrigger('runMonthlyReminders')
      .timeBased()
      .onMonthDay(5)
      .atHour(10)
      .create();

    Logger.log('Monthly reminder trigger set for the 5th of each month at 10 AM.');
  }

  // Public API
  return {
    sendWhatsApp,
    sendSms,
    sendReminderToTenant,
    sendBulkReminders,
    notifyPaymentVerified,
    notifyPaymentRejected,
    setupMonthlyReminders
  };

})();

// ── Trigger Handler ──────────────────────────────────────────
/**
 * Called automatically by the monthly time trigger.
 * Sends WhatsApp reminders to all defaulters for current month.
 */
function runMonthlyReminders() {
  const now   = new Date();
  const month = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}`;

  Logger.log(`Running monthly reminders for ${month}...`);

  const maintenanceResult = NotificationsService.sendBulkReminders(month, 'Maintenance', 'whatsapp');
  const wasteResult       = NotificationsService.sendBulkReminders(month, 'Waste Management', 'whatsapp');

  Logger.log(`Maintenance reminders — sent: ${maintenanceResult.sent}, failed: ${maintenanceResult.failed}`);
  Logger.log(`Waste reminders — sent: ${wasteResult.sent}, failed: ${wasteResult.failed}`);
}
