// ============================================================
// Bridge.gs — Client-callable wrapper for google.script.run
// ============================================================

function doPost_internal(action, data) {
  return ApiRouter.route(action, data || {});
}
