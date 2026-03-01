/**
 * HOSCAD/EMS Tracking System - API Wrapper
 *
 * Targets the Supabase Edge Function backend at:
 *   https://vnqiqxffedudfsdoadqg.supabase.co/functions/v1/hoscad
 *
 * The anon key is intentionally public — access is controlled by the
 * custom session-token auth layer inside the Edge Function, not by RLS.
 */

const API = {
  baseUrl: 'https://vnqiqxffedudfsdoadqg.supabase.co/functions/v1/hoscad',

  // Supabase anon (publishable) key — required by the Edge Function gateway
  _apiKey: 'sb_publishable_FbP38-Tm_9iIV2QHI0Ewdw_TZfEVCJc',

  /**
   * Make an API call to the Supabase Edge Function backend.
   * @param {string} action - The API function name (e.g., 'login', 'getState')
   * @param {...any} params - Parameters to pass to the API function
   * @returns {Promise<Object>} - The API response
   */
  async call(action, ...params) {
    const body = new URLSearchParams({
      action: action,
      params: JSON.stringify(params)
    });

    try {
      // POST with apikey + Authorization headers required by Supabase Edge Runtime.
      // No redirect:follow needed — Supabase functions respond directly (no 302).
      const response = await fetch(this.baseUrl, {
        method: 'POST',
        headers: {
          'Content-Type':  'application/x-www-form-urlencoded',
          'apikey':        this._apiKey,
          'Authorization': 'Bearer ' + this._apiKey
        },
        body: body.toString()
      });
      const text = await response.text();

      try {
        return JSON.parse(text);
      } catch (parseErr) {
        console.error('API response not JSON:', text.substring(0, 200));
        return { ok: false, error: 'INVALID RESPONSE FROM SERVER. IS THE EDGE FUNCTION DEPLOYED?' };
      }
    } catch (error) {
      console.error(`API call failed: ${action}`, error);
      return { ok: false, error: 'NETWORK ERROR: ' + (error.message || 'FAILED TO FETCH. CHECK BROWSER CONSOLE.') };
    }
  },

  // ============================================================
  // Authentication
  // ============================================================

  init() {
    return this.call('init');
  },

  getPositions() {
    return this.call('GET_POSITIONS');
  },

  login(role, cadIdOrUsername, password, loginTarget, force) {
    return this.call('login', role, cadIdOrUsername, password, loginTarget || 'board', force || false);
  },

  logout(token) {
    return this.call('logout', token);
  },

  // ============================================================
  // State & Data
  // ============================================================

  getState(token, sinceTs) {
    return this.call('getState', token, sinceTs || null);
  },

  getMetrics(token, hours) {
    return this.call('getMetrics', token, hours);
  },

  getSystemStatus(token) {
    return this.call('getSystemStatus', token);
  },

  // ============================================================
  // Unit Operations
  // ============================================================

  upsertUnit(token, unitId, patch, expectedUpdatedAt) {
    return this.call('upsertUnit', token, unitId, patch, expectedUpdatedAt);
  },

  logoffUnit(token, unitId, expectedUpdatedAt) {
    return this.call('logoffUnit', token, unitId, expectedUpdatedAt);
  },

  ridoffUnit(token, unitId, expectedUpdatedAt) {
    return this.call('ridoffUnit', token, unitId, expectedUpdatedAt);
  },

  touchUnit(token, unitId, expectedUpdatedAt) {
    return this.call('touchUnit', token, unitId, expectedUpdatedAt);
  },

  touchAllOS(token) {
    return this.call('touchAllOS', token);
  },

  undoUnit(token, unitId) {
    return this.call('undoUnit', token, unitId);
  },

  getUnitInfo(token, unitId) {
    return this.call('getUnitInfo', token, unitId);
  },

  getUnitHistory(token, unitId, hours) {
    return this.call('getUnitHistory', token, unitId, hours);
  },
  getUnitIncidents(token, unitId, limit) {
    return this.call('getUnitIncidents', token, unitId, limit || 20);
  },

  massDispatch(token, destination) {
    return this.call('massDispatch', token, destination);
  },

  // ============================================================
  // Incident Operations
  // ============================================================

  createQueuedIncident(token, destination, note, priority, assignUnitId, incidentType, sceneAddress, levelOfCare) {
    return this.call('createQueuedIncident', token, destination, note, priority, assignUnitId, incidentType, sceneAddress, levelOfCare);
  },

  getIncident(token, incidentId) {
    return this.call('getIncident', token, incidentId);
  },

  updateIncident(token, incidentId, message, incidentType, destination, sceneAddress, priority, levelOfCare) {
    return this.call('updateIncident', token, incidentId, message, incidentType, destination, sceneAddress, priority, levelOfCare);
  },

  appendIncidentNote(token, incidentId, message) {
    return this.call('appendIncidentNote', token, incidentId, message);
  },

  touchIncident(token, incidentId) {
    return this.call('touchIncident', token, incidentId);
  },
  ackDispatch(token, incidentId) {
    return this.call('ackDispatch', token, incidentId);
  },

  linkUnits(token, unit1Id, unit2Id, incidentId) {
    return this.call('linkUnits', token, unit1Id, unit2Id, incidentId);
  },

  transferIncident(token, fromUnitId, toUnitId, incidentId) {
    return this.call('transferIncident', token, fromUnitId, toUnitId, incidentId);
  },

  closeIncident(token, incidentId, disposition) {
    return this.call('closeIncident', token, incidentId, disposition || '');
  },

  reopenIncident(token, incidentId) {
    return this.call('reopenIncident', token, incidentId);
  },

  requeueIncident(token, incidentId) {
    return this.call('requeueIncident', token, incidentId);
  },

  // ============================================================
  // Messaging
  // ============================================================

  sendMessage(token, toRole, message, urgent) {
    return this.call('sendMessage', token, toRole, message, urgent);
  },

  sendBroadcast(token, message, urgent) {
    return this.call('sendBroadcast', token, message, urgent);
  },
  sendToDispatchers(token, message, urgent) {
    return this.call('sendToDispatchers', token, message, urgent);
  },
  sendToUnits(token, message, urgent) {
    return this.call('sendToUnits', token, message, urgent);
  },

  getMessages(token) {
    return this.call('getMessages', token);
  },

  readMessage(token, messageId) {
    return this.call('readMessage', token, messageId);
  },

  deleteMessage(token, messageId) {
    return this.call('deleteMessage', token, messageId);
  },

  deleteAllMessages(token) {
    return this.call('deleteAllMessages', token);
  },

  // ============================================================
  // Banners
  // ============================================================

  setBanner(token, kind, message) {
    return this.call('setBanner', token, kind, message);
  },

  bannerAck(token, kind) {
    return this.call('bannerAck', token, kind);
  },

  // ============================================================
  // User Management
  // ============================================================

  newUser(token, lastName, firstName) {
    return this.call('newUser', token, lastName, firstName);
  },

  delUser(token, username) {
    return this.call('delUser', token, username);
  },

  listUsers(token) {
    return this.call('listUsers', token);
  },

  listUsersAdmin(token) {
    return this.call('listUsersAdmin', token);
  },

  changePassword(token, oldPassword, newPassword) {
    return this.call('changePassword', token, oldPassword, newPassword);
  },

  changePasswordNoAuth(username, oldPassword, newPassword) {
    return this.call('changePasswordNoAuth', username, oldPassword, newPassword);
  },

  // ============================================================
  // Session Management
  // ============================================================

  who(token, filter) {
    return this.call('who', token, filter || '');
  },

  clearSessions(token) {
    return this.call('clearSessions', token);
  },

  // ============================================================
  // Reports & Export
  // ============================================================

  reportOOS(token, hours) {
    return this.call('reportOOS', token, hours);
  },

  exportAuditCsv(token, hours) {
    return this.call('exportAuditCsv', token, hours);
  },

  // ============================================================
  // Search & Data Management
  // ============================================================

  search(token, query) {
    return this.call('search', token, query);
  },

  clearData(token, what) {
    return this.call('clearData', token, what);
  },

  // ============================================================
  // Addresses
  // ============================================================

  getAddresses(token) {
    return this.call('getAddresses', token);
  },

  addAddress(token, addr_id, name, address, city, state, zip, category, aliases, phone, notes) {
    return this.call('addAddress', token, addr_id, name, address, city, state, zip, category, aliases, phone, notes);
  },

  updateAddress(token, addr_id, name, address, city, state, zip, category, aliases, phone, notes) {
    return this.call('updateAddress', token, addr_id, name, address, city, state, zip, category, aliases, phone, notes);
  },

  removeAddress(token, addr_id) {
    return this.call('removeAddress', token, addr_id);
  },

  // ============================================================
  // Destinations (admin CRUD)
  // ============================================================

  listDestinations(token) {
    return this.call('listDestinations', token);
  },

  addDestination(token, code, name) {
    return this.call('addDestination', token, code, name);
  },

  updateDestination(token, oldCode, newCode, newName) {
    return this.call('updateDestination', token, oldCode, newCode, newName);
  },

  removeDestination(token, code) {
    return this.call('removeDestination', token, code);
  },

  handoffUnit(token, unitId, expectedUpdatedAt) {
    return this.call('handoffUnit', token, unitId, expectedUpdatedAt);
  },

  adminResetPassword(token, targetUsername, newPassword) {
    return this.call('adminResetPassword', token, targetUsername, newPassword);
  },

  // Clear a brute-force lockout for a user account (requires backend 'unlockAccount' action).
  unlockAccount(token, username) {
    return this.call('unlockAccount', token, username);
  },

  // Assign a CAD ID to a user that doesn't have one yet (admin only).
  generateCadId(token, username) {
    return this.call('generateCadId', token, username);
  },

  // Update user privilege flags (admin only).
  updateUserPrivileges(token, username, privs) {
    return this.call('updateUserPrivileges', token, username, privs);
  },

  // ============================================================
  // Crew Roster (Phase B)
  // ============================================================
  crewPinLogin(cadId, pin) { return this.call('crewPinLogin', cadId, pin); },
  changeCrewPin(cadId, currentPin, newPin) { return this.call('changeCrewPin', cadId, currentPin, newPin); },
  lookupCrewByCadId(token, cadId) { return this.call('lookupCrewByCadId', token, cadId); },
  listCrewRoster(token) { return this.call('listCrewRoster', token); },
  addCrewMember(token, fullName, certLevel) { return this.call('addCrewMember', token, fullName, certLevel); },
  updateCrewMember(token, cadId, fullName, certLevel, isActive) { return this.call('updateCrewMember', token, cadId, fullName, certLevel, isActive); },
  deleteCrewMember(token, cadId) { return this.call('deleteCrewMember', token, cadId); },
  adminSetCrewPin(token, cadId, pin) { return this.call('adminSetCrewPin', token, cadId, pin); },

  // Vehicle assignment history — who was on which unit and when (admin only).
  // All filter params optional: cadId, unitId, startIso, endIso
  getCrewReport(token, cadId, unitId, startIso, endIso) {
    return this.call('getCrewReport', token, cadId || null, unitId || null, startIso || null, endIso || null);
  },

  // ============================================================
  // Maintenance
  // ============================================================

  runPurge(token) {
    return this.call('runPurge', token);
  },

  clearPpUnits(token) {
    return this.call('clearPpUnits', token);
  },

  startDemo(token) {
    return this.call('startDemo', token);
  },

  endDemo(token) {
    return this.call('endDemo', token);
  },

  getShiftReport(token, hours, startIso, endIso) {
    return this.call('getShiftReport', token, hours, startIso, endIso);
  },

  getUnitReport(token, unitId, hours) {
    return this.call('getUnitReport', token, unitId, hours);
  },

  setDiversion(token, destCode, active) {
    return this.call('setDiversion', token, destCode, active);
  },

  saveIncTypeTaxonomy(token, taxonomyJson) {
    return this.call('saveIncTypeTaxonomy', token, taxonomyJson);
  },

  // ============================================================
  // Unit / Incident Quick Actions
  // ============================================================

  clearUnitIncident(token, unitId) {
    return this.call('clearUnitIncident', token, unitId);
  },
  setUnitETA(token, unitId, minutes) {
    return this.call('setUnitETA', token, unitId, minutes);
  },
  setUnitPAT(token, unitId, patText) {
    return this.call('setUnitPAT', token, unitId, patText);
  },
  setIncidentPriority(token, incidentId, priority) {
    return this.call('setIncidentPriority', token, incidentId, priority);
  },
  getStats(token) {
    return this.call('getStats', token);
  },

  // ============================================================
  // Viewer (read-only session)
  // ============================================================

  viewerLogin(cadId) {
    return this.call('viewerLogin', cadId);
  },

  // ============================================================
  // Unit Roster
  // ============================================================
  getRoster(token) {
    return this.call('getRoster', token);
  },
  addRosterUnit(token, unitData) {
    return this.call('addRosterUnit', token, unitData);
  },
  updateRosterUnit(token, unitId, updates) {
    return this.call('updateRosterUnit', token, unitId, updates);
  },
  deleteRosterUnit(token, unitId) {
    return this.call('deleteRosterUnit', token, unitId);
  },

  // ============================================================
  // Stacked Assignments (Phase 2D)
  // ============================================================

  assignUnit(token, incidentId, unitId) {
    return this.call('assignUnit', token, { incidentId, unitId });
  },
  selfAssign(token, incidentId, status) {
    return this.call('selfAssign', token, { incidentId, status });
  },
  queueUnit(token, incidentId, unitId) {
    return this.call('queueUnit', token, { incidentId, unitId });
  },
  primaryUnit(token, incidentId, unitId) {
    return this.call('primaryUnit', token, { incidentId, unitId });
  },
  clearUnitAssignment(token, incidentId, unitId) {
    return this.call('clearUnitAssignment', token, { incidentId, unitId });
  },
  getUnitStack(token, unitId) {
    return this.call('getUnitStack', token, { unitId });
  },

  // ============================================================
  // PulsePoint Mutual Aid Feed
  // ============================================================

  upsertPpUnits(token, ppUnits, activeUnitIds) {
    return this.call('upsertPpUnits', token, ppUnits, activeUnitIds || []);
  },

  // ============================================================
  // Scope / Dispatcher Agency Access
  // ============================================================

  setScope(token, scope) {
    return this.call('setScope', token, { scope });
  },
  getDispatcherAgencies(token) {
    return this.call('getDispatcherAgencies', token);
  },
  setDispatcherAgency(token, username, agencyId, canDispatch, canView) {
    return this.call('setDispatcherAgency', token, { username, agencyId, canDispatch: canDispatch ? 'true' : 'false', canView: canView ? 'true' : 'false' });
  },
  deleteDispatcherAgency(token, username, agencyId) {
    return this.call('deleteDispatcherAgency', token, { username, agencyId });
  },
  getAgencies(token) {
    return this.call('getAgencies', token);
  },
  updateAgency(token, agencyId, updates) {
    return this.call('updateAgency', token, agencyId, updates);
  },
  linkIncidents(token, incidentId1, incidentId2, unlink) {
    return this.call('linkIncidents', token, incidentId1, incidentId2, unlink || '');
  },

  // Mutual Aid
  requestMA(token, incidentId, agencyId, notes) {
    return this.call('requestMA', token, incidentId, agencyId, notes || '');
  },
  acknowledgeMA(token, incidentId, agencyId) {
    return this.call('acknowledgeMA', token, incidentId, agencyId);
  },
  releaseMA(token, incidentId, agencyId) {
    return this.call('releaseMA', token, incidentId, agencyId);
  },
  listMA(token, incidentId) {
    return this.call('listMA', token, incidentId);
  },

  // ============================================================
  // DC911 CadView Integration
  // ============================================================
  dc911GetConfig(token) {
    return this.call('dc911GetConfig', token);
  },
  dc911SetConfig(token, config) {
    return this.call('dc911SetConfig', token, config);
  },
  getFeatureFlags(token) {
    return this.call('getFeatureFlags', token);
  },
  setFeatureFlags(token, flags) {
    return this.call('setFeatureFlags', token, flags);
  },
  clearDc911Units(token) {
    return this.call('clearDc911Units', token);
  },

  // Issue Reports
  submitIssue(token, page, severity, description, context) {
    return this.call('submitIssue', token, page, severity, description, context || '');
  },
  listIssues(token, status) {
    return this.call('listIssues', token, status || 'OPEN');
  },
  resolveIssue(token, id, adminNote) {
    return this.call('resolveIssue', token, id, adminNote || '');
  },
  updateIssue(token, id, patch) {
    return this.call('updateIssue', token, id, patch);
  },
  getUserActivityReport(token, startIso, endIso, hours) {
    return this.call('getUserActivityReport', token, startIso || null, endIso || null, hours || 24);
  },

  // ============================================================
  // Location History
  // ============================================================
  getLocationHistory(token, address, limit, offset) {
    return this.call('getLocationHistory', token, address, limit || 25, offset || 0);
  },
  searchIncidents(token, query, limit) {
    return this.call('searchIncidents', token, query, limit || 5);
  },
  searchAddressPoints(token, query, limit) {
    return this.call('searchAddressPoints', token, query, limit || 8);
  },
  nearestAddressPoint(token, lat, lon) {
    return this.call('nearestAddressPoint', token, lat, lon);
  },
  searchIncidentsFull(token, query, limit) {
    return this.call('searchIncidentsFull', token, query, limit || 10);
  },
  createFieldIncident(token, unitId, sceneAddress, notes, incidentType) {
    return this.call('createFieldIncident', token, unitId, sceneAddress || '', notes || '', incidentType || '');
  },

  // ============================================================
  // Address Flags
  // ============================================================
  getAddressFlags(token, address) {
    return this.call('getAddressFlags', token, address);
  },
  createAddressFlag(token, address, category, description, sourceIncidentId) {
    return this.call('createAddressFlag', token, address, category, description, sourceIncidentId || '');
  },
  deactivateAddressFlag(token, flagId, reason) {
    return this.call('deactivateAddressFlag', token, flagId, reason);
  },

  // ============================================================
  // Soft Presence
  // ============================================================
  upsertPresence(token, incidentId, actionHint) {
    return this.call('upsertPresence', token, incidentId, actionHint || 'viewing');
  },
  getPresence(token, incidentId) {
    return this.call('getPresence', token, incidentId);
  },
};

// Export for module systems (if used)
if (typeof module !== 'undefined' && module.exports) {
  module.exports = API;
}
