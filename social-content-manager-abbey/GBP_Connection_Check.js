/**
 * Abbey Social Content Manager — read-only GBP connection check.
 * Never creates, edits or deletes posts. Requires explicit Google API access.
 * Run scm_checkGBPConnection after setting the GBP OAuth scope and approving it.
 */
function scm_checkGBPConnection() {
  const result = { status: 'CHECKING', checks: [], postingEnabled: false };
  const props = PropertiesService.getScriptProperties();
  const fail = (stage, message) => {
    result.status = 'NEEDS SETUP';
    result.checks.push({ stage: stage, ok: false, detail: message });
  };
  const settings = scm_rows_('Settings').rows;
  const auto = settings.find(r => String(r.item.Setting) === 'Auto Publish');
  if (!auto || String(auto.item.Value).trim().toUpperCase() !== 'NO') {
    fail('SAFETY', 'Set Settings > Auto Publish to NO for connection testing.');
    return result;
  }
  result.checks.push({ stage: 'SAFETY', ok: true, detail: 'Publishing remains disabled.' });
  const token = ScriptApp.getOAuthToken();
  const headers = {
    Authorization: 'Bearer ' + token,
    'X-GOOG-API-FORMAT-VERSION': '2'
  };
  const fetchJson = url => {
    const response = UrlFetchApp.fetch(url, {
      method: 'get', headers: headers, muteHttpExceptions: true
    });
    let body = {};
    try { body = JSON.parse(response.getContentText()); } catch (_) {}
    return { code: response.getResponseCode(), body: body };
  };
  const accounts = fetchJson('https://mybusinessaccountmanagement.googleapis.com/v1/accounts?pageSize=20');
  if (accounts.code !== 200) {
    const err = accounts.body.error || {};
    fail('GOOGLE ACCESS', 'HTTP ' + accounts.code + ': ' +
      String(err.message || 'Check OAuth business.manage scope, account access and GBP API approval.').slice(0,220));
    return result;
  }
  const found = accounts.body.accounts || [];
  result.checks.push({stage:'ACCOUNTS',ok:true,detail:found.length+' accessible account(s)'});
  result.accounts = found.map(a => ({name:a.name,accountName:a.accountName || ''}));
  const configuredAccount = String(props.getProperty('SCM_GBP_ACCOUNT_ID') || '').trim();
  const configuredLocation = String(props.getProperty('SCM_GBP_LOCATION_ID') || '').trim();
  if (!configuredAccount || !configuredLocation) {
    fail('PROFILE SELECTION', 'Account and location IDs have not been configured. Set SCM_GBP_ACCOUNT_ID and SCM_GBP_LOCATION_ID after confirming the intended listing.');
    return result;
  }
  if (!/^\d+$/.test(configuredAccount) || !/^\d+$/.test(configuredLocation)) {
    fail('PROFILE SELECTION','GBP account and location IDs must contain only digits.');
    return result;
  }
  if (!found.some(a => a.name === 'accounts/' + configuredAccount)) {
    fail('PROFILE SELECTION', 'Configured account is not present in the account list for this login.');
    return result;
  }
  const localPosts = fetchJson('https://mybusiness.googleapis.com/v4/accounts/' +
    configuredAccount + '/locations/' + configuredLocation + '/localPosts?pageSize=1');
  if (localPosts.code !== 200) {
    const err = localPosts.body.error || {};
    fail('POST ACCESS', 'HTTP ' + localPosts.code + ': ' +
      String(err.message || 'Confirm this listing, Google My Business API approval and permissions.').slice(0,220));
    return result;
  }
  result.checks.push({stage:'POST ACCESS',ok:true,detail:'Read-only local posts request succeeded.'});
  result.status = 'CONNECTED — LIVE PUBLISHING STILL OFF';
  return result;
}

/** Build (but never send) the first due GBP request. */
function scm_previewGBPPayload() {
  const queue = scm_rows_('Social Posts Queue').rows;
  const row = queue.find(r => String(r.item.Platform) === 'GBP' &&
      String(r.item.Approval) === 'APPROVED' &&
      scm_validateQueuedPost_(r.item).length === 0);
  if (!row) return {status:'NO ELIGIBLE APPROVED POSTS',sent:false};
  const p = row.item;
  return {
    status:'PREVIEW ONLY',sent:false,socialPostId:String(p['Social Post ID']),
    endpoint:'https://mybusiness.googleapis.com/v4/accounts/{accountId}/locations/{locationId}/localPosts',
    body:{
      languageCode:'en-GB',
      summary:String(p['Draft Text']).trim(),
      topicType:'STANDARD',
      callToAction:{actionType:'LEARN_MORE',url:String(p['Article URL']).trim()},
      media:[{mediaFormat:'PHOTO',sourceUrl:String(p['Prepared JPG URL']).trim()}]
    }
  };
}
