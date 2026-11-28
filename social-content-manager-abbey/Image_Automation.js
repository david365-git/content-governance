/**
 * Automatic image preparation for queued GBP posts.
 * Only processes images that have generated posts. No social publication.
 * Preserves original WordPress uploads; temporary JPGs remain until explicit release.
 */
function scm_autoPrepareQueuedImages() {
  scm_checkSetup();
  const props = PropertiesService.getScriptProperties();
  const queue = scm_rows_(SCM.queue);
  const jobs = scm_rows_('Image Processing');
  const byOpportunity = new Map(jobs.rows.map(r => [String(r.item['Opportunity ID']), r]));
  const pending = new Map();
  let reused = 0, alreadyReady = 0, prepared = 0, held = 0;
  const start = Date.now();
  const put = (row, key, value) => {
    const column = queue.keys.indexOf(key);
    if (column < 0) throw new Error('Missing queue column: ' + key);
    queue.sheet.getRange(row.number, column + 1).setValue(value);
    row.item[key] = value;
  };
  for (const row of queue.rows) {
    const p = row.item;
    if (String(p.Platform) !== 'GBP' ||
        ['PUBLISHED','PUBLISHING'].includes(String(p['Publication Status'])) ||
        ['REJECTED','PAUSED'].includes(String(p.Approval))) continue;
    const source = String(p['Original Image URL'] || '').trim();
    const existing = String(p['Prepared JPG URL'] || '').trim();
    if (!/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\//i.test(source)) {
      held++; continue;
    }
    if (/\.jpe?g(?:[?#]|$)/i.test(source)) {
      if (existing !== source) put(row, 'Prepared JPG URL', source);
      reused++; continue;
    }
    if (existing) { alreadyReady++; continue; }
    const id = String(p['Opportunity ID'] || '').trim();
    if (!id) { held++; continue; }
    if (!pending.has(id)) pending.set(id, { source: source, rows: [] });
    pending.get(id).rows.push(row);
  }
  if (!pending.size) return { status:'DONE', prepared, reused, alreadyReady, held };
  if (!props.getProperty('SCM_WP_USER') || !props.getProperty('SCM_WP_APP_PASSWORD')) {
    return { status:'WAITING FOR WORDPRESS CONNECTION', prepared, reused, alreadyReady, held:held + pending.size };
  }
  for (const [id, item] of pending) {
    if (prepared >= 8 || Date.now() - start > 180000) break;
    const holdKey = 'SCM_IMAGE_HOLD_' + id;
    const prior = Number(props.getProperty(holdKey) || 0);
    if (prior && Date.now() - prior < 24 * 3600000) { held++; continue; }
    const previous = byOpportunity.get(id);
    const job = {
      'Image Job ID': previous ? previous.item['Image Job ID'] : scm_id_('JOB'),
      'Opportunity ID': id,
      'Original URL': item.source,
      'Requested UTC': scm_now_()
    };
    try {
      const output = scm_wpRequest_('prepare', {
        url: item.source, max_width:1200, max_height:1200, quality:90
      });
      const url = String(output.jpg_url || '');
      if (!/^https:\/\/www\.abbeyfloorcare\.co\.uk\/wp-content\/uploads\/abbey-social-temp\/social-[a-f0-9]{64}\.jpg$/i.test(url) ||
          !/^[a-f0-9]{64}$/.test(String(output.token || '')))
        throw new Error('Image conversion returned an unexpected URL or token.');
      job['Prepared JPG URL'] = url;
      job['Processing Status'] = 'READY';
      job['Completed UTC'] = scm_now_();
      job['Reference Count'] = item.rows.length;
      job['Last Error'] = '';
      job['Conversion Details'] = 'WordPress token: ' + output.token +
        '; ' + output.width + 'x' + output.height;
      if (previous) scm_updateRow_(jobs, previous.number, job);
      else scm_append_('Image Processing', job);
      for (const row of item.rows) put(row, 'Prepared JPG URL', url);
      props.deleteProperty(holdKey);
      prepared++;
    } catch (err) {
      job['Processing Status'] = 'FAILED';
      job['Last Error'] = String(err.message || err).slice(0,250);
      if (previous) scm_updateRow_(jobs, previous.number, job);
      else scm_append_('Image Processing', job);
      props.setProperty(holdKey, String(Date.now()));
      held++;
    }
  }
  return {status:'DONE — NO LIVE PUBLISHING', prepared, reused, alreadyReady, held,
    remaining:Math.max(0,pending.size-prepared-held)};
}
