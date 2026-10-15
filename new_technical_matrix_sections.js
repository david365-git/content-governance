/**
 * Populate missing Technical Matrix sections from the ORIGINAL specification
 * (technical!C), retaining existing M:R and excluding Terrazzo.
 * Outputs are DRAFT for independent Claude audit, not approved reference.
 */
function new_technicalPendingRows() {
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('technical');
  if(!sh)throw new Error('technical tab missing');
  var data=sh.getRange(2,2,sh.getLastRow()-1,17).getValues(); // B:R
  return data.map(function(v,i){
    var material=String(v[0]||'').trim(),source=String(v[1]||'').trim();
    var missing=[];
    ['Problems, Defects & Diagnosis','Cleaning & Surface Preparation','Restoration, Repairs & Interventions','Protection, Sealing & Aftercare','Technical Restrictions & Safeguards'].forEach(function(n,k){
      if(!String(v[12+k]||'').trim())missing.push(n);
    });
    return {row:i+2,material:material,source:source,missing:missing};
  }).filter(function(x){return x.material&&x.material.toLowerCase()!=='terrazzo'&&x.source&&x.missing.length;})
    .map(function(x){return {row:x.row,material:x.material,missing:x.missing};});
}
function new_technicalFillRow(row){
  var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName('technical');
  if(!sh)throw new Error('technical tab missing');
  row=Number(row);
  if(!Number.isInteger(row)||row<2||row>sh.getLastRow())throw new Error('Invalid technical row');
  var material=String(sh.getRange(row,2).getValue()||'').trim();
  if(!material||material.toLowerCase()==='terrazzo')throw new Error('This material is protected');
  var source=String(sh.getRange(row,3).getValue()||'').trim();
  if(!source)throw new Error('No original specification in C'+row);
  var heads=sh.getRange(1,14,1,5).getDisplayValues()[0];
  var current=sh.getRange(row,14,1,5).getValues()[0];
  var missing=heads.filter(function(h,i){return !String(current[i]||'').trim();});
  if(!missing.length)return {row:row,material:material,skipped:true,cost:0,status:'Already populated'};
  var prompt=[
    'ROLE: Independent UK hard-floor technical reference editor. Use only the supplied original specification as your source.',
    'TASK: Organise documented material-specific knowledge into the requested five headings, as DRAFT technical matrix sections for subsequent independent Claude audit.',
    'Output ONLY a JSON object whose keys are EXACTLY these requested headings: '+JSON.stringify(missing)+'. Each value must be a detailed, readable plain-text technical reference.',
    'Preserve all essential technical details relevant to each heading: distinct material behaviour, practical diagnosis, cleaning chemistry and methods, appropriate repairs, sealing, aftercare and explicit restrictions.',
    'Do not generalise from other stone types, add unsubstantiated exact numbers, invent methods, or treat speculative statements as established facts.',
    'Explicitly distinguish what is safe for this exact material from other materials. Preserve cautions and conditionality in the original. Where the source contradicts itself, include a clearly labelled NEEDS AUDIT flag rather than selecting an unsupported conclusion.',
    'Do not paraphrase too aggressively or remove relevant exceptions. No citations to sites or invented external sources.',
    'Material: '+material,
    'Requested headings: '+JSON.stringify(missing),
    'Original technical specification from Column C:\n'+source
  ].join('\n\n');
  var res=bc_sendPromptViaOpenAI(prompt,10500,MODEL_MID);
  if(!res||!res.success)throw new Error('Technical extraction failed: '+(res&&res.message||'unknown'));
  var raw=String(res.text||'').trim().replace(/^```(?:json)?\s*/i,'').replace(/\s*```$/,'');
  var parsed;try{parsed=JSON.parse(raw);}catch(e){throw new Error('Invalid JSON for '+material+': '+e.message);}
  var updates=[];
  missing.forEach(function(key){
    var value=String(parsed[key]||'').trim();
    if(value.length<150||value.length>47000)throw new Error('Missing or invalid section '+key+' for '+material+' ('+value.length+' chars)');
    updates.push({key:key,value:value});
  });
  // All sections validated BEFORE any write. Never overwrite existing cells.
  updates.forEach(function(u){
    var c=heads.indexOf(u.key)+14;
    if(String(sh.getRange(row,c).getValue()||'').trim())throw new Error('Cell '+row+':'+c+' was modified; nothing overwritten');
  });
  updates.forEach(function(u){sh.getRange(row,heads.indexOf(u.key)+14).setValue(u.value);});
  SpreadsheetApp.flush();
  return {row:row,material:material,status:'DRAFT — ready for Claude audit',sections:updates.length,cost:Number(res.cost||0)};
}
