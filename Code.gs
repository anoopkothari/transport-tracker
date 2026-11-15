/****************************************************************
 * TRANSPORT FREIGHT TRACKER — Mobile forms + Google Sheets
 *
 * VERSION: v3.14 (15-Nov-2026)
 * To check which version is running:
 *   - Sheet:  🚚 Transport menu → "Code version …" (last item)
 *   - Phone forms: shown at the top, next to the business name
 *   After pasting new code: Save, then Deploy → Manage deployments
 *   → Edit → Version: New version → Deploy (else the phone forms
 *   keep showing the old version).
 *
 * VERSION HISTORY
 *   v3.14 15-Nov-2026  Refresh reloads the whole page and stays on the
 *                      same tab and filters (browser reload too);
 *                      "on account" payments reduce a party's pending in
 *                      Summary, reminder and statement; all trips loaded
 *                      (no 1,000 limit); Summary print adds weight, rate,
 *                      notes, party details, freight-not-entered, phone
 *                      and days pending; statement adds driver, weight,
 *                      rate, party details and payments received;
 *                      weights keep decimals in PDFs
 *   v3.13 13-Nov-2026  WhatsApp reminder: your Hindi wording, commodity
 *                      added, English part mirrors the Hindi part
 *   v3.12 13-Nov-2026  WhatsApp reminder also shows the rate per quintal
 *                      (or "fixed" for fixed-amount trips)
 *   v3.11 13-Nov-2026  WhatsApp reminder: Hindi first; each trip shows
 *                      date, route, vehicle no., weight, freight, pending
 *   v3.10 12-Nov-2026  Custom Report: Transporter filter and column
 *   v3.9  12-Nov-2026  Up to 3 drivers on one trip; "+ Add new" no longer
 *                      gets stuck; saving no longer waits for the reports
 *                      to finish refreshing
 *   v3.8  12-Nov-2026  Type-to-search pickers for party, vehicle, driver,
 *                      transporter and names; Trip list filters like
 *                      Summary (by party/driver/vehicle/transporter, last
 *                      7/30 days); Print / PDF on Trip list and Summary;
 *                      missing masters created automatically from trips
 *   v3.7  12-Nov-2026  Owner's Gmail used automatically (never asked);
 *                      staff pick their name from a list (new Staff tab
 *                      + Driver master) instead of typing it
 *   v3.6  12-Nov-2026  Summary page: WhatsApp payment reminder (English +
 *                      Hindi) and party / transporter statement with Save
 *                      as PDF, for a date range and "pending above ₹";
 *                      name asked once per phone, Entered By / Updated By
 *                      saved on every record and shown in Trip list
 *   v3.5  12-Nov-2026  Vehicle numbers saved in ALL CAPS; party, driver,
 *                      transporter, commodity and place names saved as
 *                      "Sharma Traders"; menu item "Tidy capital letters
 *                      in saved entries" fixes older records
 *   v3.4  12-Nov-2026  Functions run from Apps Script no longer wait for
 *                      an OK click in the sheet (fixes "Exceeded maximum
 *                      execution time" after upgradeToMobileForms)
 *   v3.3  12-Nov-2026  Form Links tab trimmed to 2 rows (form link and
 *                      backups folder); "Show form links" rewrites it
 *   v3.2  12-Nov-2026  Summary page: totals by party, driver, vehicle or
 *                      transporter for any date range (last 7 / 30 days,
 *                      this / last month, From-To), tap through to trips;
 *                      Commodities master tab (filled from older trips,
 *                      new ones added when a trip is saved)
 *   v3.1  12-Nov-2026  Review changes: Payment trip choice shows trip no.
 *                      and date; earlier payments listed on Payment page;
 *                      Trip list hides cleared trips by default, has
 *                      This month / Last month and a From-To date range,
 *                      and an "+ Add freight" box on trips without freight
 *   v3.0  12-Nov-2026  Trip ID for every trip (T20261112-01) and payment
 *                      (P…); payments are for a chosen trip (or on
 *                      account) with an Expense column; weight, rate,
 *                      freight and advance optional; Transporter on
 *                      the main trip screen; Trip list page to view,
 *                      edit trips and add or edit payments; Trip
 *                      Summary tab; Custom Report shows paid, expense
 *                      and remaining; summaries group names ignoring
 *                      capital letters; Open backups folder menu
 *   v2.3  12-Nov-2026  Version number; "Set mobile form link…" menu
 *                      option saves the form links in Form Links tab
 *   v2.2  12-Nov-2026  "Did you mean…?" warning for similar names;
 *                      "already saved?" warning for repeat trips and
 *                      payments
 *   v2.1  12-Nov-2026  All 6 forms as mobile browser forms; "+ Add new"
 *                      party/vehicle/driver/transporter from inside
 *                      Trip and Payment; Locations master; edit an
 *                      existing master entry in place
 *   v2.0  11-Nov-2026  Mobile Trip Entry: weight in quintal, rate,
 *                      auto freight, per quintal / fixed amount,
 *                      today's date by default; Custom Report fixed
 *   v1.4  11-Nov-2026  Typing in the sheet refreshes forms and reports
 *   v1.3  11-Nov-2026  Rows typed into the sheet are counted;
 *                      "Fix forms" option
 *   v1.2  11-Nov-2026  Custom Report no longer needs the Balance column
 *   v1.1  11-Nov-2026  Setup runs in steps (fixes 6-minute timeout)
 *   v1.0  11-Nov-2026  First version with Google Forms
 *
 * 1. Create a BLANK Google Sheet.
 * 2. Extensions → Apps Script → paste this whole file.
 * 3. Fill in CONFIG below, save, choose "setup" and press Run.
 *    If the log says "Setup paused", press Run again. Repeat until
 *    it says "Setup complete". Nothing is ever created twice.
 *
 * setup() creates 6 forms (Trip Entry, Payment Received, Party,
 * Vehicle, Driver, Transporter master), links them to this sheet,
 * builds the report tabs, and installs the daily/backup triggers.
 ****************************************************************/

const CONFIG = {
  BUSINESS_NAME: 'Naman Transport',
  OWNER_EMAIL: '',                              // daily pending summary goes here (blank = your own account)
  BACKUP_EMAIL: '',                              // a SECOND Google/Gmail id: gets editor access + weekly backup file
  APP_FOLDER: 'Transport Tracker',
  BACKUP_FOLDER: 'Transport Tracker Backups',
  KEEP_BACKUPS: 60,           // number of nightly backup files kept in Drive
  TIMEZONE: 'Asia/Kolkata',
  COUNTRY_CODE: '91',         // used for WhatsApp links
  DEFAULT_CREDIT_DAYS: 30,    // used when a party has no "Credit Days" in Party Master
  REMINDER_MIN_PENDING: 1,    // email reminders only when pending ≥ this
  DOC_EXPIRY_ALERT_DAYS: 30   // insurance / permit / licence alerts
};

const CODE_VERSION = 'v3.14';   // change this line together with VERSION at the top

const CUR = '₹#,##0;-₹#,##0';
const DATE_F = 'dd-mmm-yyyy';
const PLACEHOLDER = '(add entries in the master form first)';
const SETUP_BUDGET_MS = 4 * 60 * 1000;   // Google stops scripts at 6 min; pause safely before that

const S = {
  TRIPS: 'Trips', PAYMENTS: 'Payments', PARTIES: 'Parties', VEHICLES: 'Vehicles',
  DRIVERS: 'Drivers', TRANSPORTERS: 'Transporters',
  DASH: 'Dashboard', CUSTOM: 'Custom Report', PARTY_SUM: 'Party Summary',
  VEH_SUM: 'Vehicle Summary', DRV_SUM: 'Driver Summary', MONTHLY: 'Monthly Summary',
  LINKS: 'Form Links', LOCATIONS: 'Locations', COMMODITIES: 'Commodities', STAFF: 'Staff', TRIP_SUM: 'Trip Summary'
};

/* ---------- Form definitions (question title = sheet column header) ---------- */
const FORMS = {
  PARTY: {
    key: 'PARTY_FORM_ID', sheet: S.PARTIES, title: 'Party Master', who: 'Owner / office — once per new party',
    description: 'Add a party (customer). To update details, submit again with the SAME party name and all fields filled — the latest entry is used.',
    items: [
      { t: 'Party Name', type: 'text', req: true },
      { t: 'Contact Person', type: 'text' },
      { t: 'Phone', type: 'phone', req: true, help: '10-digit mobile (used for WhatsApp reminders)' },
      { t: 'Email', type: 'email', help: 'Optional — used for email reminders' },
      { t: 'Address', type: 'para' },
      { t: 'City', type: 'text' },
      { t: 'GSTIN', type: 'text' },
      { t: 'Credit Days', type: 'number', help: 'Days allowed before payment is overdue (blank = ' + CONFIG.DEFAULT_CREDIT_DAYS + ')' },
      { t: 'Notes', type: 'para' }
    ]
  },
  VEHICLE: {
    key: 'VEHICLE_FORM_ID', sheet: S.VEHICLES, title: 'Vehicle Master', who: 'Owner / office — once per vehicle',
    description: 'Add a vehicle. To update (e.g. renewed insurance), submit again with the same vehicle number and all fields filled.',
    items: [
      { t: 'Vehicle Number', type: 'text', req: true, help: 'e.g. MH12AB1234' },
      { t: 'Vehicle Type', type: 'list', choices: ['Truck', 'Trailer', 'Tanker', 'Container', 'Tipper', 'Pickup / LCV', 'Other'] },
      { t: 'Capacity (Quintal)', type: 'number' },
      { t: 'Ownership', type: 'list', choices: ['Own', 'Attached / Market vehicle'] },
      { t: 'Owner / Transporter Name', type: 'text' },
      { t: 'Owner Phone', type: 'phone' },
      { t: 'Insurance Expiry', type: 'date' },
      { t: 'Permit Expiry', type: 'date' },
      { t: 'Fitness Expiry', type: 'date' },
      { t: 'PUC Expiry', type: 'date' },
      { t: 'Notes', type: 'para' }
    ]
  },
  DRIVER: {
    key: 'DRIVER_FORM_ID', sheet: S.DRIVERS, title: 'Driver Master', who: 'Owner / office — once per driver',
    description: 'Add a driver. To update, submit again with the same name and all fields filled.',
    items: [
      { t: 'Driver Name', type: 'text', req: true },
      { t: 'Phone', type: 'phone', req: true },
      { t: 'Licence Number', type: 'text' },
      { t: 'Licence Expiry', type: 'date' },
      { t: 'Address', type: 'para' },
      { t: 'Emergency Contact', type: 'phone' },
      { t: 'Notes', type: 'para' }
    ]
  },
  TRANSPORTER: {
    key: 'TRANSPORTER_FORM_ID', sheet: S.TRANSPORTERS, title: 'Transporter Master', who: 'Owner / office — once per transporter',
    description: 'Add a transporter / broker.',
    items: [
      { t: 'Transporter Name', type: 'text', req: true },
      { t: 'Contact Person', type: 'text' },
      { t: 'Phone', type: 'phone' },
      { t: 'Address', type: 'para' },
      { t: 'GSTIN / PAN', type: 'text' },
      { t: 'Notes', type: 'para' }
    ]
  },
  PAYMENT: {
    key: 'PAYMENT_FORM_ID', sheet: S.PAYMENTS, title: 'Payment Received', who: 'Owner / accounts — every time a party pays',
    description: 'Record money received from a party against pending balance.',
    items: [
      { t: 'Payment Date', type: 'date', req: true },
      { t: 'Party Name', type: 'dropdown', source: 'party', req: true },
      { t: 'Amount Received', type: 'number', req: true },
      { t: 'Payment Mode', type: 'list', choices: ['Cash', 'UPI', 'Bank Transfer (NEFT/RTGS/IMPS)', 'Cheque', 'Other'], req: true },
      { t: 'Reference / Trip Details', type: 'text', help: 'UTR / cheque no. / vehicle no. / bill no.' },
      { t: 'Notes', type: 'para' }
    ]
  },
  TRIP: {
    key: 'TRIP_FORM_ID', sheet: S.TRIPS, title: 'Trip Entry', who: 'Office / supervisor — one entry per trip',
    description: 'One entry per trip (load). Balance = Freight − Advance is calculated automatically.',
    items: [
      { t: 'Trip Date', type: 'date', req: true },
      { t: 'Vehicle Number', type: 'dropdown', source: 'vehicle', req: true },
      { t: 'Driver Name', type: 'dropdown', source: 'driver', req: true },
      { t: 'Source (From)', type: 'text', req: true },
      { t: 'Via Location', type: 'text', help: 'Optional — intermediate stops' },
      { t: 'Destination (To)', type: 'text', req: true },
      { t: 'Transporter Name', type: 'dropdown', source: 'transporter' },
      { t: 'Party Name', type: 'dropdown', source: 'party', req: true },
      { t: 'Party Phone', type: 'phone', help: 'Optional — leave blank to use the number in Party Master' },
      { t: 'Commodity', type: 'text' },
      { t: 'Weight (Tons)', type: 'number' },
      { t: 'Freight Amount', type: 'number', req: true, help: 'Total freight in ₹, numbers only' },
      { t: 'Advance Received', type: 'number', help: 'Leave blank if no advance' },
      { t: 'Notes / Remarks', type: 'para' }
    ]
  }
};

/* =====================================================================
 * SETUP (run until it says "Setup complete")
 * Works in steps. Each finished step is remembered, so running it
 * again continues where it stopped and never creates anything twice.
 * ===================================================================== */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SETUP_DONE')) {
    note_('Already set up', 'Setup has already finished. Use the 🚚 Transport menu in the sheet.');
    return;
  }
  const start = Date.now();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  props.setProperty('SS_ID', ss.getId());

  const steps = [
    ['Folder', () => {
      ss.setSpreadsheetTimeZone(CONFIG.TIMEZONE);
      DriveApp.getFileById(ss.getId()).moveTo(getOrCreateFolder_(CONFIG.APP_FOLDER));
    }],
    ['Party Master form', () => createForm_(FORMS.PARTY, ss)],
    ['Vehicle Master form', () => createForm_(FORMS.VEHICLE, ss)],
    ['Driver Master form', () => createForm_(FORMS.DRIVER, ss)],
    ['Transporter Master form', () => createForm_(FORMS.TRANSPORTER, ss)],
    ['Payment Received form', () => createForm_(FORMS.PAYMENT, ss)],
    ['Trip Entry form', () => createForm_(FORMS.TRIP, ss)],
    ['Balance column', addBalanceColumn_],
    ['Report tabs', () => [S.DASH, S.CUSTOM, S.TRIP_SUM, S.PARTY_SUM, S.VEH_SUM, S.DRV_SUM, S.MONTHLY].forEach(getSheet_)],
    ['Custom Report', setupCustomReport_],
    ['Form Links tab', writeFormLinks_],
    ['Form dropdowns', refreshDropdowns],
    ['Reports', buildReports],
    ['Tab order', arrangeSheets_],
    ['Daily triggers', installTriggers_],
    ['Share with backup email', shareWithBackup_],
    ['First backup', backupNow]
  ];

  for (let i = 0; i < steps.length; i++) {
    const name = steps[i][0], fn = steps[i][1];
    const flag = 'STEP_' + name;
    if (props.getProperty(flag)) continue;
    if (Date.now() - start > SETUP_BUDGET_MS) {
      note_('Setup paused', i + ' of ' + steps.length + ' steps done. Press Run (setup) again to continue — nothing will be duplicated.');
      return;
    }
    Logger.log('Step ' + (i + 1) + ' of ' + steps.length + ': ' + name);
    fn();
    props.setProperty(flag, '1');
  }

  props.setProperty('SETUP_DONE', '1');
  note_('Setup complete', 'Reload this sheet. Open the "Form Links" tab for the form links. Start by filling the master forms (parties, vehicles, drivers).');
}

function createForm_(def, ss) {
  const props = PropertiesService.getScriptProperties();
  let form;
  const existingId = props.getProperty(def.key);

  if (existingId) {
    form = FormApp.openById(existingId);              // made in an earlier run — reuse it
  } else {
    form = FormApp.create(CONFIG.BUSINESS_NAME + ' – ' + def.title)
      .setDescription(def.description)
      .setConfirmationMessage('Saved. Thank you!')
      .setShowLinkToRespondAgain(true)
      .setAllowResponseEdits(true);                   // "Edit your response" link after submitting
    def.items.forEach(s => addItem_(form, s));
    try { form.setPublished(true); } catch (e) { /* older API: forms are published by default */ }
    props.setProperty(def.key, form.getId());
    DriveApp.getFileById(form.getId()).moveTo(getOrCreateFolder_(CONFIG.APP_FOLDER));
  }

  if (ss.getSheetByName(def.sheet)) return;           // response tab already linked and renamed

  let linked = false;
  try { linked = form.getDestinationId() === ss.getId(); } catch (e) { /* no destination yet */ }
  if (!linked) form.setDestination(FormApp.DestinationType.SPREADSHEET, ss.getId());

  const sh = findResponseSheet_(ss.getId(), form.getId());
  sh.setName(def.sheet);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold');
  sh.protect().setDescription('Form data — edit only to correct mistakes').setWarningOnly(true);
}

function addItem_(form, s) {
  let item;
  switch (s.type) {
    case 'text': item = form.addTextItem(); break;
    case 'number':
      item = form.addTextItem().setValidation(FormApp.createTextValidation()
        .setHelpText('Enter a number only (no commas or ₹)').requireNumberGreaterThanOrEqualTo(0).build());
      break;
    case 'phone':
      item = form.addTextItem().setValidation(FormApp.createTextValidation()
        .setHelpText('Enter a valid phone number').requireTextMatchesPattern('^[+]?[0-9 ]{10,15}$').build());
      break;
    case 'email':
      item = form.addTextItem().setValidation(FormApp.createTextValidation()
        .setHelpText('Enter a valid email').requireTextIsEmail().build());
      break;
    case 'para': item = form.addParagraphTextItem(); break;
    case 'date': item = form.addDateItem(); break;
    case 'list': item = form.addListItem().setChoiceValues(s.choices); break;
    case 'dropdown': item = form.addListItem().setChoiceValues([PLACEHOLDER]); break;
    default: throw new Error('Unknown item type ' + s.type);
  }
  item.setTitle(s.t);
  if (s.help) item.setHelpText(s.help);
  if (s.req) item.setRequired(true);
}

// Finds the tab a form writes to. Skips tabs already set up, and checks the
// URL text first (fast) before opening any form (slow).
function findResponseSheet_(ssId, formId) {
  const ours = Object.keys(S).map(k => S[k]);
  for (let i = 0; i < 15; i++) {
    SpreadsheetApp.flush();
    const sheets = SpreadsheetApp.openById(ssId).getSheets()
      .filter(sh => ours.indexOf(sh.getName()) < 0);
    for (const sh of sheets) {
      const url = sh.getFormUrl();
      if (url && url.indexOf(formId) >= 0) return sh;
    }
    for (const sh of sheets) {
      const url = sh.getFormUrl();
      if (!url) continue;
      try { if (FormApp.openByUrl(url).getId() === formId) return sh; } catch (e) { /* keep looking */ }
    }
    Utilities.sleep(2000);
  }
  throw new Error('Could not find the response sheet for form ' + formId + '. Press Run again.');
}

function addBalanceColumn_() {
  const sh = getSheet_(S.TRIPS);
  const h = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
  if (h.indexOf('Balance') >= 0) return;
  const F = colLetter_(h.indexOf('Freight Amount') + 1);
  const A = colLetter_(h.indexOf('Advance Received') + 1);
  const c = h.length + 1;
  sh.getRange(1, c).setFormula(
    `={"Balance";ARRAYFORMULA(IF(${F}2:${F}="",,${F}2:${F}-IF(${A}2:${A}="",0,${A}2:${A})))}`);
  sh.getRange(1, c).setFontWeight('bold');
  sh.getRange(2, c, sh.getMaxRows() - 1, 1).setNumberFormat(CUR);
}

function setupCustomReport_() {
  // The Custom Report reads the Trip Summary tab, which already has Paid, Expense and Remaining per trip.
  const t = getSheet_(S.TRIP_SUM);
  if (str_(t.getRange(1, 1).getValue()) !== TRIP_SUM_HEAD[0]) {
    t.getRange(1, 1, 1, TRIP_SUM_HEAD.length).setValues([TRIP_SUM_HEAD]);
  }
  const R = title => {
    const c = colLetter_(TRIP_SUM_HEAD.indexOf(title) + 1);
    return `'${S.TRIP_SUM}'!${c}2:${c}`;
  };
  const d = R('Trip Date'), id = R('Trip ID'), v = R('Vehicle'), dr = R('Driver'), p = R('Party'), tr = R('Transporter'),
    route = R('Route'), fr = R('Freight'), ad = R('Advance'), pd = R('Paid'), ex = R('Expense'), rem = R('Remaining');

  const sh = getSheet_(S.CUSTOM);
  sh.clear();
  sh.getRange('A1:B6').setValues([['From date', ''], ['To date', ''], ['Party', 'All'], ['Vehicle', 'All'], ['Driver', 'All'],
    ['Transporter', 'All']]);
  sh.getRange('A1:A6').setFontWeight('bold');
  sh.getRange('B1:B6').setBackground('#fff8e1');
  sh.getRange('B1:B2').setNumberFormat(DATE_F)
    .setDataValidation(SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build());
  sh.getRange('D1').setValue('Change the yellow cells. Leave dates blank for all dates. Results update automatically.')
    .setFontColor('#666666');

  sh.getRange('G7').setValue('Totals').setFontWeight('bold');
  sh.getRange('H7:L7').setFormulas([['=SUM(H10:H)', '=SUM(I10:I)', '=SUM(J10:J)', '=SUM(K10:K)', '=SUM(L10:L)']])
    .setNumberFormat(CUR).setFontWeight('bold');
  sh.getRange('A9:L9').setValues([['Date', 'Trip ID', 'Vehicle', 'Driver', 'Party', 'Transporter', 'Route', 'Freight', 'Advance',
    'Paid', 'Expense', 'Remaining']])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');

  sh.getRange('A10').setFormula(
    `=IFERROR(FILTER({${d},${id},${v},${dr},${p},${tr},${route},${fr},${ad},${pd},${ex},${rem}},` +
    `${d}<>"",${d}>=IF($B$1="",DATE(2000,1,1),$B$1),${d}<=IF($B$2="",DATE(2100,1,1),$B$2),` +
    `($B$3="All")+(${p}=$B$3),($B$4="All")+(${v}=$B$4),($B$5="All")+ISNUMBER(SEARCH(", "&$B$5&",", ", "&${dr}&",")),` +
    `($B$6="All")+(${tr}=$B$6)),"No trips match these filters")`);
  sh.getRange('A10:A').setNumberFormat(DATE_F);
  sh.getRange('H10:L').setNumberFormat(CUR);
  sh.setFrozenRows(9);
  sh.setColumnWidth(1, 110);
  sh.setColumnWidth(7, 220);
}

function writeFormLinks_() {
  let backups = '';
  try { backups = getOrCreateFolder_(CONFIG.BACKUP_FOLDER).getUrl(); } catch (e) { /* ignore */ }
  const backupRow = ['Daily backups folder (one Excel file per day)', 'Owner', backups];
  if (prop_('APP_ON')) {
    const url = webAppUrl_();
    const L = f => url ? url + (f ? '?f=' + f : '') : 'Not set yet: use 🚚 Transport → Set mobile form link…';
    writeTable_(S.LINKS, ['What', 'Who uses it', 'Link'], [
      ['Mobile forms (one link: New trip, Payment, Trip list, Summary and masters are tabs at the top)', 'Everyone who enters data', L('')],
      backupRow
    ], {});
    return;
  }
  const rows = ['TRIP', 'PAYMENT', 'PARTY', 'VEHICLE', 'DRIVER', 'TRANSPORTER'].map(k => {
    const f = FormApp.openById(prop_(FORMS[k].key));
    return [FORMS[k].title, FORMS[k].who, f.getPublishedUrl(), f.getEditUrl()];
  });
  rows.push(backupRow.concat(['']));
  writeTable_(S.LINKS, ['Form', 'Who fills it', 'Link to fill (bookmark on phone)', 'Edit link (owner only)'], rows, {});
}

function arrangeSheets_() {
  const ss = ss_();
  const order = [S.DASH, S.CUSTOM, S.TRIP_SUM, S.PARTY_SUM, S.MONTHLY, S.VEH_SUM, S.DRV_SUM,
    S.TRIPS, S.PAYMENTS, S.PARTIES, S.VEHICLES, S.DRIVERS, S.TRANSPORTERS, S.LOCATIONS, S.COMMODITIES, S.STAFF, S.LINKS];
  const colors = { report: '#2e7d32', data: '#1565c0', master: '#ef6c00' };
  order.forEach((n, i) => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(i + 1);
    sh.setTabColor(i < 7 ? colors.report : i < 9 ? colors.data : colors.master);
  });
  ss.getSheets().forEach(sh => {
    if (order.indexOf(sh.getName()) < 0 && sh.getLastRow() === 0 && !sh.getFormUrl()) ss.deleteSheet(sh);
  });
  ss.setActiveSheet(ss.getSheetByName(S.LINKS));
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('handleSubmit').forSpreadsheet(ss_()).onFormSubmit().create();
  ScriptApp.newTrigger('handleEdit').forSpreadsheet(ss_()).onEdit().create();
  ScriptApp.newTrigger('dailyJob').timeBased().everyDays(1).atHour(8).inTimezone(CONFIG.TIMEZONE).create();
  ScriptApp.newTrigger('backupNow').timeBased().everyDays(1).atHour(22).inTimezone(CONFIG.TIMEZONE).create();
}

function shareWithBackup_() {
  if (!CONFIG.BACKUP_EMAIL) return;
  [CONFIG.APP_FOLDER, CONFIG.BACKUP_FOLDER].forEach(n => {
    try { getOrCreateFolder_(n).addEditor(CONFIG.BACKUP_EMAIL); }
    catch (e) { Logger.log('Could not share folder "' + n + '" with ' + CONFIG.BACKUP_EMAIL + ': ' + e); }
  });
}

/* =====================================================================
 * MENU + TRIGGERS
 * ===================================================================== */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const m = ui.createMenu('🚚 Transport');
  if (!PropertiesService.getScriptProperties().getProperty('SETUP_DONE')) {
    m.addItem('Run first-time setup', 'setup').addItem('Code version ' + CODE_VERSION, 'showVersion').addToUi();
    return;
  }
  m.addItem('Refresh reports', 'buildReports')
    .addItem('Refresh form dropdowns', 'refreshDropdowns')
    .addItem('Fix forms (publish + accept entries)', 'fixForms')
    .addSeparator()
    .addItem('Email reminders to pending parties…', 'sendPartyReminders')
    .addItem('Email me the pending summary now', 'dailyJob')
    .addSeparator()
    .addItem('Backup now', 'backupNowUi')
    .addItem('Open backups folder', 'openBackups')
    .addItem('Show form links', 'showFormLinks')
    .addItem('Set mobile form link…', 'setFormLink')
    .addItem('Rebuild Custom Report tab', 'rebuildCustomReport')
    .addItem('Tidy capital letters in saved entries', 'tidyNames')
    .addItem('Add missing masters from trips', 'addMissingMasters')
    .addSeparator()
    .addItem('Code version ' + CODE_VERSION, 'showVersion')
    .addToUi();
}

function handleSubmit(e) {
  const name = e && e.range ? e.range.getSheet().getName() : '';
  if (name === S.VEHICLES) {
    const sh = e.range.getSheet();
    const col = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('Vehicle Number') + 1;
    if (col > 0) {
      const cell = sh.getRange(e.range.getRow(), col);
      cell.setValue(normVehicle_(cell.getValue()));
    }
  }
  if ([S.PARTIES, S.VEHICLES, S.DRIVERS, S.TRANSPORTERS].indexOf(name) >= 0) refreshDropdowns();
  buildReports();
}

// Runs when you type into the sheet: master tabs refresh the form dropdowns,
// data tabs refresh the reports.
function handleEdit(e) {
  const name = e && e.range ? e.range.getSheet().getName() : '';
  try {
    if ([S.PARTIES, S.VEHICLES, S.DRIVERS, S.TRANSPORTERS].indexOf(name) >= 0) refreshDropdowns();
    if ([S.TRIPS, S.PAYMENTS, S.PARTIES, S.VEHICLES, S.DRIVERS].indexOf(name) >= 0) buildReports();
  } catch (err) {
    Logger.log('handleEdit: ' + err);
  }
}

// Run once from the editor after updating the script, to switch on the automatic updates.
function reinstallTriggers() {
  installTriggers_();
  note_('Done', 'Automatic updates are on: form entries AND typing in the sheet now refresh the forms and reports.');
}

function dailyJob() {
  buildReports();
  sendDailySummary_();
}

// Menu: puts capital letters right in everything already saved — vehicle numbers in ALL CAPS, and
// party, driver, transporter, commodity and place names as "Sharma Traders". Also removes repeated
// names from the Locations and Commodities lists. Safe to run again at any time.
function tidyNames() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let changed = 0;
  try {
    const fixCols = (sheetName, cols) => {
      const sh = ss_().getSheetByName(sheetName);
      if (!sh || sh.getLastRow() < 2) return;
      const n = sh.getLastRow(), w = sh.getLastColumn();
      const head = sh.getRange(1, 1, 1, w).getValues()[0].map(x => String(x).trim());
      const v = sh.getRange(2, 1, n - 1, w).getValues();
      Object.keys(cols).forEach(h => {
        const c = head.indexOf(h);
        if (c < 0) return;
        let any = false;
        const out = v.map(r => {
          const old = r[c];
          if (str_(old) === '') return [old];
          const nv = cols[h](old);
          if (nv !== old) { any = true; changed++; }
          return [nv];
        });
        if (any) sh.getRange(2, c + 1, n - 1, 1).setValues(out);
      });
    };
    const P = properCase_, V = normVehicle_;
    fixCols(S.TRIPS, { 'Vehicle Number': V, 'Party Name': P, 'Driver Name': P, 'Source (From)': P, 'Via Location': P,
      'Destination (To)': P, 'Commodity': P, 'Transporter Name': P });
    fixCols(S.PAYMENTS, { 'Party Name': P });
    fixCols(S.PARTIES, { 'Party Name': P });
    fixCols(S.VEHICLES, { 'Vehicle Number': V });
    fixCols(S.DRIVERS, { 'Driver Name': P });
    fixCols(S.TRANSPORTERS, { 'Transporter Name': P });
    fixCols(S.LOCATIONS, { 'Location': P });
    fixCols(S.COMMODITIES, { 'Commodity': P });
    [[S.LOCATIONS, 'Location'], [S.COMMODITIES, 'Commodity']].forEach(x => {   // drop repeated names
      const sh = ss_().getSheetByName(x[0]);
      if (!sh || sh.getLastRow() < 2) return;
      const n = sh.getLastRow(), w = sh.getLastColumn();
      const c = sh.getRange(1, 1, 1, w).getValues()[0].map(h => String(h).trim()).indexOf(x[1]);
      const seen = {}, keep = [];
      sh.getRange(2, 1, n - 1, w).getValues().forEach(r => {
        const k = key_(r[c]);
        if (k && !seen[k]) { seen[k] = 1; keep.push(r); }
      });
      if (keep.length < n - 1) {
        changed += n - 1 - keep.length;
        sh.getRange(2, 1, n - 1, w).clearContent();
        if (keep.length) sh.getRange(2, 1, keep.length, w).setValues(keep);
      }
    });
  } finally {
    lock.releaseLock();
  }
  buildReports();
  alert_('Names tidied', changed ? changed + ' cells were corrected. Reports are refreshed.' : 'Everything was already correct.');
}

function showVersion() {
  alert_('Code version', 'This sheet is running code ' + CODE_VERSION + '.\n\n' +
    'The phone forms show their version at the top. If the phone shows an older one, go to ' +
    'Apps Script → Deploy → Manage deployments → Edit → Version: New version → Deploy.');
}

function showFormLinks() { writeFormLinks_(); ss_().setActiveSheet(getSheet_(S.LINKS)); }

// Makes sure all 6 forms are published and accepting entries. Safe to run any time.
function fixForms() {
  const notes = [];
  Object.keys(FORMS).forEach(k => {
    const id = prop_(FORMS[k].key);
    if (!id) { notes.push(FORMS[k].title + ': form not found'); return; }
    const f = FormApp.openById(id);
    try { f.setPublished(true); } catch (e) { notes.push(FORMS[k].title + ' publish: ' + e.message); }
    try { f.setAcceptingResponses(true); } catch (e) { notes.push(FORMS[k].title + ' accepting: ' + e.message); }
    try { f.setAllowResponseEdits(true).setShowLinkToRespondAgain(true); } catch (e) { /* not essential */ }
  });
  refreshDropdowns();
  note_('Forms checked', notes.length ? 'Some problems:\n' + notes.join('\n') : 'All 6 forms are published, accepting entries, and dropdowns are refreshed.');
}
function rebuildCustomReport() { setupCustomReport_(); buildReports(); }

/* =====================================================================
 * MOBILE FORMS — browser forms that replace all 6 Google Forms
 * One link with a menu at the top: Trip, Payment, Party, Vehicle,
 * Driver, Transporter. Add "?f=payment" (etc.) to open one directly.
 * ===================================================================== */
const TRIP_HEADERS = ['Timestamp', 'Trip Date', 'Vehicle Number', 'Driver Name', 'Source (From)', 'Via Location',
  'Destination (To)', 'Transporter Name', 'Party Name', 'Party Phone', 'Commodity', 'Weight (Quintal)',
  'Rate (₹/Quintal)', 'Freight Type', 'Freight Amount', 'Advance Received', 'Notes / Remarks', 'Trip ID', 'Updated On', 'Entered By', 'Updated By'];
const PAY_HEADERS = ['Timestamp', 'Payment Date', 'Party Name', 'Amount Received', 'Payment Mode',
  'Reference / Trip Details', 'Notes', 'Payment ID', 'Trip ID', 'Expense', 'Updated On', 'Entered By', 'Updated By'];
const LOC_HEADERS = ['Location', 'Added On'];
const COM_HEADERS = ['Commodity', 'Added On'];
const STAFF_HEADERS = ['Name', 'Phone', 'Notes'];   // office staff who use the forms (drivers come from Driver master)
const HEADER_RENAMES = { 'Weight (Tons)': 'Weight (Quintal)', 'Capacity (Tons)': 'Capacity (Quintal)' };
const PAY_MODES = ['Cash', 'UPI', 'Bank Transfer (NEFT/RTGS/IMPS)', 'Cheque', 'Other'];

// Master forms. "quick" = the fields asked when adding a new entry from inside the Trip or Payment form.
const APP_MASTERS = {
  party: {
    sheet: S.PARTIES, label: 'Party', key: 'Party Name', quick: ['Party Name', 'Phone'],
    fields: [
      { h: 'Party Name', req: 1, full: 1, proper: 1 },
      { h: 'Phone', type: 'tel', req: 1 },
      { h: 'Contact Person' },
      { h: 'City' },
      { h: 'Credit Days', type: 'number' },
      { h: 'Email', type: 'email', more: 1 },
      { h: 'GSTIN', more: 1 },
      { h: 'Address', more: 1, full: 1, area: 1 },
      { h: 'Notes', more: 1, full: 1, area: 1 }
    ]
  },
  vehicle: {
    sheet: S.VEHICLES, label: 'Vehicle', key: 'Vehicle Number', quick: ['Vehicle Number', 'Vehicle Type'],
    fields: [
      { h: 'Vehicle Number', req: 1, upper: 1 },
      { h: 'Vehicle Type', type: 'select', choices: ['Truck', 'Trailer', 'Tanker', 'Container', 'Tipper', 'Pickup / LCV', 'Other'] },
      { h: 'Capacity (Quintal)', type: 'number' },
      { h: 'Ownership', type: 'select', choices: ['Own', 'Attached / Market vehicle'] },
      { h: 'Insurance Expiry', type: 'date', sec: 'Document expiry' },
      { h: 'Permit Expiry', type: 'date' },
      { h: 'Fitness Expiry', type: 'date' },
      { h: 'PUC Expiry', type: 'date' },
      { h: 'Owner / Transporter Name', more: 1 },
      { h: 'Owner Phone', type: 'tel', more: 1 },
      { h: 'Notes', more: 1, full: 1, area: 1 }
    ]
  },
  driver: {
    sheet: S.DRIVERS, label: 'Driver', key: 'Driver Name', quick: ['Driver Name', 'Phone'],
    fields: [
      { h: 'Driver Name', req: 1, proper: 1 },
      { h: 'Phone', type: 'tel', req: 1 },
      { h: 'Licence Number' },
      { h: 'Licence Expiry', type: 'date' },
      { h: 'Address', more: 1, full: 1, area: 1 },
      { h: 'Emergency Contact', type: 'tel', more: 1 },
      { h: 'Notes', more: 1, full: 1, area: 1 }
    ]
  },
  transporter: {
    sheet: S.TRANSPORTERS, label: 'Transporter', key: 'Transporter Name', quick: ['Transporter Name', 'Phone'],
    fields: [
      { h: 'Transporter Name', req: 1, full: 1, proper: 1 },
      { h: 'Contact Person' },
      { h: 'Phone', type: 'tel' },
      { h: 'GSTIN / PAN', more: 1 },
      { h: 'Address', more: 1, full: 1, area: 1 },
      { h: 'Notes', more: 1, full: 1, area: 1 }
    ]
  }
};

// Run ONCE after deploying the web app: switches all 6 forms to the mobile forms.
function upgradeToMobileForms() {
  const props = PropertiesService.getScriptProperties();
  Object.keys(FORMS).forEach(k => {
    const id = prop_(FORMS[k].key);
    if (!id) return;
    try {
      const f = FormApp.openById(id);
      try { f.removeDestination(); } catch (e) { /* already unlinked */ }
      try {
        f.setAcceptingResponses(false)
          .setCustomClosedFormMessage('This form has moved to the new mobile forms. Ask the office for the new link.');
      } catch (e) { /* not essential */ }
    } catch (e) {
      Logger.log('Could not close ' + FORMS[k].title + ': ' + e);
    }
  });
  props.setProperty('APP_ON', '1');
  props.setProperty('TRIP_APP_ON', '1');

  ensureHeaders_(getSheet_(S.TRIPS), TRIP_HEADERS);
  ensureHeaders_(getSheet_(S.PAYMENTS), PAY_HEADERS);
  Object.keys(APP_MASTERS).forEach(k =>
    ensureHeaders_(getSheet_(APP_MASTERS[k].sheet), ['Timestamp'].concat(APP_MASTERS[k].fields.map(f => f.h))));
  ensureHeaders_(getSheet_(S.LOCATIONS), LOC_HEADERS);
  ensureHeaders_(getSheet_(S.COMMODITIES), COM_HEADERS);
  ensureHeaders_(getSheet_(S.STAFF), STAFF_HEADERS);
  setupCustomReport_();
  writeFormLinks_();
  buildReports();

  const url = webAppUrl_();
  note_('Mobile forms are on', url
    ? 'Open this link on the phone (it has all 6 forms):\n' + url
    : 'Done. Now reload the sheet and use 🚚 Transport → Set mobile form link… to save the link.');
}

// Kept so the earlier instructions still work.
function upgradeTripEntry() { upgradeToMobileForms(); }

// The link saved with "Set mobile form link…" is used first. Google's own guess is used only if it is a
// real /exec link (a /dev link works only for the owner, not for staff).
function webAppUrl_() {
  const saved = prop_('WEB_APP_URL');
  if (saved) return saved;
  try {
    const u = ScriptApp.getService().getUrl() || '';
    return /\/exec$/.test(u) ? u : '';
  } catch (e) { return ''; }
}

// Menu: paste the Web app link once; the Form Links tab is rewritten with a link for every form.
function setFormLink() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Mobile form link',
    'Paste the Web app URL (Deploy → Manage deployments → copy the URL that ends in /exec):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  const url = str_(r.getResponseText()).replace(/[?#].*$/, '').replace(/\/+$/, '');
  if (!/^https:\/\/script\.google\.com\/.+\/exec$/.test(url)) {
    ui.alert('That link does not look right. It must start with https://script.google.com/ and end with /exec.');
    return;
  }
  PropertiesService.getScriptProperties().setProperty('WEB_APP_URL', url);
  writeFormLinks_();
  ss_().setActiveSheet(getSheet_(S.LINKS));
  ui.alert('Saved. The Form Links tab now has the link for every form.');
}

// Makes sure every wanted column exists (adds missing ones at the end) and renames old tons headers.
function ensureHeaders_(sh, wanted) {
  const last = sh.getLastColumn();
  let head = last ? sh.getRange(1, 1, 1, last).getValues()[0].map(x => String(x).trim()) : [];
  head.forEach((h, i) => {
    const nh = HEADER_RENAMES[h];
    if (nh && head.indexOf(nh) < 0) { sh.getRange(1, i + 1).setValue(nh); head[i] = nh; }
  });
  const missing = wanted.filter(h => head.indexOf(h) < 0);
  if (missing.length) {
    const need = head.length + missing.length;
    if (sh.getMaxColumns() < need) sh.insertColumnsAfter(sh.getMaxColumns(), need - sh.getMaxColumns());
    sh.getRange(1, head.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    head = head.concat(missing);
  }
  if (!sh.getFrozenRows()) sh.setFrozenRows(1);
  return head;
}

// First empty row after the last row that has data (ignores the Balance formula column).
function nextRow_(sh, width, skipCol) {
  const n = sh.getLastRow();
  if (n < 2) return 2;
  const v = sh.getRange(1, 1, n, width).getValues();
  for (let i = n - 1; i >= 1; i--) {
    if (v[i].some((x, c) => c !== skipCol && x !== '' && x !== null)) return i + 2;
  }
  return 2;
}

// Writes one row, skipping the Balance formula column so its formula keeps working.
function writeRow_(sh, r, row, skipCol) {
  if (r > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 100);
  const seg = (a, b) => { if (b > a) sh.getRange(r, a + 1, 1, b - a).setValues([row.slice(a, b)]); };
  if (skipCol < 0) seg(0, row.length); else { seg(0, skipCol); seg(skipCol + 1, row.length); }
}

function parseYmd_(s) {
  const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(str_(s));
  return m ? new Date(+m[1], +m[2] - 1, +m[3]) : '';
}

function ymd_(v) {
  const d = toDate_(v);
  return d ? Utilities.formatDate(d, CONFIG.TIMEZONE, 'yyyy-MM-dd') : '';
}

// "sharma  TRADERS" → "Sharma Traders"; "r.k. sharma" → "R.K. Sharma". Used for party, driver,
// transporter, commodity and place names. Vehicle numbers use normVehicle_ (ALL CAPS, no spaces).
// Who saved a record: the name typed once on the phone, plus the Google email when Google gives it
// (with the form open to "Anyone", that is only for the sheet owner).
function who_(d) {
  const raw = str_(d && d.by);
  const name = raw && raw.indexOf('@') < 0 ? properCase_(raw) : '';
  return userEmail_() || name || raw;
}

// The Google email of the person using the form. With the form open to "Anyone", Google gives it only
// for the sheet owner; for everyone else it is blank and the name chosen on the phone is used.
function userEmail_() {
  try { return Session.getActiveUser().getEmail() || ''; } catch (e) { return ''; }
}

// One trip can have up to 3 drivers, saved together in "Driver Name" as "Ramesh, Suresh".
function drivers_(v) {
  const seen = {}, out = [];
  (Array.isArray(v) ? v : [v]).forEach(x => str_(x).split(',').forEach(n => {
    const p = properCase_(n), k = key_(p);
    if (k && !seen[k]) { seen[k] = 1; out.push(p); }
  }));
  return out;
}

function properCase_(v) {
  const s = str_(v).replace(/\s+/g, ' ').toLowerCase();
  if (s === 'self / own vehicle') return 'Self / Own vehicle';
  return s.replace(/(^|[\s.(-])([a-z])/g, (m, p, c) => p + c.toUpperCase())
    .replace(/ (And|Of|The) /g, (m, w) => ' ' + w.toLowerCase() + ' ');   // "Gupta and Sons"
}

function cellVal_(f, v) {
  if (f.proper) return properCase_(v);
  if (f.type === 'date') return parseYmd_(v);
  if (f.type === 'number') return str_(v) === '' ? '' : num_(v);
  if (f.upper) return normVehicle_(v);
  return str_(v);
}

function recordOut_(def, r) {
  const o = {};
  def.fields.forEach(f => {
    const v = r[f.h];
    o[f.h] = f.type === 'date' ? ymd_(v) : f.upper ? normVehicle_(v) : str_(v);
  });
  return o;
}

function masterNorm_(type) { return type === 'vehicle' ? normVehicle_ : key_; }

function appData_(start) {
  const rec = {};
  Object.keys(APP_MASTERS).forEach(k => {
    const def = APP_MASTERS[k], norm = masterNorm_(k);
    rec[k] = latestBy_(readTable_(def.sheet), r => norm(r[def.key]))
      .map(r => recordOut_(def, r))
      .filter(r => r[def.key])
      .sort((a, b) => a[def.key].localeCompare(b[def.key]));
  });
  const trips = readTable_(S.TRIPS);
  const locs = readTable_(S.LOCATIONS).map(r => r['Location'])
    .concat(trips.map(r => r['Source (From)']), trips.map(r => r['Destination (To)']), trips.map(r => r['Via Location']));
  const live = appTrips_();
  return {
    biz: CONFIG.BUSINESS_NAME,
    ver: CODE_VERSION,
    today: Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd'),
    start: start,
    defs: APP_MASTERS,
    rec: rec,
    trips: live.trips,
    pending: live.pending,
    onacc: live.onacc,
    locations: uniq_(locs),
    commodities: uniq_(readTable_(S.COMMODITIES).map(r => r['Commodity']).concat(trips.map(r => r['Commodity']))),
    modes: PAY_MODES,
    cc: CONFIG.COUNTRY_CODE,
    email: userEmail_(),
    staff: uniq_(readTable_(S.STAFF).map(r => properCase_(r['Name']))
      .concat(readTable_(S.DRIVERS).map(r => properCase_(r['Driver Name']))))
  };
}

function doGet(e) {
  const tabs = ['trip', 'payment', 'trips', 'summary', 'party', 'vehicle', 'driver', 'transporter'];
  const f = e && e.parameter && tabs.indexOf(e.parameter.f) >= 0 ? e.parameter.f : 'trip';
  const data = appData_(f);
  data.qs = cleanQs_(e && e.parameter);
  try { data.url = ScriptApp.getService().getUrl(); } catch (err) { data.url = ''; }
  const html = APP_HTML.replace('__DATA__', JSON.stringify(data).replace(/</g, '\\u003c'));
  return HtmlService.createHtmlOutput(html)
    .setTitle(CONFIG.BUSINESS_NAME + ' – Entry')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Tab state from the page address (?f=summary&by=party&from=...), so Refresh and a browser
// reload come back to the same tab with the same filters.
function cleanQs_(p) {
  const out = {};
  ['by', 'which', 'from', 'to', 'min', 'lf', 'lby', 'lwhich', 'lfrom', 'lto', 'q'].forEach(k => {
    if (p && Object.prototype.hasOwnProperty.call(p, k)) out[k] = String(p[k]).slice(0, 100);
  });
  return out;
}

// Called by the forms right after a save, so the save itself stays quick.
function refreshAfterSave() {
  try { buildReports(); } catch (e) { Logger.log('Reports not refreshed: ' + e); }
}

// Save or update a master entry (party, vehicle, driver, transporter).
// An existing name or vehicle number is updated in place; a new one is added.
function saveMaster(type, d) {
  const def = APP_MASTERS[type];
  if (!def) throw new Error('Unknown form.');
  d = d || {};
  const quick = !!d._quick;
  def.fields.forEach(f => {
    if (f.req && (!quick || def.quick.indexOf(f.h) >= 0) && !str_(d[f.h])) throw new Error(f.h + ' is required.');
  });
  const norm = masterNorm_(type);
  let name = type === 'vehicle' ? normVehicle_(d[def.key]) : properCase_(d[def.key]);
  let updated = false, out = null;

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(def.sheet);
    const head = ensureHeaders_(sh, ['Timestamp'].concat(def.fields.map(f => f.h), ['Entered By', 'Updated By']));
    const kc = head.indexOf(def.key);
    const n = sh.getLastRow();
    let r = 0, row = null;
    if (n >= 2) {
      const v = sh.getRange(2, 1, n - 1, head.length).getValues();
      for (let i = v.length - 1; i >= 0; i--) {
        if (norm(v[i][kc]) === norm(name)) { r = i + 2; row = v[i]; break; }
      }
    }
    if (row) {
      updated = true;
    } else {
      row = head.map(() => '');
      r = nextRow_(sh, head.length, -1);
    }
    row[head.indexOf('Timestamp')] = new Date();
    row[head.indexOf(updated ? 'Updated By' : 'Entered By')] = who_(d);
    def.fields.forEach(f => {
      if (f.h === def.key) row[kc] = name;
      else if (Object.prototype.hasOwnProperty.call(d, f.h)) row[head.indexOf(f.h)] = cellVal_(f, d[f.h]);
    });
    writeRow_(sh, r, row, -1);
    def.fields.filter(f => f.type === 'date').forEach(f => sh.getRange(r, head.indexOf(f.h) + 1).setNumberFormat(DATE_F));
    const obj = {};
    head.forEach((h, i) => { obj[h] = row[i]; });
    out = recordOut_(def, obj);
  } finally {
    lock.releaseLock();
  }
  return { name: name, updated: updated, record: out };
}

// Looks for an already-saved row with the same key details. Returns a short description, or ''.
function findDup_(sh, head, same, describe) {
  const n = sh.getLastRow();
  if (n < 2) return '';
  const v = sh.getRange(2, 1, n - 1, head.length).getValues();
  for (let i = v.length - 1; i >= 0; i--) {
    const o = {};
    head.forEach((h, j) => { o[h] = v[i][j]; });
    if (same(o)) return describe(o);
  }
  return '';
}

// Saves a new payment, or updates one when d.id is given. A payment is for one trip (d.tripId),
// or "on account" for the party when no trip is chosen. Expense reduces the balance like a payment.
function savePayment(d) {
  d = d || {};
  if (!d.date || !str_(d.party)) throw new Error('Fill all required fields.');
  const amt = str_(d.amount) === '' ? 0 : num_(d.amount);
  const exp = str_(d.expense) === '' ? 0 : num_(d.expense);
  if (amt < 0 || exp < 0) throw new Error('Amounts cannot be negative.');
  if (!(amt > 0 || exp > 0)) throw new Error('Enter the amount received or an expense.');
  if (amt > 0 && !str_(d.mode)) throw new Error('Choose the payment mode.');
  const tripId = str_(d.tripId);
  let id = str_(d.id);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(S.PAYMENTS);
    const head = ensureHeaders_(sh, PAY_HEADERS);
    const vals = {
      'Payment Date': parseYmd_(d.date),
      'Party Name': properCase_(d.party),
      'Trip ID': tripId,
      'Amount Received': amt > 0 ? amt : '',
      'Expense': exp > 0 ? exp : '',
      'Payment Mode': amt > 0 ? str_(d.mode) : '',
      'Reference / Trip Details': str_(d.ref),
      'Notes': str_(d.notes)
    };
    let r, row;
    if (id) {
      r = findRowById_(sh, head, 'Payment ID', id);
      if (!r) throw new Error('Payment ' + id + ' was not found. Reload the page and try again.');
      row = sh.getRange(r, 1, 1, head.length).getValues()[0];
      vals['Updated On'] = new Date();
      vals['Updated By'] = who_(d);
    } else {
      if (!d.force) {
        const dup = findDup_(sh, head,
          o => ymd_(o['Payment Date']) === ymd_(vals['Payment Date']) && key_(o['Party Name']) === key_(vals['Party Name']) &&
            str_(o['Trip ID']) === tripId && num_(o['Amount Received']) === amt && num_(o['Expense']) === exp,
          o => fmtD_(toDate_(o['Payment Date'])) + ', ' + str_(o['Party Name']) +
            (str_(o['Trip ID']) ? ', trip ' + str_(o['Trip ID']) : '') +
            (num_(o['Amount Received']) ? ', ₹' + inr_(num_(o['Amount Received'])) + ' ' + str_(o['Payment Mode']) : '') +
            (num_(o['Expense']) ? ', expense ₹' + inr_(num_(o['Expense'])) : ''));
        if (dup) return { dup: true, info: dup };
      }
      id = newId_('P', vals['Payment Date'] || new Date(), usedIds_(sh, head, 'Payment ID'));
      vals['Payment ID'] = id;
      vals['Timestamp'] = new Date();
      vals['Entered By'] = who_(d);
      r = nextRow_(sh, head.length, -1);
      row = head.map(() => '');
    }
    head.forEach((h, i) => { if (h in vals) row[i] = vals[h]; });
    writeRow_(sh, r, row, -1);
    sh.getRange(r, head.indexOf('Payment Date') + 1).setNumberFormat(DATE_F);
    sh.getRange(r, head.indexOf('Amount Received') + 1).setNumberFormat(CUR);
    sh.getRange(r, head.indexOf('Expense') + 1).setNumberFormat(CUR);
  } finally {
    lock.releaseLock();
  }
  return { saved: true, id: id };
}

// Saves a new trip, or updates one when d.id is given. Only date, vehicle, driver, party, From and To
// are required; weight, rate, freight and advance can be added later.
function saveTrip(d) {
  d = d || {};
  if (!d.date || !str_(d.vehicle) || !str_(d.driver) || !str_(d.party) || !str_(d.from) || !str_(d.to)) {
    throw new Error('Fill all required fields.');
  }
  const fixed = d.type === 'Fixed amount';
  let freight = '';
  if (fixed) { if (str_(d.fixed) !== '') freight = num_(d.fixed); }
  else if (str_(d.weight) !== '' && str_(d.rate) !== '') freight = Math.round(num_(d.weight) * num_(d.rate) * 100) / 100;
  let id = str_(d.id);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_(S.TRIPS);
    const head = ensureHeaders_(sh, TRIP_HEADERS);
    const vals = {
      'Trip Date': parseYmd_(d.date),
      'Vehicle Number': normVehicle_(d.vehicle),
      'Driver Name': drivers_([d.driver, d.driver2, d.driver3]).join(', '),
      'Source (From)': properCase_(d.from),
      'Via Location': properCase_(d.via),
      'Destination (To)': properCase_(d.to),
      'Transporter Name': properCase_(d.transporter),
      'Party Name': properCase_(d.party),
      'Party Phone': str_(d.phone),
      'Commodity': properCase_(d.commodity),
      'Weight (Quintal)': str_(d.weight) === '' ? '' : num_(d.weight),
      'Rate (₹/Quintal)': fixed || str_(d.rate) === '' ? '' : num_(d.rate),
      'Freight Type': fixed ? 'Fixed amount' : 'Per quintal',
      'Freight Amount': freight,
      'Advance Received': str_(d.advance) === '' ? '' : num_(d.advance),
      'Notes / Remarks': str_(d.notes)
    };
    const bal = head.indexOf('Balance');
    let r, row;
    if (id) {
      r = findRowById_(sh, head, 'Trip ID', id);
      if (!r) throw new Error('Trip ' + id + ' was not found. Reload the page and try again.');
      row = sh.getRange(r, 1, 1, head.length).getValues()[0];
      vals['Updated On'] = new Date();
      vals['Updated By'] = who_(d);
    } else {
      if (!d.force) {
        const dup = findDup_(sh, head,
          o => ymd_(o['Trip Date']) === ymd_(vals['Trip Date']) && normVehicle_(o['Vehicle Number']) === vals['Vehicle Number'] &&
            key_(o['Party Name']) === key_(vals['Party Name']) && key_(o['Source (From)']) === key_(vals['Source (From)']) &&
            key_(o['Destination (To)']) === key_(vals['Destination (To)']),
          o => (str_(o['Trip ID']) ? str_(o['Trip ID']) + ', ' : '') + fmtD_(toDate_(o['Trip Date'])) + ', ' +
            normVehicle_(o['Vehicle Number']) + ', ' + str_(o['Source (From)']) + ' → ' + str_(o['Destination (To)']) + ', ' +
            str_(o['Party Name']) + (str_(o['Freight Amount']) !== '' ? ', freight ₹' + inr_(num_(o['Freight Amount'])) : ''));
        if (dup) return { dup: true, info: dup };
      }
      id = newId_('T', vals['Trip Date'] || new Date(), usedIds_(sh, head, 'Trip ID'));
      vals['Trip ID'] = id;
      vals['Timestamp'] = new Date();
      vals['Entered By'] = who_(d);
      r = nextRow_(sh, head.length, bal);
      row = head.map(() => '');
    }
    head.forEach((h, i) => { if (h in vals) row[i] = vals[h]; });
    writeRow_(sh, r, row, bal);
    sh.getRange(r, head.indexOf('Trip Date') + 1).setNumberFormat(DATE_F);
    sh.getRange(r, head.indexOf('Freight Amount') + 1).setNumberFormat(CUR);
    addLocations_([d.from, d.via, d.to]);
    addToList_(S.COMMODITIES, COM_HEADERS, 'Commodity', [d.commodity]);
  } finally {
    lock.releaseLock();
  }
  return { saved: true, id: id };
}

function findRowById_(sh, head, idCol, id) {
  const n = sh.getLastRow(), c = head.indexOf(idCol);
  if (n < 2 || c < 0) return 0;
  const v = sh.getRange(2, c + 1, n - 1, 1).getValues();
  for (let i = 0; i < v.length; i++) if (str_(v[i][0]).toUpperCase() === str_(id).toUpperCase()) return i + 2;
  return 0;
}

function usedIds_(sh, head, idCol) {
  const used = {}, n = sh.getLastRow(), c = head.indexOf(idCol);
  if (n >= 2 && c >= 0) sh.getRange(2, c + 1, n - 1, 1).getValues().forEach(r => { if (str_(r[0])) used[str_(r[0])] = 1; });
  return used;
}

// Date-based ID: T20261112-01, T20261112-02 … (P… for payments). The number restarts each day.
function newId_(prefix, date, used) {
  const base = prefix + Utilities.formatDate(date, CONFIG.TIMEZONE, 'yyyyMMdd') + '-';
  let n = 1;
  while (used[base + String(n).padStart(2, '0')]) n++;
  return base + String(n).padStart(2, '0');
}

// Gives an ID to every trip and payment that does not have one yet (older rows, rows typed in the sheet).
function ensureIds_() {
  idsFor_(S.TRIPS, TRIP_HEADERS, 'Trip ID', 'T', 'Trip Date');
  idsFor_(S.PAYMENTS, PAY_HEADERS, 'Payment ID', 'P', 'Payment Date');
}

function idsFor_(name, headers, idCol, prefix, dateCol) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return;
  const head = ensureHeaders_(sh, headers);
  const n = sh.getLastRow();
  const v = sh.getRange(2, 1, n - 1, head.length).getValues();
  const ic = head.indexOf(idCol), dc = head.indexOf(dateCol), tc = head.indexOf('Timestamp'), bal = head.indexOf('Balance');
  const used = {};
  v.forEach(r => { if (str_(r[ic])) used[str_(r[ic])] = 1; });
  v.forEach((r, i) => {
    if (str_(r[ic])) return;
    if (!r.some((x, c) => c !== bal && c !== ic && x !== '' && x !== null)) return;
    const id = newId_(prefix, toDate_(r[dc]) || toDate_(r[tc]) || new Date(), used);
    used[id] = 1;
    sh.getRange(i + 2, ic + 1).setValue(id);
  });
}

// Reads Trips and Payments and works out, for every trip, what was paid, the expenses and what is left.
function tripData_() {
  ensureIds_();
  const trips = readTable_(S.TRIPS).map(r => {
    const fr = r['Freight Amount'], wt = r['Weight (Quintal)'] !== undefined ? r['Weight (Quintal)'] : r['Weight (Tons)'];
    const t = {
      id: str_(r['Trip ID']),
      date: toDate_(r['Trip Date']) || toDate_(r['Timestamp']),
      vehicle: normVehicle_(r['Vehicle Number']),
      driver: drivers_(r['Driver Name']).join(', '),
      party: str_(r['Party Name']),
      transporter: str_(r['Transporter Name']),
      from: str_(r['Source (From)']),
      via: str_(r['Via Location']),
      to: str_(r['Destination (To)']),
      phone: str_(r['Party Phone']),
      commodity: str_(r['Commodity']),
      weightRaw: str_(wt) === '' ? '' : num_(wt),
      rateRaw: str_(r['Rate (₹/Quintal)']) === '' ? '' : num_(r['Rate (₹/Quintal)']),
      ftype: str_(r['Freight Type']) || 'Per quintal',
      freightRaw: str_(fr) === '' ? '' : num_(fr),
      advanceRaw: str_(r['Advance Received']) === '' ? '' : num_(r['Advance Received']),
      notes: str_(r['Notes / Remarks']),
      enteredBy: str_(r['Entered By']), updatedBy: str_(r['Updated By']),
      created: toDate_(r['Timestamp']), changed: toDate_(r['Updated On']),
      paid: 0, expense: 0, pays: []
    };
    t.weight = num_(t.weightRaw);
    t.freight = num_(t.freightRaw);
    t.advance = num_(t.advanceRaw);
    t.balance = t.freight - t.advance;
    return t;
  });
  const byId = {};
  trips.forEach(t => { if (t.id) byId[t.id.toUpperCase()] = t; });
  const pays = readTable_(S.PAYMENTS).map(r => ({
    id: str_(r['Payment ID']),
    tripId: str_(r['Trip ID']),
    date: toDate_(r['Payment Date']) || toDate_(r['Timestamp']),
    party: str_(r['Party Name']),
    amount: num_(r['Amount Received']),
    expense: num_(r['Expense']),
    mode: str_(r['Payment Mode']),
    ref: str_(r['Reference / Trip Details']),
    notes: str_(r['Notes'])
  }));
  pays.forEach(p => {
    const t = p.tripId ? byId[p.tripId.toUpperCase()] : null;
    p.linked = !!t;
    if (t) { t.paid += p.amount; t.expense += p.expense; t.pays.push(p); }
  });
  trips.forEach(t => {
    t.remaining = t.balance - t.paid - t.expense;
    t.status = t.freightRaw === '' || t.freight <= 0 ? 'Freight not entered' : t.remaining <= 0 ? 'Cleared' : 'Due';
  });
  return { trips: trips, pays: pays };
}

// Pending per party (lower-case key): all trip balances minus every payment and expense.
function pendingMap_(data) {
  const m = {};
  data.trips.forEach(t => { if (t.party) m[key_(t.party)] = (m[key_(t.party)] || 0) + t.balance; });
  data.pays.forEach(p => { if (p.party) m[key_(p.party)] = (m[key_(p.party)] || 0) - p.amount - p.expense; });
  return m;
}

// All trips for the phone forms (newest first), pending per party, and payments that are not
// linked to a trip ("on account"), which reduce the party's pending in Summary and statements.
function appTrips_() {
  const data = tripData_();
  const out = data.trips.slice()
    .sort((a, b) => (b.date || 0) - (a.date || 0) || (b.id < a.id ? -1 : 1))
    .map(t => ({
      id: t.id, date: t.date ? ymd_(t.date) : '', vehicle: t.vehicle, driver: t.driver, party: t.party,
      transporter: t.transporter, from: t.from, via: t.via, to: t.to, phone: t.phone, commodity: t.commodity,
      enteredBy: t.enteredBy, updatedBy: t.updatedBy,
      created: t.created ? Utilities.formatDate(t.created, CONFIG.TIMEZONE, 'd MMM HH:mm') : '',
      changed: t.changed ? Utilities.formatDate(t.changed, CONFIG.TIMEZONE, 'd MMM HH:mm') : '',
      weight: t.weightRaw, rate: t.rateRaw, ftype: t.ftype, freight: t.freightRaw, advance: t.advanceRaw, notes: t.notes,
      paid: t.paid, expense: t.expense, remaining: t.remaining, status: t.status,
      pays: t.pays.map(p => ({ id: p.id, date: p.date ? ymd_(p.date) : '', amount: p.amount, expense: p.expense,
        mode: p.mode, ref: p.ref, notes: p.notes, tripId: p.tripId, party: p.party }))
    }));
  const onacc = data.pays.filter(p => !p.linked && p.party).map(p => ({
    id: p.id, date: p.date ? ymd_(p.date) : '', party: p.party, amount: p.amount, expense: p.expense,
    mode: p.mode, ref: p.ref }));
  return { trips: out, pending: pendingMap_(data), onacc: onacc };
}

// Party / transporter statement as a PDF in Drive (folder "Party Statements" inside the app folder).
// Only trips with money still due are listed. spec = { by: 'party'|'transporter', name, from, to, min }.
// Adds to the master tabs every party, driver, vehicle and transporter that appears in Trips but is
// missing from its master (for example trips typed into the sheet, or older entries). Only the name is
// filled (plus the party phone from the trip, if any); Entered By says "Added from trips".
function syncMasters_(trips) {
  const specs = [
    ['party', t => t.party, t => ({ 'Phone': t.phone })],
    ['driver', t => drivers_(t.driver), () => ({})],
    ['vehicle', t => t.vehicle, () => ({})],
    ['transporter', t => (key_(t.transporter) === 'self / own vehicle' ? '' : t.transporter), () => ({})]
  ];
  const newestFirst = trips.slice().sort((a, b) => (b.date || 0) - (a.date || 0));
  let added = 0;
  specs.forEach(sp => {
    const type = sp[0], getName = sp[1], extra = sp[2];
    const def = APP_MASTERS[type], norm = masterNorm_(type);
    const have = {};
    readTable_(def.sheet).forEach(r => { const k = norm(r[def.key]); if (k) have[k] = 1; });
    const fresh = [];
    newestFirst.forEach(t => {
      [].concat(getName(t)).forEach(nm => {
        const k = norm(nm);
        if (!k || have[k]) return;
        have[k] = 1;
        const vals = Object.assign({ 'Timestamp': new Date(), 'Entered By': 'Added from trips' }, extra(t));
        vals[def.key] = type === 'vehicle' ? normVehicle_(nm) : properCase_(nm);
        fresh.push(vals);
      });
    });
    if (!fresh.length) return;
    const sh = getSheet_(def.sheet);
    const head = ensureHeaders_(sh, ['Timestamp'].concat(def.fields.map(f => f.h), ['Entered By', 'Updated By']));
    const r = nextRow_(sh, head.length, -1);
    if (r + fresh.length > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), fresh.length + 50);
    sh.getRange(r, 1, fresh.length, head.length).setValues(fresh.map(v => head.map(h => (h in v ? v[h] : ''))));
    added += fresh.length;
  });
  return added;
}

// Menu: same check on demand, with a message saying how many were added.
function addMissingMasters() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let n = 0;
  try { n = syncMasters_(tripData_().trips); } finally { lock.releaseLock(); }
  buildReports();
  alert_('Masters checked', n
    ? n + (n === 1 ? ' missing entry was' : ' missing entries were') + ' added to the Party, Driver, Vehicle and Transporter tabs. Add their phone numbers and other details there or in the forms.'
    : 'Nothing was missing. Every party, driver, vehicle and transporter in your trips is already in its master.');
}

function subFolder_(name) {
  const app = getOrCreateFolder_(CONFIG.APP_FOLDER);
  const it = app.getFoldersByName(name);
  return it.hasNext() ? it.next() : app.createFolder(name);
}

// Print / PDF from the phone forms: makes a PDF of the table on screen (Trip list or Summary) in
// Drive → Transport Tracker → Reports. spec = { title, subtitle, head: [...], rows: [[...]], foot: [...],
// numCols: [...], info: [[label, value]] (party details at the top), lines: [[label, amount]] (totals
// under the table), notes: ['...'] }.
function makeReportPdf(spec) {
  spec = spec || {};
  const head = (spec.head || []).map(x => str_(x));
  const rows = (spec.rows || []).slice(0, 5000);
  const foot = spec.foot || null, numCols = spec.numCols || [];
  if (!rows.length) throw new Error('Nothing to print for these filters.');
  const font = 'font-family:Arial,sans-serif';
  const td = i => 'style="border:1px solid #ccc;padding:3px 4px;font-size:9px;' + font + (numCols.indexOf(i) >= 0 ? ';text-align:right' : '') + '"';
  const cell = (v, i) => '<td ' + td(i) + '>' + esc_(typeof v === 'number' ? fmtNum_(v) : str_(v)) + '</td>';
  const info = (spec.info || []).filter(x => x && str_(x[1]));
  const lines = spec.lines || [], notes = (spec.notes || []).filter(x => str_(x));
  const html = '<div style="' + font + ';font-size:9px;color:#111">' +
    '<h2 style="margin:0;font-size:15px">' + esc_(CONFIG.BUSINESS_NAME) + ' — ' + esc_(spec.title) + '</h2>' +
    '<p style="margin:4px 0 6px">' + esc_(spec.subtitle) + '<br>Printed: ' +
    esc_(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm')) + '</p>' +
    (info.length ? '<p style="margin:0 0 8px;font-size:10px">' +
      info.map(x => '<b>' + esc_(x[0]) + ':</b> ' + esc_(x[1])).join(' &nbsp;·&nbsp; ') + '</p>' : '') +
    '<table style="border-collapse:collapse;width:100%;font-size:9px"><tr style="background:#1f4e78;color:#fff">' +
    head.map((h, i) => '<th ' + td(i) + '>' + esc_(h) + '</th>').join('') + '</tr>' +
    rows.map(r => '<tr>' + (r || []).map(cell).join('') + '</tr>').join('') +
    (foot ? '<tr style="font-weight:bold;background:#f1f3f6">' + foot.map(cell).join('') + '</tr>' : '') + '</table>' +
    (lines.length ? '<table style="border-collapse:collapse;margin-top:8px;font-size:10px">' +
      lines.map((x, i) => '<tr' + (i === lines.length - 1 ? ' style="font-weight:bold"' : '') + '><td style="padding:2px 12px 2px 0;' + font + '">' +
        esc_(x[0]) + '</td><td style="padding:2px 0;text-align:right;' + font + '">Rs. ' + fmtNum_(num_(x[1])) + '</td></tr>').join('') + '</table>' : '') +
    notes.map(n => '<p style="margin:6px 0 0;color:#555">' + esc_(n) + '</p>').join('') + '</div>';
  const fname = str_(spec.title).replace(/[^\w ]+/g, '').replace(/\s+/g, '_') + '_' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HHmm') + '.pdf';
  const file = subFolder_('Reports').createFile(HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(fname));
  return { url: file.getUrl(), name: fname };
}

// Party / transporter details for the top of a PDF, from the latest master entry.
function masterInfo_(by, name) {
  const def = APP_MASTERS[by];
  if (!def) return [];
  const k = key_(name);
  const r = latestBy_(readTable_(def.sheet), x => key_(x[def.key])).filter(x => key_(x[def.key]) === k)[0];
  if (!r) return [];
  const fields = by === 'party' ? ['Contact Person', 'Phone', 'Address', 'City', 'GSTIN', 'Credit Days']
    : ['Contact Person', 'Phone', 'Address', 'GSTIN / PAN'];
  return fields.filter(h => str_(r[h])).map(h => [h, str_(r[h])]);
}

function makeStatementPdf(spec) {
  spec = spec || {};
  const by = spec.by === 'transporter' ? 'transporter' : 'party';
  const k = key_(spec.name), min = num_(spec.min);
  const inPeriod = d => (!spec.from || d >= spec.from) && (!spec.to || d <= spec.to);
  const data = tripData_();
  const rows = data.trips.filter(t => {
    const nm = by === 'party' ? t.party : (t.transporter || 'Self / Own vehicle');
    return key_(nm) === k && t.status === 'Due' && t.remaining >= min && inPeriod(t.date ? ymd_(t.date) : '');
  }).sort((a, b) => (a.date || 0) - (b.date || 0));
  if (!rows.length) throw new Error('No pending trips for these filters.');
  const name = by === 'party' ? rows[0].party : (rows[0].transporter || 'Self / Own vehicle');
  const period = (spec.from ? fmtD_(parseYmd_(spec.from)) : 'Start') + ' to ' + (spec.to ? fmtD_(parseYmd_(spec.to)) : fmtD_(new Date()));
  const f = 'font-family:Arial,sans-serif', b = 'border:1px solid #ccc;padding:4px 5px;font-size:10px;' + f;
  const td = 'style="' + b + '"', tdr = 'style="' + b + ';text-align:right"';
  const th = h => '<th ' + td + '>' + h + '</th>';
  const total = rows.reduce((s, t) => s + t.remaining, 0);

  // Payments received: those against the listed trips, plus "on account" payments of this party in the period.
  const onacc = by === 'party' ? data.pays.filter(p => !p.linked && key_(p.party) === k && inPeriod(p.date ? ymd_(p.date) : '')) : [];
  const oa = onacc.reduce((s, p) => s + p.amount + p.expense, 0);
  const pays = [];
  rows.forEach(t => t.pays.forEach(p => pays.push({ p: p, against: t.id })));
  onacc.forEach(p => pays.push({ p: p, against: 'On account' }));
  pays.sort((x, y) => (x.p.date || 0) - (y.p.date || 0));

  const info = [[by === 'party' ? 'Party' : 'Transporter', name]].concat(masterInfo_(by, name));
  const rate = t => t.ftype === 'Fixed amount' ? 'Fixed' : (t.rateRaw === '' ? '' : fmtNum_(t.rateRaw));
  const html = '<div style="' + f + ';font-size:11px;color:#111">' +
    '<h2 style="margin:0">' + esc_(CONFIG.BUSINESS_NAME) + ' — Statement</h2>' +
    '<p style="margin:4px 0 12px">' + info.map(x => esc_(x[0]) + ': <b>' + esc_(x[1]) + '</b>').join('<br>') +
    '<br>Period: ' + esc_(period) + '<br>Date: ' + esc_(fmtD_(new Date())) + '</p>' +
    '<table style="border-collapse:collapse;width:100%;font-size:10px"><tr style="background:#1f4e78;color:#fff">' +
    ['Date', 'Trip no.', 'Route', 'Vehicle', 'Driver', 'Weight (qtl)', 'Rate (Rs./qtl)', 'Freight', 'Advance', 'Paid', 'Expense', 'Pending'].map(th).join('') + '</tr>' +
    rows.map(t => '<tr><td ' + td + '>' + esc_(fmtD_(t.date)) + '</td><td ' + td + '>' + esc_(t.id) + '</td><td ' + td + '>' +
      esc_(t.from + ' → ' + t.to) + '</td><td ' + td + '>' + esc_(t.vehicle) + '</td><td ' + td + '>' + esc_(t.driver) + '</td>' +
      '<td ' + tdr + '>' + (t.weightRaw === '' ? '' : fmtNum_(t.weightRaw)) + '</td><td ' + tdr + '>' + esc_(rate(t)) + '</td>' +
      [t.freight, t.advance, t.paid, t.expense].map(v => '<td ' + tdr + '>' + fmtNum_(v) + '</td>').join('') +
      '<td ' + tdr + '><b>' + fmtNum_(t.remaining) + '</b></td></tr>').join('') +
    '<tr><td colspan="11" ' + tdr + '><b>Total pending on these trips (Rs.)</b></td><td ' + tdr + '><b>' + fmtNum_(total) + '</b></td></tr>' +
    (oa ? '<tr><td colspan="11" ' + tdr + '>Less: payments received on account (Rs.)</td><td ' + tdr + '>-' + fmtNum_(oa) + '</td></tr>' +
      '<tr><td colspan="11" ' + tdr + '><b>Net pending (Rs.)</b></td><td ' + tdr + '><b>' + fmtNum_(total - oa) + '</b></td></tr>' : '') +
    '</table>' +
    (pays.length ? '<h3 style="margin:14px 0 4px;font-size:12px">Payments received</h3>' +
      '<table style="border-collapse:collapse;width:100%;font-size:10px"><tr style="background:#1f4e78;color:#fff">' +
      ['Date', 'Payment no.', 'Against', 'Mode', 'Reference', 'Amount', 'Expense'].map(th).join('') + '</tr>' +
      pays.map(x => '<tr><td ' + td + '>' + esc_(x.p.date ? fmtD_(x.p.date) : '') + '</td><td ' + td + '>' + esc_(x.p.id) + '</td><td ' + td + '>' +
        esc_(x.against) + '</td><td ' + td + '>' + esc_(x.p.mode) + '</td><td ' + td + '>' + esc_(x.p.ref) + '</td><td ' + tdr + '>' +
        fmtNum_(x.p.amount) + '</td><td ' + tdr + '>' + fmtNum_(x.p.expense) + '</td></tr>').join('') + '</table>' : '') +
    '<p style="color:#555;margin-top:10px">Only trips with payment pending are shown. Amounts in Rupees.</p></div>';
  const app = getOrCreateFolder_(CONFIG.APP_FOLDER);
  const it = app.getFoldersByName('Party Statements');
  const folder = it.hasNext() ? it.next() : app.createFolder('Party Statements');
  const fname = 'Statement_' + name.replace(/[^\w ]+/g, '').replace(/\s+/g, '_') + '_' +
    Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd_HHmm') + '.pdf';
  const file = folder.createFile(HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf').setName(fname));
  return { url: file.getUrl(), name: fname };
}

// Called by the phone forms to refresh the trip list after a save.
function getAppTrips() { return appTrips_(); }

// New From / Via / To places typed in a trip are added to the Locations master.
function addLocations_(names) { addToList_(S.LOCATIONS, LOC_HEADERS, 'Location', names); }

// Adds names that are not in a simple list tab yet (Locations, Commodities). Capital letters are ignored.
function addToList_(sheetName, headers, col, names) {
  const sh = getSheet_(sheetName);
  const head = ensureHeaders_(sh, headers);
  const have = {};
  readTable_(sheetName).forEach(r => { have[key_(r[col])] = 1; });
  const add = [];
  names.forEach(nm => {
    const s = properCase_(nm), k = key_(s);
    if (s && !have[k]) { have[k] = 1; add.push(s); }
  });
  add.forEach(s => {
    const vals = { 'Added On': new Date() };
    vals[col] = s;
    writeRow_(sh, nextRow_(sh, head.length, -1), head.map(h => (h in vals ? vals[h] : '')), -1);
  });
}

const APP_HTML = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
*{box-sizing:border-box}body{margin:0;font-family:system-ui,-apple-system,Roboto,sans-serif;background:#f1f3f6;color:#111}
.w{max-width:480px;margin:0 auto;padding:0 8px 16px}
.top{position:sticky;top:0;background:#f1f3f6;padding:8px 0 6px;z-index:5}
.biz{font-size:12px;color:#666;margin:0 2px 6px}
.tabs{display:grid;grid-template-columns:repeat(4,1fr);gap:6px}
.tabs2{grid-template-columns:repeat(4,1fr);margin-top:6px}
.tabs2 button{height:30px;font-size:13px}
.tabs button{height:34px;padding:0 4px;border:1px solid #c8ccd2;border-radius:17px;background:#fff;font-size:14px;color:#333}
.tabs .on{background:#1f4e78;color:#fff;border-color:#1f4e78}
.c{background:#fff;border-radius:12px;padding:12px;display:flex;flex-direction:column;gap:10px}
h1{font-size:18px;margin:0;font-weight:600}
.sub{font-size:12px;color:#666;margin-top:2px}
.g{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.full{grid-column:1/-1}
label{display:block;font-size:12px;color:#555;margin-bottom:3px}
input,select,textarea{width:100%;height:42px;font-size:16px;padding:6px 8px;border:1px solid #c8ccd2;border-radius:8px;background:#fff;color:#111}
textarea{height:64px}
.req:after{content:" *";color:#c62828}
.seg{display:flex;border:1px solid #c8ccd2;border-radius:8px;overflow:hidden}
.seg button{flex:1;height:38px;border:0;background:#fff;font-size:14px;color:#444}
.seg .on{background:#1f4e78;color:#fff}
.chips{display:flex;flex-wrap:wrap;gap:6px}
.chips button{height:38px;padding:0 14px;border:1px solid #c8ccd2;border-radius:8px;background:#fff;font-size:14px;color:#333}
.chips .on{background:#1f4e78;color:#fff;border-color:#1f4e78}
.out{height:42px;padding:9px 8px;border-radius:8px;background:#e6eef8;color:#1f4e78;font-weight:600;font-size:16px}
.bal{display:flex;justify-content:space-between;font-size:15px}
.more{color:#1f4e78;font-size:14px;background:none;border:0;text-align:left;padding:2px 0}
.save{height:50px;border:0;border-radius:10px;background:#1f4e78;color:#fff;font-size:17px;font-weight:600;width:100%}
.save:disabled{background:#7a96b3}
.btn2{height:50px;border:1px solid #c8ccd2;border-radius:10px;background:#fff;font-size:16px;color:#333}
.row{display:flex;gap:8px}.row .save{flex:2}.row .btn2{flex:1}
.msg{font-size:14px;padding:10px;border-radius:8px;display:none;margin-bottom:8px}
.ok{background:#d4edda;color:#155724}.err{background:#f8d7da;color:#721c24}
.bad{border-color:#c62828;background:#fff5f5}
.info{font-size:13px;padding:8px 10px;border-radius:8px;background:#e6eef8;color:#1f4e78}
.note{font-size:13px;padding:8px 10px;border-radius:8px;background:#fff3cd;color:#7a5a00}
.sect{grid-column:1/-1;font-size:12px;color:#888;border-top:1px solid #eee;padding-top:6px;margin-top:2px}
.qa{border:1.5px dashed #1f4e78;border-radius:10px;padding:10px;background:#f5f9fd;display:flex;flex-direction:column;gap:8px}
.qa h2{font-size:15px;margin:0;font-weight:600;color:#1f4e78}
.qa .row button{height:44px}
.view{display:none}
.pick{display:flex;flex-wrap:wrap;gap:6px;margin-top:8px}
.pick button{min-height:40px;padding:4px 12px;border:1px solid #1f4e78;border-radius:8px;background:#1f4e78;color:#fff;font-size:14px}
.pick .alt{background:#fff;color:#333;border-color:#c8ccd2}
.info{white-space:pre-line}
.hint{font-size:12px;color:#777;margin-top:-4px}
.prev{font-size:13px;color:#333;white-space:pre-line;border:1px solid #e1e4e8;border-radius:8px;padding:8px 10px;line-height:1.5}
.hd2{display:flex;justify-content:space-between;align-items:center}
.totals{font-size:13px;color:#444}
.trip{border:1px solid #e1e4e8;border-radius:10px;padding:10px;margin-bottom:8px;background:#fff;cursor:pointer;font-size:14px;display:flex;flex-direction:column;gap:3px}
.trip .l1{display:flex;justify-content:space-between;align-items:center;gap:8px}
.pill{font-size:12px;padding:2px 8px;border-radius:10px;white-space:nowrap}
.pill.due{background:#fff3cd;color:#7a5a00}.pill.ok{background:#d4edda;color:#155724}.pill.inc{background:#e2e3e5;color:#383d41}
.amts{font-size:13px;color:#444}.rem{font-weight:600;color:#1f4e78}
.det{border-top:1px solid #eee;margin-top:6px;padding-top:8px}
.det .h3{font-size:13px;font-weight:600;margin:4px 0}
.prow{display:flex;justify-content:space-between;align-items:center;gap:8px;font-size:13px;padding:4px 0;border-bottom:1px dashed #eee}
.mini{height:32px;padding:0 12px;border:1px solid #c8ccd2;border-radius:8px;background:#fff;font-size:13px}
.fq{border:1.5px dashed #1f4e78;border-radius:10px;padding:10px;background:#f5f9fd;display:flex;flex-direction:column;gap:8px;margin-top:6px;cursor:default}
.fq .row button{height:44px;font-size:15px}
.addf{align-self:flex-start;margin-top:4px;border-color:#1f4e78;color:#1f4e78;font-weight:600}
.kpis{display:grid;grid-template-columns:1fr 1fr;gap:8px}
.kpi{background:#f1f3f6;border-radius:8px;padding:8px 10px}
.kpi span{display:block;font-size:12px;color:#555}.kpi b{font-size:17px}
.kpi.main{background:#e6eef8;color:#1f4e78}
.srow{display:flex;justify-content:space-between;align-items:center;gap:8px;padding:9px 2px;border-bottom:1px solid #eee;cursor:pointer;font-size:14px}
.srow .r{text-align:right;white-space:nowrap}.srow small{display:block;font-size:12px;color:#666}
.bizrow{display:flex;justify-content:space-between;align-items:baseline;gap:8px}
.bizrow .biz{margin-bottom:6px}
.whoTxt{font-size:12px;color:#555;white-space:nowrap}.whoTxt a{color:#1f4e78;text-decoration:underline;cursor:pointer}
.nameBox{background:#fff3cd;border-radius:10px;padding:10px;margin-bottom:8px}
.nameBox .row{margin-top:4px}.nameBox select{flex:1}.nameBox .row .save{flex:0 0 90px;height:42px;font-size:16px}
.who{font-size:12px;color:#555;margin-top:6px}
.wa{background:#1f8b4c !important;color:#fff !important;text-decoration:none;display:flex;align-items:center;justify-content:center;text-align:center}
.wa.off{background:#9bb8a6 !important;pointer-events:none}
.minrow{display:flex;align-items:center;gap:8px}.minrow label{flex:1;margin:0}.minrow input{flex:0 0 110px}
.st{border:1px solid #e1e4e8;border-radius:10px;padding:10px;display:flex;flex-direction:column;gap:6px}
.st .sl{display:flex;justify-content:space-between;gap:8px;font-size:13px;border-bottom:1px solid #eee;padding:6px 0}
.st .sl small{display:block;color:#666;font-size:12px}
.st .tot{display:flex;justify-content:space-between;font-weight:700;font-size:15px;padding-top:4px}
.combo{position:relative}
.comboList{position:absolute;left:0;right:0;top:calc(100% + 2px);z-index:30;background:#fff;border:1px solid #c8ccd2;border-radius:8px;max-height:280px;overflow-y:auto;box-shadow:0 8px 20px rgba(0,0,0,.15)}
.comboList div{padding:11px 10px;font-size:15px;border-bottom:1px solid #f0f2f5;cursor:pointer;color:#111}
.comboList div.on{background:#e6eef8}
.comboList .add{color:#1f4e78;font-weight:600}
.comboList .none{color:#888;cursor:default}
.nameBox .combo{flex:1}
.hbtns{display:flex;gap:12px;align-items:center}
.dl{display:flex;justify-content:space-between;align-items:baseline}.dl .more{font-size:13px;padding:0}
.det .row{margin-top:8px}.det .row button{height:44px;font-size:15px}
</style></head><body><div class="w">
<div class="top"><div class="bizrow"><div class="biz" id="biz"></div><div class="whoTxt" id="whoTxt"></div></div>
<div class="nameBox" id="nameBox" style="display:none"><label>Who are you? (asked once on this phone)</label>
<div class="row"><select id="whoSel"></select><input id="whoIn" autocomplete="name" style="display:none"><button type="button" class="save" id="whoSave">Save</button></div>
<div class="sub" id="whoHint"></div></div><div class="tabs" id="tabs"></div><div class="tabs tabs2" id="tabs2"></div></div>
<div class="msg" id="msg"></div>

<div class="view" id="v_trip"><div class="c">
<div><h1 id="t_title">New trip</h1><div class="sub" id="t_sub"></div></div>
<div class="g">
<div><label class="req">Date</label><input type="date" id="t_date"></div>
<div><label class="req">Party</label><select id="t_party"></select></div>
<div><label class="req">Vehicle</label><select id="t_vehicle"></select></div>
<div><label class="req">Driver</label><select id="t_driver"></select></div>
<div class="full" id="t_d2box" style="display:none"><div class="dl"><label>Second driver</label><button type="button" class="more" id="t_d2x">Remove</button></div><select id="t_driver2"></select></div>
<div class="full" id="t_d3box" style="display:none"><div class="dl"><label>Third driver</label><button type="button" class="more" id="t_d3x">Remove</button></div><select id="t_driver3"></select></div>
<div class="full"><button type="button" class="more" id="t_addDrv">+ Add another driver</button></div>
<div><label>Transporter</label><select id="t_transporter"></select></div>
<div><label>Commodity</label><input id="t_commodity" list="comms" autocomplete="off"></div>
<div><label class="req">From</label><input id="t_from" list="locs" autocomplete="off"></div>
<div><label class="req">To</label><input id="t_to" list="locs" autocomplete="off"></div>
</div>
<div class="seg"><button type="button" id="t_tq" class="on">Per quintal</button><button type="button" id="t_tf">Fixed amount</button></div>
<div class="g">
<div><label>Weight (quintal)</label><input id="t_weight" type="number" inputmode="decimal" min="0" step="any"></div>
<div id="t_rateBox"><label>Rate ₹/quintal</label><input id="t_rate" type="number" inputmode="decimal" min="0" step="any"></div>
<div id="t_fixBox" style="display:none"><label>Fixed freight ₹</label><input id="t_fixed" type="number" inputmode="decimal" min="0" step="any"></div>
<div><label>Advance ₹</label><input id="t_advance" type="number" inputmode="decimal" min="0" step="any"></div>
<div><label>Freight ₹ (auto)</label><div class="out" id="t_freight">–</div></div>
</div>
<div class="bal"><span id="t_balLbl">Balance</span><b id="t_balance">–</b></div>
<div class="info" id="t_payinfo" style="display:none"></div>
<button type="button" class="more" id="t_moreBtn">+ More (via, party phone, notes)</button>
<div id="t_more" class="g" style="display:none">
<div><label>Via</label><input id="t_via" list="locs" autocomplete="off"></div>
<div><label>Party phone</label><input id="t_phone" type="tel" inputmode="tel"></div>
<div class="full"><label>Notes</label><textarea id="t_notes"></textarea></div>
</div>
<div class="row"><button type="button" class="btn2" id="t_cancel" style="display:none">Cancel</button><button type="button" class="save" id="t_save">Save trip</button></div>
</div></div>

<div class="view" id="v_payment"><div class="c">
<div><h1 id="p_title">Payment received</h1></div>
<div class="g">
<div><label class="req">Date</label><input type="date" id="p_date"></div>
<div><label class="req">Party</label><select id="p_party"></select></div>
<div class="full"><label class="req">Trip</label><select id="p_trip"></select></div>
</div>
<div class="info" id="p_info" style="display:none"></div>
<div class="prev" id="p_prev" style="display:none"></div>
<div class="g">
<div><label>Amount received ₹</label><input id="p_amount" type="number" inputmode="decimal" min="0" step="any"></div>
<div><label>Expense ₹</label><input id="p_expense" type="number" inputmode="decimal" min="0" step="any"></div>
</div>
<div class="hint">Expense reduces what is left, for example unloading, shortage or commission. Enter an amount, an expense, or both.</div>
<div class="bal"><span>Left after this</span><b id="p_left">–</b></div>
<div><label id="p_modeLbl">Mode</label><div class="chips" id="p_mode"></div></div>
<div><label>UTR / cheque no. / bill no.</label><input id="p_ref" autocomplete="off"></div>
<button type="button" class="more" id="p_moreBtn">+ More (notes)</button>
<div id="p_more" style="display:none"><label>Notes</label><textarea id="p_notes"></textarea></div>
<div class="row"><button type="button" class="btn2" id="p_cancel" style="display:none">Cancel</button><button type="button" class="save" id="p_save">Save payment</button></div>
</div></div>

<div class="view" id="v_trips"><div class="c">
<div class="hd2"><h1>Trip list</h1><div class="hbtns"><button type="button" class="more" id="l_print">Print / PDF</button><button type="button" class="more" id="l_refresh">Refresh</button></div></div>
<div class="sub" id="l_printOut"></div>
<input id="l_q" type="search" placeholder="Search party, vehicle, driver, place or trip ID" autocomplete="off">
<div class="chips" id="l_f"></div>
<div><label>Filter by</label><div class="seg" id="l_by"></div></div>
<div><label id="l_whichLbl">Party</label><select id="l_which"></select></div>
<div class="chips" id="l_d"></div>
<div class="g"><div><label>From date</label><input type="date" id="l_from"></div><div><label>To date</label><input type="date" id="l_to"></div></div>
<div class="totals" id="l_tot"></div>
<div id="l_list"></div>
<button type="button" class="btn2" id="l_moreBtn" style="display:none">Show more</button>
</div></div>

<div class="view" id="v_summary"><div><div class="c" id="s_main">
<div class="hd2"><h1>Summary</h1><div class="hbtns"><button type="button" class="more" id="s_print">Print / PDF</button><button type="button" class="more" id="s_refresh">Refresh</button></div></div>
<div class="sub" id="s_printOut"></div>
<div><label>Summary by</label><div class="seg" id="s_by"></div></div>
<div><label id="s_whichLbl">Party</label><select id="s_which"></select></div>
<div class="chips" id="s_d"></div>
<div class="g"><div><label>From date</label><input type="date" id="s_from"></div><div><label>To date</label><input type="date" id="s_to"></div></div>
<div class="kpis" id="s_kpi"></div>
<div class="sub" id="s_note"></div>
<div id="s_act" style="display:none;flex-direction:column;gap:8px">
<div class="minrow"><label for="s_min">Only trips pending above ₹</label><input id="s_min" type="number" inputmode="decimal" min="0" step="any" placeholder="0"></div>
<div class="row"><button type="button" class="btn2" id="s_stmtBtn">View statement</button><a class="save wa" id="s_wa" target="_blank" rel="noopener">WhatsApp reminder</a></div>
<div class="sub" id="s_actNote"></div>
</div>
<div class="h3" id="s_listTitle" style="font-weight:600;font-size:15px"></div>
<div id="s_list"></div>
</div>
<div id="s_stmt" class="c" style="display:none"></div>
</div></div>

<div class="view" id="v_party"></div>
<div class="view" id="v_vehicle"></div>
<div class="view" id="v_driver"></div>
<div class="view" id="v_transporter"></div>
<datalist id="locs"></datalist><datalist id="comms"></datalist>
</div>
<script>
var D = __DATA__;
var sPickInit = '', lPickInit = '';
var TABS = [['trip','New trip'],['payment','Payment'],['trips','Trip list'],['summary','Summary']];
var TABS2 = [['party','Party'],['vehicle','Vehicle'],['driver','Driver'],['transporter','Transporter']];
var NEW = '__new__', ACCT = '__acct__';
var SELS = { t_vehicle: 'vehicle', t_driver: 'driver', t_driver2: 'driver', t_driver3: 'driver', t_party: 'party', t_transporter: 'transporter', p_party: 'party' };
var FILTERS = ['Not cleared', 'Due', 'Freight not entered', 'Cleared', 'All'];
var tripMode = 'Per quintal', payMode = '', qaBox = null, editTrip = null, editPay = null, backToList = false;
var listFilter = 'Not cleared', listLimit = 50, openId = '';
var DATES = ['Last 7 days', 'Last 30 days', 'This month', 'Last month', 'Any date'];
var LBY = [['party','Party','parties'],['driver','Driver','drivers'],['vehicle','Vehicle','vehicles'],['transporter','Transporter','transporters']], lBy = 'party', lRows = [];
var MONTHS = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function $(i){ return document.getElementById(i); }
function inr(x){ return Math.round(x).toLocaleString('en-IN'); }
function qty(x){ return Number(x || 0).toLocaleString('en-IN', { maximumFractionDigits: 2 }); }
function num(id){ var v = parseFloat($(id).value); return isNaN(v) ? 0 : v; }
function norm(type, s){ s = String(s || '').trim(); return type === 'vehicle' ? s.toUpperCase().replace(/[\\s-]/g, '') : s.toLowerCase().replace(/\\s+/g, ' '); }
function names(type){ var k = D.defs[type].key; return D.rec[type].map(function(r){ return r[k]; }); }
function findRec(type, name){ var k = D.defs[type].key, n = norm(type, name); for (var i = 0; i < D.rec[type].length; i++) if (norm(type, D.rec[type][i][k]) === n) return D.rec[type][i]; return null; }
function opt(s, v, t){ var o = document.createElement('option'); o.value = v; o.textContent = t; s.appendChild(o); }
function show(t, cls){ var m = $('msg'); m.textContent = t; m.className = 'msg ' + cls; m.style.display = 'block'; window.scrollTo(0, 0); }
function hideMsg(){ $('msg').style.display = 'none'; }
function errText(e){ return e && e.message ? e.message : String(e); }
function clearBad(root){ root.querySelectorAll('.bad').forEach(function(x){ x.classList.remove('bad'); }); }
function markMissing(ids){ var miss = ids.filter(function(i){ return !String($(i).value).trim(); }); miss.forEach(function(i){ $(i).classList.add('bad'); }); return miss.length; }
function busy(b, on, label){ b.disabled = on; b.textContent = on ? 'Saving…' : label; }
function askBox(before, text, buttons){
  var old = document.getElementById('askbox'); if (old) old.remove();
  var box = document.createElement('div'); box.className = 'note'; box.id = 'askbox';
  var t = document.createElement('div'); t.textContent = text; box.appendChild(t);
  var row = document.createElement('div'); row.className = 'pick'; box.appendChild(row);
  buttons.forEach(function(x){
    var b = document.createElement('button'); b.type = 'button'; b.textContent = x.t; if (x.cls) b.className = x.cls;
    b.onclick = function(){ box.remove(); if (x.fn) x.fn(); }; row.appendChild(b);
  });
  before.parentNode.insertBefore(box, before);
  box.scrollIntoView({ block: 'center', behavior: 'smooth' });
}

/* ---------- similar names ---------- */
var NOISE = ['ms','messrs','shri','sri','pvt','private','ltd','limited','llp','co','company','and'];
function core(s){
  s = String(s || '').toLowerCase().replace(/\\bm\\s*\\/\\s*s\\b\\.?/g, ' ').replace(/&/g, ' and ').replace(/[^a-z0-9\\u0900-\\u097f]+/g, ' ');
  return s.split(' ').filter(function(w){ return w && NOISE.indexOf(w) < 0; }).join(' ');
}
function lev(a, b){
  if (Math.abs(a.length - b.length) > 2) return 9;
  var d = [], i, j;
  for (i = 0; i <= a.length; i++) d[i] = [i];
  for (j = 0; j <= b.length; j++) d[0][j] = j;
  for (i = 1; i <= a.length; i++) for (j = 1; j <= b.length; j++) {
    d[i][j] = Math.min(d[i - 1][j] + 1, d[i][j - 1] + 1, d[i - 1][j - 1] + (a[i - 1] === b[j - 1] ? 0 : 1));
    if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1]) d[i][j] = Math.min(d[i][j], d[i - 2][j - 2] + 1);
  }
  return d[a.length][b.length];
}
function similarNames(type, name){
  var out = [], k = D.defs[type].key, veh = type === 'vehicle';
  var a = (veh ? norm('vehicle', name) : core(name)).replace(/ /g, '');
  if (!a) return out;
  D.rec[type].forEach(function(r){
    var nm = r[k];
    if (norm(type, nm) === norm(type, name) || out.length >= 3) return;
    var b = (veh ? norm('vehicle', nm) : core(nm)).replace(/ /g, '');
    if (!b) return;
    var big = Math.max(a.length, b.length), small = Math.min(a.length, b.length);
    var hit = veh ? lev(a, b) <= 1
      : a === b || (small >= 4 && (a.indexOf(b) >= 0 || b.indexOf(a) >= 0)) || lev(a, b) <= (big <= 3 ? 0 : big <= 6 ? 1 : 2);
    if (hit) out.push(nm);
  });
  return out;
}
function refreshReports(){ google.script.run.withFailureHandler(function(){}).refreshAfterSave(); }



function pc(s){
  s = String(s || '').replace(/\\s+/g, ' ').trim().toLowerCase();
  if (s === 'self / own vehicle') return 'Self / Own vehicle';
  return s.replace(/(^|[\\s.(-])([a-z])/g, function(m, p, c){ return p + c.toUpperCase(); })
    .replace(/ (And|Of|The) /g, function(m, w){ return ' ' + w.toLowerCase() + ' '; });
}
/* ---------- who is entering (asked once per phone) ---------- */
var WHO = '';
try { WHO = localStorage.getItem('tt_name') || ''; } catch (e) { WHO = ''; }
// Owner: Google gives the email, nothing is asked. Staff: pick a name from the list once per phone
// (Staff tab + Driver master). Typing is only used when that list is still empty.
function whoList(){
  var s = $('whoSel'), list = D.staff || [];
  s.innerHTML = ''; opt(s, '', 'Choose your name');
  list.forEach(function(n){ opt(s, n, n); });
  var typing = !list.length;
  (s._combo || s).style.display = typing ? 'none' : ''; $('whoIn').style.display = typing ? '' : 'none';
  $('whoHint').textContent = typing ? 'No names in the Staff list yet, so type your name. The owner can add names in the Staff tab of the sheet.' :
    'Not in the list? Ask the owner to add your name in the Staff tab of the sheet.';
}
function showWho(){
  var t = $('whoTxt'); t.innerHTML = '';
  if (D.email) { WHO = D.email; t.textContent = D.email; $('nameBox').style.display = 'none'; return; }
  if (WHO) {
    t.appendChild(document.createTextNode(WHO + ' · '));
    var a = document.createElement('a'); a.textContent = 'change';
    a.onclick = function(){ whoList(); selVal($('whoSel'), WHO); $('whoIn').value = WHO; $('nameBox').style.display = 'block'; };
    t.appendChild(a); $('nameBox').style.display = 'none';
  } else { whoList(); $('nameBox').style.display = 'block'; }
}
function needWho(){
  if (WHO) return false;
  $('nameBox').style.display = 'block'; $('whoSel').classList.add('bad'); $('whoIn').classList.add('bad');
  show('Choose your name at the top first. It is asked only once on this phone.', 'err');
  return true;
}
function fmtDate(s){ if (!s) return ''; var p = String(s).split('-'); return (+p[2]) + ' ' + MONTHS[+p[1] - 1] + ' ' + p[0]; }
function fmtShort(s){ if (!s) return ''; var p = String(s).split('-'); return (+p[2]) + ' ' + MONTHS[+p[1] - 1]; }
function money(v){ return v === '' || v == null ? '–' : '₹' + inr(v); }
function findTrip(id){ for (var i = 0; i < D.trips.length; i++) if (D.trips[i].id === id) return D.trips[i]; return null; }
function selVal(sel, v){
  v = v == null ? '' : String(v);
  var k = v.trim().toLowerCase(), hit = null;
  Array.prototype.forEach.call(sel.options, function(o){ if (!hit && o.value !== NEW && o.value.trim().toLowerCase() === k) hit = o.value; });
  if (v && hit === null) { opt(sel, v, v); hit = v; }
  sel.value = v ? hit : '';
}

/* ---------- tabs ---------- */
function go(t, noReload){
  TABS.concat(TABS2).forEach(function(x){ $('tab_' + x[0]).className = x[0] === t ? 'on' : ''; $('v_' + x[0]).style.display = x[0] === t ? 'block' : 'none'; });
  hideMsg(); closeQuick();
  var a = $('askbox'); if (a) a.remove();
  if (t === 'trips') { renderList(); if (!noReload) reloadTrips(); }
  if (t === 'summary') { $('s_stmt').style.display = 'none'; $('s_main').style.display = 'flex'; renderSummary(); if (!noReload) reloadTrips(); }
  saveState();
}
/* ---------- page address keeps the tab and filters, so Refresh and a browser reload come back to them ---------- */
function curTab(){ var t = D.start || 'trip'; TABS.concat(TABS2).forEach(function(x){ if ($('v_' + x[0]).style.display === 'block') t = x[0]; }); return t; }
function stateParams(){
  var t = curTab(), p = { f: t };
  if (t === 'summary') { p.by = sBy; p.which = $('s_which').value; p.from = $('s_from').value; p.to = $('s_to').value; p.min = $('s_min').value; }
  if (t === 'trips') { p.lf = listFilter; p.lby = lBy; p.lwhich = $('l_which').value; p.lfrom = $('l_from').value; p.lto = $('l_to').value; p.q = $('l_q').value; }
  return p;
}
function saveState(){ try { google.script.history.replace(null, stateParams()); } catch (e) {} }
// Refresh = open the page again (all trips, payments, parties, vehicles, drivers and today's date are
// loaded fresh) on the same tab with the same filters.
function fullReload(btn){
  btn.disabled = true; btn.textContent = 'Refreshing…';
  var p = stateParams();
  var qs = Object.keys(p).map(function(k){ return encodeURIComponent(k) + '=' + encodeURIComponent(p[k] == null ? '' : p[k]); }).join('&');
  function fallback(){
    reloadTrips(function(ok){
      btn.disabled = false; btn.textContent = ok ? 'Updated' : 'Refresh';
      setTimeout(function(){ btn.textContent = 'Refresh'; }, 2500);
    });
  }
  if (D.url) {
    try { window.open(D.url + '?' + qs, '_top'); } catch (e) { fallback(); return; }
    setTimeout(fallback, 8000);   // only runs if the browser did not open the page again
    return;
  }
  reloadTrips(function(ok){
    btn.disabled = false; btn.textContent = ok ? 'Updated' : 'Refresh';
    setTimeout(function(){ btn.textContent = 'Refresh'; }, 2500);
  });
}
[[TABS, 'tabs'], [TABS2, 'tabs2']].forEach(function(g){
  g[0].forEach(function(x){
    var b = document.createElement('button'); b.type = 'button'; b.id = 'tab_' + x[0]; b.textContent = x[1];
    b.onclick = function(){ go(x[0]); }; $(g[1]).appendChild(b);
  });
});

function reloadTrips(cb){
  $('l_tot').textContent = 'Refreshing…';
  google.script.run
    .withSuccessHandler(function(r){ D.trips = r.trips; D.pending = r.pending; D.onacc = r.onacc || []; renderList(); renderSummary(); payTripOpts(true); if (cb) cb(true); })
    .withFailureHandler(function(e){ renderList(); show('Could not load the latest trips: ' + errText(e) + '. Check the internet and tap Refresh.', 'err'); if (cb) cb(false); })
    .getAppTrips();
}

/* ---------- dropdowns with "+ Add new" ---------- */
function fillSel(id){
  var type = SELS[id], s = $(id), keep = s.value;
  s.innerHTML = '';
  opt(s, '', 'Choose');
  if (id === 't_transporter') opt(s, 'Self / Own vehicle', 'Self / Own vehicle');
  names(type).forEach(function(v){ opt(s, v, v); });
  opt(s, NEW, '+ Add new ' + D.defs[type].label.toLowerCase());
  if (keep && keep !== NEW) selVal(s, keep);
}
function refreshType(type){
  Object.keys(SELS).forEach(function(id){ if (SELS[id] === type) fillSel(id); });
  var dl = $('dl_' + type);
  if (dl) { dl.innerHTML = ''; names(type).forEach(function(v){ opt(dl, v, v); }); }
}
function upsertRec(type, rec){
  var k = D.defs[type].key, n = norm(type, rec[k]), list = D.rec[type], done = false;
  for (var i = 0; i < list.length; i++) if (norm(type, list[i][k]) === n) { list[i] = rec; done = true; break; }
  if (!done) list.push(rec);
  list.sort(function(a, b){ return String(a[k]).localeCompare(String(b[k])); });
  refreshType(type);
}
function fillList(id, arr){ var dl = $(id); dl.innerHTML = ''; arr.forEach(function(v){ opt(dl, v, v); }); }
function addList(arr, vals){ vals.forEach(function(v){ v = String(v || '').trim(); if (!v) return; var k = v.toLowerCase(); if (!arr.some(function(x){ return x.toLowerCase() === k; })) arr.push(v); }); arr.sort(); }

/* ---------- field builder (master forms + quick add) ---------- */
function field(f, id){
  var wrap = document.createElement('div'); if (f.full) wrap.className = 'full';
  var lb = document.createElement('label'); lb.textContent = f.h; if (f.req) lb.className = 'req'; wrap.appendChild(lb);
  var inp;
  if (f.type === 'select') {
    inp = document.createElement('select'); opt(inp, '', 'Choose');
    f.choices.forEach(function(c){ opt(inp, c, c); });
  } else if (f.area) {
    inp = document.createElement('textarea');
  } else {
    inp = document.createElement('input');
    inp.type = { number: 'number', date: 'date', tel: 'tel', email: 'email' }[f.type] || 'text';
    if (f.type === 'number') { inp.inputMode = 'decimal'; inp.min = '0'; inp.step = 'any'; }
    if (f.type === 'tel') inp.inputMode = 'tel';
    inp.autocomplete = 'off';
    if (f.upper) inp.style.textTransform = 'uppercase';
    if (f.upper) inp.addEventListener('change', function(){ inp.value = inp.value.toUpperCase().replace(/[\\s-]/g, ''); });
    if (f.proper) inp.addEventListener('change', function(){ inp.value = pc(inp.value); });
  }
  inp.id = id;
  inp.addEventListener('input', function(){ inp.classList.remove('bad'); });
  wrap.appendChild(inp);
  return wrap;
}
function setVal(inp, v){
  if (inp.tagName === 'SELECT') { selVal(inp, v); return; }
  inp.value = v == null ? '' : String(v);
}

/* ---------- quick add (inside Trip / Payment) ---------- */
function closeQuick(){
  if (qaBox) { qaBox.remove(); qaBox = null; }
  Object.keys(SELS).forEach(function(id){ if ($(id).value === NEW) $(id).value = ''; });   // never leave a box stuck on "+ Add new"
}
function openQuick(selId){
  closeQuick();
  var type = SELS[selId], def = D.defs[type], sel = $(selId);
  var box = document.createElement('div'); box.className = 'qa';
  var h = document.createElement('h2'); h.textContent = 'Add new ' + def.label.toLowerCase(); box.appendChild(h);
  var g = document.createElement('div'); g.className = 'g'; box.appendChild(g);
  var flds = def.fields.filter(function(f){ return def.quick.indexOf(f.h) >= 0; });
  flds.forEach(function(f, i){ var x = field(f, 'qa_' + i); x.className = ''; g.appendChild(x); });
  var p = document.createElement('div'); p.className = 'sub'; p.textContent = 'Saved to the ' + def.label + ' master. Add the other details later in the ' + def.label + ' tab.'; box.appendChild(p);
  var row = document.createElement('div'); row.className = 'row';
  var cancel = document.createElement('button'); cancel.type = 'button'; cancel.className = 'btn2'; cancel.textContent = 'Cancel';
  var save = document.createElement('button'); save.type = 'button'; save.className = 'save'; save.textContent = 'Add';
  row.appendChild(cancel); row.appendChild(save); box.appendChild(row);
  var grid = sel.closest('.g') || sel.parentNode;
  grid.parentNode.insertBefore(box, grid.nextSibling);
  qaBox = box;
  $('qa_0').focus();
  cancel.onclick = function(){ sel.value = ''; closeQuick(); };
  save.onclick = function(){ tryAdd(false); };
  function tryAdd(force){
    if (needWho()) return;
    var ids = []; flds.forEach(function(f, i){ if (f.req) ids.push('qa_' + i); });
    if (markMissing(ids)) return;
    var nm = $('qa_0').value, ex = findRec(type, nm);
    if (ex) { sel.value = ex[def.key]; closeQuick(); sel.dispatchEvent(new Event('change')); show(def.label + ' ' + ex[def.key] + ' already exists, selected it.', 'ok'); return; }
    var sim = force ? [] : similarNames(type, nm);
    if (sim.length) {
      askBox(row, 'Did you mean ' + sim.join(' or ') + '?', sim.map(function(x){
        return { t: 'Use ' + x, fn: function(){ sel.value = x; closeQuick(); sel.dispatchEvent(new Event('change')); show('Selected existing ' + def.label.toLowerCase() + ' ' + x + '.', 'ok'); } };
      }).concat([{ t: 'No, add as new', cls: 'alt', fn: function(){ tryAdd(true); } }]));
      return;
    }
    var d = { _quick: true, by: WHO };
    flds.forEach(function(f, i){ d[f.h] = $('qa_' + i).value; });
    busy(save, true, 'Add');
    var slow = setTimeout(function(){ if (save.disabled) show('Still saving… please wait a few seconds.', 'ok'); }, 6000);
    google.script.run
      .withSuccessHandler(function(res){ clearTimeout(slow); upsertRec(type, res.record); sel.value = res.name; closeQuick(); sel.dispatchEvent(new Event('change')); show('Added ' + def.label.toLowerCase() + ' ' + res.name + ' to the master.', 'ok'); })
      .withFailureHandler(function(e){ clearTimeout(slow); busy(save, false, 'Add'); show('Not added: ' + errText(e), 'err'); })
      .saveMaster(type, d);
  }
}

Object.keys(SELS).forEach(function(id){
  $(id).addEventListener('change', function(){
    $(id).classList.remove('bad');
    if ($(id).value === NEW) openQuick(id); else if (qaBox) closeQuick();
    if (id === 'p_party') payTripOpts(false);
  });
});

/* ---------- new trip / edit trip ---------- */
function tripFreight(){
  if (tripMode === 'Fixed amount') return $('t_fixed').value.trim() === '' ? null : num('t_fixed');
  if ($('t_weight').value.trim() === '' || $('t_rate').value.trim() === '') return null;
  return Math.round(num('t_weight') * num('t_rate') * 100) / 100;
}
function tripCalc(){
  var f = tripFreight();
  $('t_freight').textContent = f === null ? '–' : inr(f);
  $('t_balLbl').textContent = editTrip ? 'Remaining (after payments and expenses)' : 'Balance';
  $('t_balance').textContent = f === null ? '–' : '₹' + inr(f - num('t_advance') - (editTrip ? editTrip.paid + editTrip.expense : 0));
}
function setTripMode(m){
  tripMode = m; var fx = m === 'Fixed amount';
  $('t_tq').className = fx ? '' : 'on'; $('t_tf').className = fx ? 'on' : '';
  $('t_rateBox').style.display = fx ? 'none' : ''; $('t_fixBox').style.display = fx ? '' : 'none';
  tripCalc();
}
function tripMore(open){ $('t_more').style.display = open ? 'grid' : 'none'; $('t_moreBtn').textContent = open ? '− Less' : '+ More (via, party phone, notes)'; }
function drvList(s){ return String(s || '').split(',').map(function(x){ return x.trim(); }).filter(Boolean); }
function showDrivers(n){
  $('t_d2box').style.display = n >= 2 ? '' : 'none';
  $('t_d3box').style.display = n >= 3 ? '' : 'none';
  $('t_addDrv').style.display = n >= 3 ? 'none' : '';
  if (n < 3) $('t_driver3').value = '';
  if (n < 2) $('t_driver2').value = '';
}
function driverCount(){ return $('t_d3box').style.display !== 'none' ? 3 : $('t_d2box').style.display !== 'none' ? 2 : 1; }
$('t_addDrv').onclick = function(){ showDrivers(driverCount() + 1); var q = $(driverCount() === 3 ? 't_driver3_q' : 't_driver2_q'); if (q) q.focus(); };
$('t_d2x').onclick = function(){ $('t_driver2').value = $('t_driver3').value; showDrivers(driverCount() - 1); };
$('t_d3x').onclick = function(){ $('t_driver3').value = ''; showDrivers(2); };
function tripReset(){
  ['t_from','t_to','t_weight','t_rate','t_fixed','t_advance','t_via','t_commodity','t_phone','t_notes','t_vehicle','t_driver','t_driver2','t_driver3','t_party'].forEach(function(i){ $(i).value = ''; });
  showDrivers(1);
  $('t_transporter').value = 'Self / Own vehicle'; $('t_date').value = D.today;
  editTrip = null;
  $('t_title').textContent = 'New trip';
  $('t_sub').textContent = 'Only boxes with * are needed now. Weight, rate and advance can be added later from Trip list.';
  $('t_save').textContent = 'Save trip'; $('t_cancel').style.display = 'none'; $('t_payinfo').style.display = 'none';
  clearBad($('v_trip')); tripMore(false); setTripMode('Per quintal'); syncCombos();
}
function startEditTrip(t){
  go('trip'); tripReset(); editTrip = t;
  $('t_title').textContent = 'Edit trip ' + t.id;
  $('t_sub').textContent = 'Change what you need and tap Update trip.';
  $('t_date').value = t.date;
  var dl = drvList(t.driver);
  selVal($('t_party'), t.party); selVal($('t_vehicle'), t.vehicle); selVal($('t_driver'), dl[0] || '');
  showDrivers(Math.min(3, Math.max(1, dl.length)));
  if (dl[1]) selVal($('t_driver2'), dl[1]);
  if (dl[2]) selVal($('t_driver3'), dl[2]);
  selVal($('t_transporter'), t.transporter);
  $('t_commodity').value = t.commodity; $('t_from').value = t.from; $('t_to').value = t.to;
  $('t_via').value = t.via; $('t_phone').value = t.phone; $('t_notes').value = t.notes;
  setTripMode(t.ftype === 'Fixed amount' ? 'Fixed amount' : 'Per quintal');
  $('t_weight').value = t.weight; $('t_rate').value = t.rate;
  $('t_fixed').value = t.ftype === 'Fixed amount' ? t.freight : '';
  $('t_advance').value = t.advance;
  tripMore(!!(t.via || t.phone || t.notes));
  $('t_payinfo').textContent = 'So far for this trip: paid ₹' + inr(t.paid) + ', expense ₹' + inr(t.expense) + '. Payments are changed from Trip list.';
  $('t_payinfo').style.display = 'block';
  $('t_save').textContent = 'Update trip'; $('t_cancel').style.display = '';
  tripCalc(); syncCombos(); window.scrollTo(0, 0);
}
$('t_tq').onclick = function(){ setTripMode('Per quintal'); };
$('t_tf').onclick = function(){ setTripMode('Fixed amount'); };
$('t_moreBtn').onclick = function(){ tripMore($('t_more').style.display === 'none'); };
['t_weight','t_rate','t_fixed','t_advance'].forEach(function(i){ $(i).addEventListener('input', tripCalc); });
$('t_cancel').onclick = function(){ tripReset(); go('trips', true); };
$('t_save').onclick = function(){
  if (needWho()) return;
  clearBad($('v_trip'));
  var req = ['t_date','t_party','t_vehicle','t_driver','t_from','t_to'];
  if (markMissing(req) || [$('t_vehicle'), $('t_driver'), $('t_driver2'), $('t_driver3'), $('t_party')].some(function(s){ return s.value === NEW; })) { show('Fill the boxes marked in red.', 'err'); return; }
  var dsel = ['t_driver','t_driver2','t_driver3'].map(function(i){ return $(i).value; }).filter(Boolean);
  if (dsel.length !== drvList(dsel.join(',')).filter(function(n, i, a){ return a.map(function(x){ return x.toLowerCase(); }).indexOf(n.toLowerCase()) === i; }).length) {
    show('The same driver is chosen twice. Pick a different second driver or remove it.', 'err'); return;
  }
  var d = { type: tripMode, id: editTrip ? editTrip.id : '', by: WHO };
  ['date','vehicle','driver','driver2','driver3','party','from','to','weight','rate','fixed','advance','via','commodity','transporter','phone','notes'].forEach(function(k){ d[k] = $('t_' + k).value; });
  if (d.driver2 === NEW) d.driver2 = ''; if (d.driver3 === NEW) d.driver3 = '';
  if (d.transporter === NEW) d.transporter = '';
  var b = $('t_save'), label = editTrip ? 'Update trip' : 'Save trip', f = tripFreight(), wasEdit = !!editTrip;
  function send(force){
    d.force = force;
    busy(b, true, label);
    google.script.run
      .withSuccessHandler(function(res){
        busy(b, false, label);
        if (res && res.dup) {
          askBox(b.parentNode, 'This trip looks already saved: ' + res.info + '. Save it again?',
            [{ t: 'Save anyway', fn: function(){ send(true); } }, { t: 'Cancel', cls: 'alt' }]);
          return;
        }
        addList(D.locations, [d.from, d.via, d.to]); fillList('locs', D.locations);
        addList(D.commodities, [d.commodity]); fillList('comms', D.commodities);
        var txt = 'Trip ' + res.id + (wasEdit ? ' updated.' : ' saved.') +
          (f === null ? ' Add weight and rate later from Trip list.' : ' Freight ₹' + inr(f) + '.');
        tripReset(); refreshReports();
        if (wasEdit) { openId = res.id; go('trips', true); }
        reloadTrips();
        show(txt, 'ok');
      })
      .withFailureHandler(function(e){ show('Not saved: ' + errText(e), 'err'); busy(b, false, label); })
      .saveTrip(d);
  }
  send(false);
};

/* ---------- payment (trip-wise) ---------- */
function partyTrips(party){ var k = norm('party', party); return D.trips.filter(function(t){ return norm('party', t.party) === k; }); }
function payTripOpts(keep){
  var s = $('p_trip'), cur = keep ? s.value : '', party = $('p_party').value;
  s.innerHTML = '';
  if (!party || party === NEW) { opt(s, '', 'Choose the party first'); payCalc(); return; }
  opt(s, '', 'Choose trip');
  var list = partyTrips(party);
  list.filter(function(t){ return t.status !== 'Cleared'; }).concat(list.filter(function(t){ return t.status === 'Cleared'; })).forEach(function(t){
    opt(s, t.id, t.id + ' · ' + fmtShort(t.date));
  });
  opt(s, ACCT, 'On account (not for one trip)');
  if (cur) { s.value = cur; if (s.value !== cur) { opt(s, cur, cur); s.value = cur; } }
  payCalc();
}
function payCalc(){
  var box = $('p_info'), tid = $('p_trip').value, party = $('p_party').value, base = null, txt = '', prev = '';
  var mine = function(f){ return editPay ? editPay[f] : 0; };
  if (tid && tid !== ACCT) {
    var t = findTrip(tid);
    if (t) {
      var same = editPay && editPay.tripId === tid;
      var paid = t.paid - (same ? mine('amount') : 0), ex = t.expense - (same ? mine('expense') : 0);
      base = (t.freight === '' ? 0 : t.freight) - (t.advance === '' ? 0 : t.advance) - paid - ex;
      var others = t.pays.filter(function(p){ return !(editPay && p.id === editPay.id); });
      prev = others.length ? (editPay ? 'Other payments for this trip:\\n' : 'Earlier payments for this trip:\\n') + others.map(function(p){
        return fmtShort(p.date) + (p.amount ? ' · ₹' + inr(p.amount) + ' ' + (p.mode.indexOf('Bank') === 0 ? 'Bank' : p.mode) : '') +
          (p.expense ? ' · expense ₹' + inr(p.expense) : '') + (p.ref ? ' · ' + p.ref : '');
      }).join('\\n') : (editPay ? 'No other payments for this trip.' : 'No earlier payments for this trip.');
      txt = t.from + ' → ' + t.to + ' · ' + fmtDate(t.date) + ' · ' + t.vehicle + '\\n' +
        (t.freight === '' ? 'Freight not entered yet for this trip.' :
          'Freight ₹' + inr(t.freight) + ' − advance ₹' + inr(t.advance || 0) + ' − paid ₹' + inr(paid) + ' − expense ₹' + inr(ex) + ' = remaining ₹' + inr(base));
    }
  } else if (party && party !== NEW) {
    base = (D.pending[norm('party', party)] || 0) + (editPay && norm('party', editPay.party) === norm('party', party) ? mine('amount') + mine('expense') : 0);
    txt = base > 0 ? 'Party pending (all trips): ₹' + inr(base) : 'No pending amount for this party.';
  }
  box.textContent = txt; box.style.display = txt ? 'block' : 'none';
  $('p_prev').textContent = prev; $('p_prev').style.display = prev ? 'block' : 'none';
  $('p_left').textContent = base === null ? '–' : '₹' + inr(base - num('p_amount') - num('p_expense'));
  $('p_modeLbl').className = num('p_amount') > 0 ? 'req' : '';
}
function setPayMode(m){
  payMode = m;
  Array.prototype.forEach.call($('p_mode').children, function(x){ x.className = x.dataset.m === m ? 'on' : ''; });
}
D.modes.forEach(function(m){
  var b = document.createElement('button'); b.type = 'button'; b.dataset.m = m; b.textContent = m.indexOf('Bank') === 0 ? 'Bank' : m;
  b.onclick = function(){ setPayMode(m); }; $('p_mode').appendChild(b);
});
function payMore(open){ $('p_more').style.display = open ? 'block' : 'none'; $('p_moreBtn').textContent = open ? '− Less' : '+ More (notes)'; }
function payReset(){
  ['p_party','p_amount','p_expense','p_ref','p_notes'].forEach(function(i){ $(i).value = ''; });
  $('p_date').value = D.today; setPayMode(''); editPay = null; backToList = false;
  $('p_title').textContent = 'Payment received'; $('p_save').textContent = 'Save payment'; $('p_cancel').style.display = 'none';
  clearBad($('v_payment')); payMore(false); payTripOpts(false); syncCombos();
}
function startPayment(t){
  go('payment'); payReset(); backToList = true;
  selVal($('p_party'), t.party); payTripOpts(false); $('p_trip').value = t.id; payCalc(); syncCombos();
  $('p_cancel').style.display = '';
}
function startEditPayment(p){
  go('payment'); payReset(); editPay = p; backToList = true;
  $('p_title').textContent = 'Edit payment ' + p.id;
  $('p_date').value = p.date; selVal($('p_party'), p.party); payTripOpts(false);
  var s = $('p_trip'); s.value = p.tripId || ACCT; if (s.value !== (p.tripId || ACCT)) { opt(s, p.tripId, p.tripId); s.value = p.tripId; }
  $('p_amount').value = p.amount || ''; $('p_expense').value = p.expense || '';
  setPayMode(p.mode); $('p_ref').value = p.ref; $('p_notes').value = p.notes; payMore(!!p.notes);
  $('p_save').textContent = 'Update payment'; $('p_cancel').style.display = '';
  payCalc(); syncCombos(); window.scrollTo(0, 0);
}
['p_amount','p_expense'].forEach(function(i){ $(i).addEventListener('input', payCalc); });
$('p_trip').addEventListener('change', function(){ $('p_trip').classList.remove('bad'); payCalc(); });
$('p_moreBtn').onclick = function(){ payMore($('p_more').style.display === 'none'); };
$('p_cancel').onclick = function(){ payReset(); go('trips', true); };
$('p_save').onclick = function(){
  if (needWho()) return;
  clearBad($('v_payment'));
  if (markMissing(['p_date','p_party','p_trip']) || $('p_party').value === NEW) { show('Fill the boxes marked in red.', 'err'); return; }
  if (num('p_amount') <= 0 && num('p_expense') <= 0) { $('p_amount').classList.add('bad'); $('p_expense').classList.add('bad'); show('Enter the amount received or an expense.', 'err'); return; }
  if (num('p_amount') > 0 && !payMode) { show('Choose the payment mode.', 'err'); return; }
  var tid = $('p_trip').value;
  var d = { by: WHO, id: editPay ? editPay.id : '', date: $('p_date').value, party: $('p_party').value, tripId: tid === ACCT ? '' : tid,
    amount: $('p_amount').value, expense: $('p_expense').value, mode: payMode, ref: $('p_ref').value, notes: $('p_notes').value };
  var b = $('p_save'), label = editPay ? 'Update payment' : 'Save payment', wasEdit = !!editPay, back = backToList;
  var amt = num('p_amount'), ex = num('p_expense');
  function send(force){
    d.force = force;
    busy(b, true, label);
    google.script.run
      .withSuccessHandler(function(res){
        busy(b, false, label);
        if (res && res.dup) {
          askBox(b.parentNode, 'This payment looks already saved: ' + res.info + '. Save it again?',
            [{ t: 'Save anyway', fn: function(){ send(true); } }, { t: 'Cancel', cls: 'alt' }]);
          return;
        }
        var txt = 'Payment ' + res.id + (wasEdit ? ' updated.' : ' saved.') + (amt > 0 ? ' ₹' + inr(amt) + ' received.' : '') + (ex > 0 ? ' Expense ₹' + inr(ex) + '.' : '');
        payReset(); refreshReports();
        if (back) { openId = d.tripId; go('trips', true); }
        reloadTrips();
        show(txt, 'ok');
      })
      .withFailureHandler(function(e){ show('Not saved: ' + errText(e), 'err'); busy(b, false, label); })
      .savePayment(d);
  }
  send(false);
};

/* ---------- trip list ---------- */
FILTERS.forEach(function(f){
  var b = document.createElement('button'); b.type = 'button'; b.dataset.f = f; b.textContent = f;
  b.onclick = function(){ listFilter = f; listLimit = 50; renderList(); };
  $('l_f').appendChild(b);
});
// Quick date buttons fill the From / To boxes; the boxes can also be changed by hand.
function monthRange(back){
  var y = +D.today.slice(0, 4), m = +D.today.slice(5, 7) - back;
  if (m < 1) { m += 12; y--; }
  var mm = (m < 10 ? '0' : '') + m, last = new Date(y, m, 0).getDate();
  return [y + '-' + mm + '-01', y + '-' + mm + '-' + (last < 10 ? '0' : '') + last];
}
function dateRange(f){ return f === 'This month' ? monthRange(0) : f === 'Last month' ? monthRange(1) : ['', '']; }
DATES.forEach(function(f){
  var b = document.createElement('button'); b.type = 'button'; b.dataset.d = f; b.textContent = f;
  b.onclick = function(){ var r = sRange(f); $('l_from').value = r[0]; $('l_to').value = r[1]; listLimit = 50; renderList(); };
  $('l_d').appendChild(b);
});
LBY.forEach(function(x){
  var b = document.createElement('button'); b.type = 'button'; b.dataset.k = x[0]; b.textContent = x[1];
  b.onclick = function(){ lBy = x[0]; $('l_which').value = ''; listLimit = 50; renderList(); };
  $('l_by').appendChild(b);
});
$('l_which').addEventListener('change', function(){ listLimit = 50; renderList(); });
function lName(t){ return lBy === 'transporter' ? (t.transporter || 'Self / Own vehicle') : t[lBy]; }
function lNames(t){ return lBy === 'driver' ? drvList(t.driver) : [lName(t)]; }
function lKey(v){ return lBy === 'vehicle' ? norm('vehicle', v) : norm('party', v); }
['l_from', 'l_to'].forEach(function(i){ $(i).addEventListener('change', function(){ listLimit = 50; renderList(); }); });
$('l_q').addEventListener('input', function(){ listLimit = 50; renderList(); });
$('l_refresh').onclick = function(){ fullReload($('l_refresh')); };
$('l_moreBtn').onclick = function(){ listLimit += 50; renderList(); };
function el(tag, cls, text){ var e = document.createElement(tag); if (cls) e.className = cls; if (text != null) e.textContent = text; return e; }
function renderList(){
  Array.prototype.forEach.call($('l_f').children, function(x){ x.className = x.dataset.f === listFilter ? 'on' : ''; });
  var df = $('l_from').value, dt = $('l_to').value;
  Array.prototype.forEach.call($('l_d').children, function(x){ var r = sRange(x.dataset.d); x.className = r[0] === df && r[1] === dt ? 'on' : ''; });
  var lb = LBY.filter(function(x){ return x[0] === lBy; })[0];
  Array.prototype.forEach.call($('l_by').children, function(x){ x.className = x.dataset.k === lBy ? 'on' : ''; });
  $('l_whichLbl').textContent = lb[1];
  var ws = $('l_which'), wcur = lPickInit || ws.value, seen = {}, names = [];
  lPickInit = '';
  D.trips.forEach(function(t){ lNames(t).forEach(function(n){ var k = lKey(n); if (k && !seen[k]) { seen[k] = n; names.push(k); } }); });
  ws.innerHTML = ''; opt(ws, '', 'All ' + lb[2]);
  names.sort(function(a, b){ return seen[a].localeCompare(seen[b]); }).forEach(function(k){ opt(ws, k, seen[k]); });
  ws.value = wcur && seen[wcur] ? wcur : '';
  var lw = ws.value;
  var q = $('l_q').value.trim().toLowerCase();
  var rows = D.trips.filter(function(t){
    if (lw && !lNames(t).some(function(n){ return lKey(n) === lw; })) return false;
    if (listFilter === 'Not cleared' ? t.status === 'Cleared' : listFilter !== 'All' && t.status !== listFilter) return false;
    if (df && String(t.date) < df) return false;
    if (dt && String(t.date) > dt) return false;
    if (!q) return true;
    return [t.id, t.party, t.vehicle, t.driver, t.from, t.to, t.via, t.transporter, t.commodity, fmtDate(t.date)].join(' ').toLowerCase().indexOf(q) >= 0;
  });
  var sf = 0, sr = 0;
  lRows = rows;
  rows.forEach(function(t){ sf += t.freight === '' ? 0 : t.freight; if (t.status === 'Due') sr += t.remaining; });
  $('l_tot').textContent = rows.length + ' trip' + (rows.length === 1 ? '' : 's') + ' · freight ₹' + inr(sf) + ' · remaining ₹' + inr(sr);
  var box = $('l_list'); box.innerHTML = '';
  if (!rows.length) box.appendChild(el('div', 'sub', 'No trips match.'));
  rows.slice(0, listLimit).forEach(function(t){ box.appendChild(card(t)); });
  $('l_moreBtn').style.display = rows.length > listLimit ? '' : 'none';
  if (curTab() === 'trips') saveState();
  if (openId) { var o = $('trip_' + openId); if (o) o.scrollIntoView({ block: 'center' }); openId = ''; }
  syncCombos();
}
function numBox(label, id, val){
  var w = el('div'); w.appendChild(el('label', '', label));
  var i = document.createElement('input'); i.type = 'number'; i.inputMode = 'decimal'; i.min = '0'; i.step = 'any'; i.id = id;
  i.value = val === '' || val == null ? '' : val; w.appendChild(i); return w;
}
function freightPanel(t){
  var p = el('div', 'fq'), mode = t.ftype === 'Fixed amount' ? 'Fixed amount' : 'Per quintal', k = t.id;
  p.onclick = function(ev){ ev.stopPropagation(); };
  p.appendChild(el('b', '', 'Add freight for ' + t.id));
  var seg = el('div', 'seg'), bq = el('button', '', 'Per quintal'), bf = el('button', '', 'Fixed amount');
  bq.type = bf.type = 'button'; seg.appendChild(bq); seg.appendChild(bf); p.appendChild(seg);
  var g = el('div', 'g');
  var wq = numBox('Weight (quintal)', 'fq_w_' + k, t.weight), rq = numBox('Rate ₹/quintal', 'fq_r_' + k, t.rate);
  var fx = numBox('Fixed freight ₹', 'fq_f_' + k, ''), av = numBox('Advance ₹', 'fq_a_' + k, t.advance);
  var outW = el('div'); outW.appendChild(el('label', '', 'Freight ₹ (auto)')); var out = el('div', 'out', '–'); outW.appendChild(out);
  [wq, rq, fx, av, outW].forEach(function(x){ g.appendChild(x); }); p.appendChild(g);
  var bal = el('div', 'bal'); bal.appendChild(el('span', '', 'Remaining')); var bv = el('b', '', '–'); bal.appendChild(bv); p.appendChild(bal);
  var row = el('div', 'row'), cn = el('button', 'btn2', 'Cancel'), sv = el('button', 'save', 'Save freight');
  cn.type = sv.type = 'button'; row.appendChild(cn); row.appendChild(sv); p.appendChild(row);
  function v(id){ var x = p.querySelector('#' + id).value.trim(); return x; }
  function fr(){
    if (mode === 'Fixed amount') return v('fq_f_' + k) === '' ? null : +v('fq_f_' + k);
    if (v('fq_w_' + k) === '' || v('fq_r_' + k) === '') return null;
    return Math.round(+v('fq_w_' + k) * +v('fq_r_' + k) * 100) / 100;
  }
  function calc(){
    var f = fr(); out.textContent = f === null ? '–' : inr(f);
    bv.textContent = f === null ? '–' : '₹' + inr(f - (+v('fq_a_' + k) || 0) - t.paid - t.expense);
  }
  function setMode(m){
    mode = m; var isF = m === 'Fixed amount';
    bq.className = isF ? '' : 'on'; bf.className = isF ? 'on' : '';
    wq.style.display = rq.style.display = isF ? 'none' : ''; fx.style.display = isF ? '' : 'none'; calc();
  }
  bq.onclick = function(){ setMode('Per quintal'); }; bf.onclick = function(){ setMode('Fixed amount'); };
  p.querySelectorAll('input').forEach(function(i){ i.addEventListener('input', function(){ i.classList.remove('bad'); calc(); }); });
  cn.onclick = function(){ p.remove(); };
  sv.onclick = function(){
    if (needWho()) return;
    var need = mode === 'Fixed amount' ? ['fq_f_' + k] : ['fq_w_' + k, 'fq_r_' + k];
    var miss = need.filter(function(id){ return v(id) === ''; });
    miss.forEach(function(id){ p.querySelector('#' + id).classList.add('bad'); });
    if (miss.length) return;
    var d = { by: WHO, id: t.id, date: t.date, vehicle: t.vehicle, driver: t.driver, party: t.party, from: t.from, to: t.to, via: t.via,
      commodity: t.commodity, transporter: t.transporter, phone: t.phone, notes: t.notes, type: mode,
      weight: mode === 'Fixed amount' ? t.weight : v('fq_w_' + k), rate: mode === 'Fixed amount' ? '' : v('fq_r_' + k),
      fixed: mode === 'Fixed amount' ? v('fq_f_' + k) : '', advance: v('fq_a_' + k) };
    var f = fr();
    busy(sv, true, 'Save freight');
    google.script.run
      .withSuccessHandler(function(){
        show('Freight ₹' + inr(f) + ' added to trip ' + t.id + '.', 'ok');
        openId = t.id; refreshReports(); reloadTrips();
      })
      .withFailureHandler(function(e){ busy(sv, false, 'Save freight'); show('Not saved: ' + errText(e), 'err'); })
      .saveTrip(d);
  };
  setMode(mode);
  return p;
}
function card(t){
  var c = el('div', 'trip'); c.id = 'trip_' + t.id;
  var l1 = el('div', 'l1');
  l1.appendChild(el('b', '', t.id + ' · ' + fmtDate(t.date)));
  l1.appendChild(el('span', 'pill ' + (t.status === 'Due' ? 'due' : t.status === 'Cleared' ? 'ok' : 'inc'), t.status));
  c.appendChild(l1);
  c.appendChild(el('div', '', t.party));
  c.appendChild(el('div', 'sub', t.from + (t.via ? ' → ' + t.via : '') + ' → ' + t.to + ' · ' + t.vehicle + ' · ' + t.driver + (t.transporter && t.transporter !== 'Self / Own vehicle' ? ' · ' + t.transporter : '')));
  c.appendChild(el('div', 'amts', 'Freight ' + money(t.freight) + ' · Adv ' + money(t.advance === '' ? 0 : t.advance) + ' · Paid ₹' + inr(t.paid) + ' · Exp ₹' + inr(t.expense)));
  c.appendChild(el('div', 'rem', 'Remaining ' + (t.freight === '' ? '– (freight not entered)' : '₹' + inr(t.remaining))));
  if (t.status === 'Freight not entered') {
    var ab = el('button', 'mini addf', '+ Add freight'); ab.type = 'button';
    ab.onclick = function(ev){ ev.stopPropagation(); var x = c.querySelector('.fq'); if (x) { x.remove(); return; } c.insertBefore(freightPanel(t), ab.nextSibling); };
    c.appendChild(ab);
  }
  var det = el('div', 'det'); det.style.display = openId === t.id ? 'block' : 'none';
  if (t.weight !== '' || t.commodity) det.appendChild(el('div', 'sub', (t.weight !== '' ? t.weight + ' quintal' + (t.rate !== '' ? ' × ₹' + t.rate : '') : '') + (t.commodity ? (t.weight !== '' ? ' · ' : '') + t.commodity : '')));
  det.appendChild(el('div', 'h3', t.pays.length ? 'Payments for this trip' : 'No payments yet for this trip'));
  t.pays.forEach(function(p){
    var r = el('div', 'prow');
    r.appendChild(el('span', '', fmtShort(p.date) + (p.amount ? ' · ₹' + inr(p.amount) + ' ' + (p.mode.indexOf('Bank') === 0 ? 'Bank' : p.mode) : '') + (p.expense ? ' · exp ₹' + inr(p.expense) : '') + (p.ref ? ' · ' + p.ref : '')));
    var e = el('button', 'mini', 'Edit'); e.type = 'button';
    e.onclick = function(ev){ ev.stopPropagation(); startEditPayment(p); };
    r.appendChild(e); det.appendChild(r);
  });
  if (t.enteredBy) det.appendChild(el('div', 'who', 'Added by ' + t.enteredBy + (t.created ? ' · ' + t.created : '')));
  if (t.updatedBy) det.appendChild(el('div', 'who', 'Changed by ' + t.updatedBy + (t.changed ? ' · ' + t.changed : '')));
  var row = el('div', 'row');
  var eb = el('button', 'btn2', 'Edit trip'); eb.type = 'button'; eb.onclick = function(ev){ ev.stopPropagation(); startEditTrip(t); };
  var pb = el('button', 'save', 'Add payment'); pb.type = 'button'; pb.onclick = function(ev){ ev.stopPropagation(); startPayment(t); };
  row.appendChild(eb); row.appendChild(pb); det.appendChild(row);
  c.appendChild(det);
  c.onclick = function(){ det.style.display = det.style.display === 'none' ? 'block' : 'none'; };
  return c;
}

/* ---------- master forms (Party, Vehicle, Driver, Transporter) ---------- */
function buildMaster(type){
  var def = D.defs[type], v = $('v_' + type);
  var c = document.createElement('div'); c.className = 'c'; v.appendChild(c);
  var hd = document.createElement('div');
  var h = document.createElement('h1'); h.textContent = def.label + ' master'; hd.appendChild(h);
  var sub = document.createElement('div'); sub.className = 'sub'; sub.textContent = 'Type or pick an existing ' + def.label.toLowerCase() + ' to edit it, or enter a new one.'; hd.appendChild(sub);
  c.appendChild(hd);
  var note = document.createElement('div'); note.className = 'note'; note.style.display = 'none'; note.id = 'note_' + type; c.appendChild(note);
  var g = document.createElement('div'); g.className = 'g'; c.appendChild(g);
  var mb = document.createElement('button'); mb.type = 'button'; mb.className = 'more';
  var mg = document.createElement('div'); mg.className = 'g'; mg.style.display = 'none';
  var moreNames = def.fields.filter(function(f){ return f.more; }).map(function(f){ return f.h.toLowerCase(); });
  var moreLabel = '+ More (' + moreNames.join(', ') + ')';
  mb.textContent = moreLabel;
  mb.onclick = function(){ var open = mg.style.display === 'none'; mg.style.display = open ? 'grid' : 'none'; mb.textContent = open ? '− Less' : moreLabel; };
  def.fields.forEach(function(f, i){
    var target = f.more ? mg : g;
    if (f.sec) { var s = document.createElement('div'); s.className = 'sect'; s.textContent = f.sec; target.appendChild(s); }
    target.appendChild(field(f, 'm_' + type + '_' + i));
  });
  if (moreNames.length) { c.appendChild(mb); c.appendChild(mg); }
  var dl = document.createElement('datalist'); dl.id = 'dl_' + type; c.appendChild(dl);
  var keyInp = $('m_' + type + '_0'); keyInp.setAttribute('list', 'dl_' + type);
  var row = document.createElement('div'); row.className = 'row';
  var clr = document.createElement('button'); clr.type = 'button'; clr.className = 'btn2'; clr.textContent = 'New';
  var save = document.createElement('button'); save.type = 'button'; save.className = 'save'; save.id = 'ms_' + type; save.textContent = 'Save ' + def.label.toLowerCase();
  row.appendChild(clr); row.appendChild(save); c.appendChild(row);

  var loaded = false;
  function label(){ return (loaded ? 'Update ' : 'Save ') + def.label.toLowerCase(); }
  function fillFrom(rec){
    def.fields.forEach(function(f, i){ if (i > 0) setVal($('m_' + type + '_' + i), rec ? rec[f.h] : ''); });
  }
  function reset(){ def.fields.forEach(function(f, i){ setVal($('m_' + type + '_' + i), ''); }); loaded = false; note.style.display = 'none'; save.textContent = label(); clearBad(v); }
  keyInp.addEventListener('input', function(){
    var rec = findRec(type, keyInp.value);
    if (rec) {
      fillFrom(rec); loaded = true; note.textContent = 'Editing existing ' + def.label.toLowerCase() + ' ' + rec[def.key] + '. Saving will update it.'; note.style.display = 'block';
    } else if (loaded) {
      loaded = false; note.style.display = 'none';
    }
    save.textContent = label();
  });
  clr.onclick = function(){ reset(); hideMsg(); keyInp.focus(); };
  save.onclick = function(){ trySave(false); };
  function trySave(force){
    if (needWho()) return;
    clearBad(v);
    var ids = []; def.fields.forEach(function(f, i){ if (f.req) ids.push('m_' + type + '_' + i); });
    if (markMissing(ids)) { show('Fill the boxes marked in red.', 'err'); return; }
    var sim = (force || loaded) ? [] : similarNames(type, keyInp.value);
    if (sim.length) {
      askBox(row, 'Did you mean ' + sim.join(' or ') + '? It is already saved.', sim.map(function(x){
        return { t: 'Edit ' + x, fn: function(){ keyInp.value = x; keyInp.dispatchEvent(new Event('input')); show('Loaded ' + x + '. Change what you need and tap Update.', 'ok'); } };
      }).concat([{ t: 'No, add as new', cls: 'alt', fn: function(){ trySave(true); } }]));
      return;
    }
    var d = { by: WHO }; def.fields.forEach(function(f, i){ d[f.h] = $('m_' + type + '_' + i).value; });
    var txt = label();
    busy(save, true, txt);
    google.script.run
      .withSuccessHandler(function(res){ upsertRec(type, res.record); show((res.updated ? 'Updated ' : 'Added ') + def.label.toLowerCase() + ' ' + res.name + '.', 'ok'); busy(save, false, txt); reset(); refreshReports(); })
      .withFailureHandler(function(e){ show('Not saved: ' + errText(e), 'err'); busy(save, false, txt); })
      .saveMaster(type, d);
  }
}



/* ---------- summary (totals with filters) ---------- */
var SBY = [['party','Party','parties'],['driver','Driver','drivers'],['vehicle','Vehicle','vehicles'],['transporter','Transporter','transporters']];
var SDATES = ['Last 7 days', 'Last 30 days', 'This month', 'Last month', 'Any date'];
var sBy = 'party';
function ymdStr(d){ var m = d.getMonth() + 1, x = d.getDate(); return d.getFullYear() + '-' + (m < 10 ? '0' : '') + m + '-' + (x < 10 ? '0' : '') + x; }
function daysBack(n){ var p = D.today.split('-'); return ymdStr(new Date(+p[0], +p[1] - 1, +p[2] - n)); }
function sRange(f){
  if (f === 'Last 7 days') return [daysBack(6), D.today];
  if (f === 'Last 30 days') return [daysBack(29), D.today];
  return dateRange(f);
}
function sName(t){ return sBy === 'transporter' ? (t.transporter || 'Self / Own vehicle') : t[sBy]; }
function sKey(v){ return sBy === 'vehicle' ? norm('vehicle', v) : norm('party', v); }
SBY.forEach(function(x){
  var b = document.createElement('button'); b.type = 'button'; b.dataset.k = x[0]; b.textContent = x[1];
  b.onclick = function(){ sBy = x[0]; $('s_which').value = ''; renderSummary(); };
  $('s_by').appendChild(b);
});
SDATES.forEach(function(f){
  var b = document.createElement('button'); b.type = 'button'; b.dataset.d = f; b.textContent = f;
  b.onclick = function(){ var r = sRange(f); $('s_from').value = r[0]; $('s_to').value = r[1]; renderSummary(); };
  $('s_d').appendChild(b);
});
['s_from', 's_to', 's_which'].forEach(function(i){ $(i).addEventListener('change', renderSummary); });
$('s_min').addEventListener('input', renderSummary);
$('s_refresh').onclick = function(){ fullReload($('s_refresh')); };
function kpi(label, value, main){ var k = el('div', 'kpi' + (main ? ' main' : '')); k.appendChild(el('span', '', label)); k.appendChild(el('b', '', value)); return k; }
function renderSummary(){
  if (!$('s_by').children.length) return;
  var sb = SBY.filter(function(x){ return x[0] === sBy; })[0], lbl = sb[1];
  Array.prototype.forEach.call($('s_by').children, function(x){ x.className = x.dataset.k === sBy ? 'on' : ''; });
  $('s_whichLbl').textContent = lbl;
  var df = $('s_from').value, dt = $('s_to').value;
  Array.prototype.forEach.call($('s_d').children, function(x){ var r = sRange(x.dataset.d); x.className = r[0] === df && r[1] === dt ? 'on' : ''; });
  var inDates = D.trips.filter(function(t){ return (!df || String(t.date) >= df) && (!dt || String(t.date) <= dt); });
  // names for the dropdown (only those with trips in the chosen dates)
  var groups = {}, order = [];
  inDates.forEach(function(t){
    (sBy === 'driver' ? drvList(t.driver) : [sName(t)]).forEach(function(n){
      var k = sKey(n);
      if (!k) return;
      if (!groups[k]) { groups[k] = { name: n, trips: [] }; order.push(k); }
      groups[k].trips.push(t);
    });
  });
  var sel = $('s_which'), cur = sPickInit || sel.value;
  sPickInit = '';
  sel.innerHTML = ''; opt(sel, '', 'All ' + sb[2]);
  order.slice().sort(function(a, b){ return groups[a].name.localeCompare(groups[b].name); }).forEach(function(k){ opt(sel, k, groups[k].name); });
  if (cur && groups[cur]) sel.value = cur; else sel.value = '';
  var pick = sel.value, rows = pick ? groups[pick].trips : inDates;
  function tot(list){
    var o = { n: list.length, w: 0, f: 0, a: 0, p: 0, e: 0, r: 0, nf: 0 };
    list.forEach(function(t){
      o.w += t.weight === '' ? 0 : +t.weight; o.f += t.freight === '' ? 0 : t.freight; o.a += t.advance === '' ? 0 : t.advance;
      o.p += t.paid; o.e += t.expense;
      if (t.freight === '') o.nf++; else o.r += t.remaining;
    });
    return o;
  }
  // Payments saved "on account" (not against a trip) reduce the party's pending. Counted when dated in these dates.
  var oaList = sBy === 'party' ? (D.onacc || []).filter(function(p){ return (!df || String(p.date) >= df) && (!dt || String(p.date) <= dt); }) : [];
  function oaFor(key){ var x = 0; oaList.forEach(function(p){ if (norm('party', p.party) === key) x += p.amount + p.expense; }); return x; }
  var oa = pick ? oaFor(pick) : order.reduce(function(x, kk){ return x + oaFor(kk); }, 0);
  var multi = sBy === 'driver' && !pick && inDates.some(function(t){ return drvList(t.driver).length > 1; });
  var o = tot(rows), k = $('s_kpi');
  k.innerHTML = '';
  [kpi('Trips', String(o.n)), kpi('Weight', qty(o.w) + ' qtl'), kpi('Freight', '₹' + inr(o.f)), kpi('Advance', '₹' + inr(o.a)),
   kpi('Paid', '₹' + inr(o.p)), kpi('Expense', '₹' + inr(o.e)), kpi(oa ? 'Remaining on trips' : 'Remaining', '₹' + inr(o.r), !oa), kpi('Freight not entered', o.nf + ' trip' + (o.nf === 1 ? '' : 's'))
  ].concat(oa ? [kpi('Paid on account', '₹' + inr(oa)), kpi('Net remaining', '₹' + inr(o.r - oa), true)] : [])
   .forEach(function(x){ k.appendChild(x); });
  $('s_note').textContent = (df || dt ? 'Trips dated ' + (df ? fmtDate(df) : 'start') + ' to ' + (dt ? fmtDate(dt) : 'today') + '. ' : 'All dates. ') +
    'Paid and expense are payments entered against these trips.' +
    (oa ? ' "Paid on account" is money received without choosing a trip, dated in these dates; net remaining = remaining on trips − paid on account.' : '') +
    (multi ? ' A trip with 2 or 3 drivers is counted under each driver, so the driver rows add up to more than the total.' : '');
  var box = $('s_list'); box.innerHTML = '';
  if (!pick) {
    $('s_listTitle').textContent = order.length ? 'By ' + lbl.toLowerCase() + ' (tap one for its trips)' : 'No trips in these dates.';
    order.map(function(kk){ var to = tot(groups[kk].trips), a = oaFor(kk); return { k: kk, g: groups[kk], o: to, oa: a, net: to.r - a }; })
      .sort(function(a, b){ return b.net - a.net || b.o.f - a.o.f; })
      .forEach(function(x){
        var r = el('div', 'srow');
        var l = el('div'); l.appendChild(el('b', '', x.g.name)); l.appendChild(el('small', '', x.o.n + ' trip' + (x.o.n === 1 ? '' : 's') + ' · freight ₹' + inr(x.o.f)));
        var rt = el('div', 'r'); rt.appendChild(el('b', '', '₹' + inr(x.net))); rt.appendChild(el('small', '', x.oa ? 'after ₹' + inr(x.oa) + ' on account' : 'remaining'));
        r.appendChild(l); r.appendChild(rt);
        r.onclick = function(){ $('s_which').value = x.k; renderSummary(); window.scrollTo(0, 0); };
        box.appendChild(r);
      });
  } else {
    $('s_listTitle').textContent = 'Trips of ' + groups[pick].name + ' (tap to open in Trip list)';
    rows.forEach(function(t){
      var r = el('div', 'srow');
      var l = el('div'); l.appendChild(el('b', '', t.id + ' · ' + fmtShort(t.date))); l.appendChild(el('small', '', t.from + ' → ' + t.to + ' · ' + (sBy === 'party' ? t.vehicle : t.party)));
      var rt = el('div', 'r'); rt.appendChild(el('b', '', t.freight === '' ? '–' : '₹' + inr(t.remaining))); rt.appendChild(el('small', '', t.status === 'Freight not entered' ? 'freight not entered' : t.status === 'Cleared' ? 'cleared' : 'remaining'));
      r.appendChild(l); r.appendChild(rt);
      r.onclick = function(){
        listFilter = 'All'; $('l_from').value = ''; $('l_to').value = ''; $('l_q').value = t.id; $('l_which').value = '';
        openId = t.id; go('trips', true);
      };
      box.appendChild(r);
    });
  }
  sView = { by: sBy, lbl: lbl, plural: sb[2], df: df, dt: dt, pick: pick, name: pick ? groups[pick].name : '', rows: rows, o: o, oa: oa, multi: multi,
    groups: order.map(function(kk){ var to = tot(groups[kk].trips), a = oaFor(kk); return { name: groups[kk].name, trips: groups[kk].trips, o: to, oa: a, net: to.r - a }; })
      .sort(function(a, b){ return b.net - a.net || b.o.f - a.o.f; }) };
  sActions(pick ? groups[pick].name : '', rows, df, dt, oa);
  syncCombos();
  if (curTab() === 'summary') saveState();
}

/* ---------- Print / PDF (Trip list and Summary) ---------- */
var sView = null;
function datesText(df, dt){ return (df || dt) ? (df ? fmtDate(df) : 'start') + ' – ' + (dt ? fmtDate(dt) : fmtDate(D.today)) : 'All dates'; }
function tripRowsForPrint(rows){
  return rows.map(function(t){
    return [fmtDate(t.date), t.id, t.party, t.from + ' → ' + t.to, t.vehicle, t.driver, t.freight === '' ? '' : t.freight,
      t.advance === '' ? '' : t.advance, t.paid, t.expense, t.freight === '' ? '' : t.remaining, t.status];
  });
}
function tripFoot(rows){
  var f = 0, a = 0, p = 0, e = 0, r = 0;
  rows.forEach(function(t){ f += t.freight === '' ? 0 : t.freight; a += t.advance === '' ? 0 : t.advance; p += t.paid; e += t.expense; if (t.status === 'Due') r += t.remaining; });
  return ['Total', rows.length + (rows.length === 1 ? ' trip' : ' trips'), '', '', '', '', f, a, p, e, r, ''];
}
var TRIP_HEAD = ['Date', 'Trip no.', 'Party', 'Route', 'Vehicle', 'Driver', 'Freight', 'Advance', 'Paid', 'Expense', 'Remaining', 'Status'];
function printPdf(btn, out, spec){
  if (!spec.rows.length) { out.textContent = 'Nothing to print for these filters.'; return; }
  btn.disabled = true; btn.textContent = 'Making PDF…'; out.textContent = '';
  google.script.run
    .withSuccessHandler(function(r){
      btn.disabled = false; btn.textContent = 'Print / PDF'; out.innerHTML = '';
      out.appendChild(document.createTextNode('PDF ready: '));
      var l = el('a', '', 'Open PDF'); l.href = r.url; l.target = '_blank'; l.rel = 'noopener'; out.appendChild(l);
      out.appendChild(document.createTextNode(' Print or share it from there. Saved in Drive → Transport Tracker → Reports.'));
    })
    .withFailureHandler(function(e){ btn.disabled = false; btn.textContent = 'Print / PDF'; out.textContent = 'PDF not made: ' + errText(e); })
    .makeReportPdf(spec);
}
$('l_print').onclick = function(){
  var lb = LBY.filter(function(x){ return x[0] === lBy; })[0], ws = $('l_which');
  var parts = ['Status: ' + listFilter];
  if (ws.value) parts.push(lb[1] + ': ' + ws.options[ws.selectedIndex].text);
  parts.push('Dates: ' + datesText($('l_from').value, $('l_to').value));
  if ($('l_q').value.trim()) parts.push('Search: "' + $('l_q').value.trim() + '"');
  printPdf($('l_print'), $('l_printOut'), { title: 'Trip list', subtitle: parts.join(' · ') + ' · Remaining total counts trips that are due', head: TRIP_HEAD,
    rows: tripRowsForPrint(lRows), foot: tripFoot(lRows), numCols: [6, 7, 8, 9, 10] });
};
function daysSince(ymd){ var x = String(ymd).split('-'), y = D.today.split('-'); return Math.round((Date.UTC(+y[0], +y[1] - 1, +y[2]) - Date.UTC(+x[0], +x[1] - 1, +x[2])) / 86400000); }
function daysPending(trips){ var d = ''; trips.forEach(function(t){ if (t.status === 'Due' && t.date && (!d || String(t.date) < d)) d = String(t.date); }); return d ? daysSince(d) : ''; }
// Party / transporter details from the master, for the top of the PDF.
function infoFor(type, name){
  if (type !== 'party' && type !== 'transporter') return [];
  var r = findRec(type, name) || {};
  var f = type === 'party' ? ['Contact Person', 'Phone', 'Address', 'City', 'GSTIN', 'Credit Days'] : ['Contact Person', 'Phone', 'Address', 'GSTIN / PAN'];
  var out = f.filter(function(h){ return r[h]; }).map(function(h){ return [h, r[h]]; });
  if (!r['Phone']) { var ph = phoneFor(name); if (ph) out.unshift(['Phone', ph]); }
  return out;
}
$('s_print').onclick = function(){
  var v = sView; if (!v) return;
  var sub = 'Summary by ' + v.lbl + ': ' + (v.pick ? v.name : 'all ' + v.plural) + ' · Dates: ' + datesText(v.df, v.dt);
  var o = v.o, notes = [], lines = [];
  if (o.nf) notes.push(o.nf + ' trip' + (o.nf === 1 ? ' has' : 's have') + ' no freight entered; ' + (o.nf === 1 ? 'it is' : 'they are') + ' not in the Freight and Remaining totals.');
  if (v.oa) notes.push('"Paid on account" is money received without choosing a trip, dated in these dates.');
  if (v.multi) notes.push('A trip with 2 or 3 drivers is counted under each driver, so the driver rows add up to more than the Total row.');
  if (v.pick) {
    var cols = [['Date', function(t){ return fmtDate(t.date); }]];
    if (v.by !== 'party') cols.push(['Party', function(t){ return t.party; }]);
    cols.push(['Route', function(t){ return t.from + ' → ' + t.to; }]);
    if (v.by !== 'vehicle') cols.push(['Vehicle', function(t){ return t.vehicle; }]);
    if (v.by !== 'driver') cols.push(['Driver', function(t){ return t.driver; }]);
    cols.push(['Weight (qtl)', function(t){ return t.weight === '' || t.weight == null ? '' : +t.weight; }, 'w']);
    cols.push(['Rate (₹/qtl)', function(t){ return t.ftype === 'Fixed amount' ? 'Fixed' : (t.rate === '' || t.rate == null ? '' : +t.rate); }, '-']);
    cols.push(['Freight', function(t){ return t.freight === '' ? '' : t.freight; }, 'f']);
    cols.push(['Advance', function(t){ return t.advance === '' ? '' : t.advance; }, 'a']);
    cols.push(['Paid', function(t){ return t.paid; }, 'p']);
    cols.push(['Expense', function(t){ return t.expense; }, 'e']);
    cols.push(['Remaining', function(t){ return t.freight === '' ? '' : t.remaining; }, 'r']);
    cols.push(['Status', function(t){ return t.status; }]);
    cols.push(['Notes', function(t){ return t.notes || ''; }]);
    var tots = { w: o.w, f: o.f, a: o.a, p: o.p, e: o.e, r: o.r };
    var numCols = []; cols.forEach(function(c, i){ if (c[2]) numCols.push(i); });
    if (v.oa) lines = [['Remaining on trips', o.r], ['Less: paid on account', -v.oa], ['Net remaining', o.r - v.oa]];
    var min = num('s_min');
    if (min && (v.by === 'party' || v.by === 'transporter')) notes.push('"Only trips pending above ₹' + inr(min) + '" is used for the WhatsApp reminder and statement, not for this list.');
    printPdf($('s_print'), $('s_printOut'), { title: 'Summary – ' + v.name, subtitle: sub, info: infoFor(v.by, v.name),
      head: cols.map(function(c){ return c[0]; }),
      rows: v.rows.map(function(t){ return cols.map(function(c){ return c[1](t); }); }),
      foot: cols.map(function(c, i){ return i === 0 ? 'Total (' + o.n + ' trip' + (o.n === 1 ? '' : 's') + ')' : (c[2] && c[2] in tots ? tots[c[2]] : ''); }),
      numCols: numCols, lines: lines, notes: notes });
  } else {
    var isP = v.by === 'party';
    printPdf($('s_print'), $('s_printOut'), { title: 'Summary by ' + v.lbl, subtitle: sub,
      head: [v.lbl, 'Trips', 'Weight (qtl)', 'Freight', 'Advance', 'Paid', 'Expense', 'Remaining', 'Freight not entered']
        .concat(isP ? ['Paid on account', 'Net remaining', 'Phone', 'Days pending'] : []),
      rows: v.groups.map(function(g){ return [g.name, g.o.n, g.o.w, g.o.f, g.o.a, g.o.p, g.o.e, g.o.r, g.o.nf]
        .concat(isP ? [g.oa, g.net, phoneFor(g.name), daysPending(g.trips)] : []); }),
      foot: ['Total', o.n, o.w, o.f, o.a, o.p, o.e, o.r, o.nf].concat(isP ? [v.oa, o.r - v.oa, '', ''] : []),
      numCols: [1, 2, 3, 4, 5, 6, 7, 8].concat(isP ? [9, 10, 12] : []), notes: notes });
  }
};

/* ---------- type-to-search pickers ----------
   The original <select> stays (hidden) and keeps the value, so every other part of the form works
   unchanged. The search box shows only names that contain what is typed, e.g. "ra" → Ramesh, Suraj. */
var COMBOS = [];
function makeCombo(id, allowEmpty){
  var sel = $(id); sel.style.display = 'none';
  var wrap = el('div', 'combo'); sel.parentNode.insertBefore(wrap, sel.nextSibling); sel._combo = wrap;
  var inp = document.createElement('input'); inp.type = 'text'; inp.autocomplete = 'off'; inp.id = id + '_q';
  inp.placeholder = 'Type to search'; wrap.appendChild(inp);
  var list = el('div', 'comboList'); list.style.display = 'none'; wrap.appendChild(list);
  function label(){
    var o = sel.options[sel.selectedIndex];
    if (!o || o.value === NEW || (!o.value && !allowEmpty)) return '';
    return o.text;
  }
  function sync(){
    if (document.activeElement !== inp) inp.value = label();
    inp.classList.toggle('bad', sel.classList.contains('bad'));
  }
  function pick(v, typed){
    list.style.display = 'none'; sel.value = v; sel.classList.remove('bad'); inp.classList.remove('bad');
    inp.value = label(); inp.blur();
    sel.dispatchEvent(new Event('change'));
    if (v === NEW && typed && $('qa_0')) { $('qa_0').value = typed; $('qa_0').dispatchEvent(new Event('change')); }
  }
  function render(){
    var raw = inp.value.trim(), q = raw.toLowerCase();
    if (q === label().toLowerCase()) q = '';
    list.innerHTML = '';
    var n = 0, exact = false, addOpt = null;
    Array.prototype.forEach.call(sel.options, function(o){
      if (o.value === NEW) { addOpt = o; return; }
      if (!o.value && !allowEmpty) return;
      var t = o.text.toLowerCase();
      if (q && t.indexOf(q) < 0) return;
      if (t === q) exact = true;
      n++;
      if (n > 60) return;
      var d = el('div', o.value === sel.value ? 'on' : '', o.text);
      d.addEventListener('mousedown', function(e){ e.preventDefault(); pick(o.value); });
      list.appendChild(d);
    });
    if (n > 60) list.appendChild(el('div', 'none', (n - 60) + ' more. Keep typing to narrow the list.'));
    if (!n) list.appendChild(el('div', 'none', q ? 'No match for "' + raw + '"' : 'Nothing to choose yet'));
    if (addOpt && !exact) {
      var typed = q ? raw : '';
      var a = el('div', 'add', addOpt.text + (typed ? ' "' + typed + '"' : ''));
      a.addEventListener('mousedown', function(e){ e.preventDefault(); pick(NEW, typed); });
      list.appendChild(a);
    }
    list.style.display = 'block';
  }
  inp.addEventListener('focus', function(){ inp.select(); render(); });
  inp.addEventListener('input', render);
  inp.addEventListener('keydown', function(e){
    if (e.key === 'Enter') { e.preventDefault(); var f = list.querySelector('div:not(.none)'); if (f) f.dispatchEvent(new Event('mousedown')); }
    if (e.key === 'Escape') inp.blur();
  });
  inp.addEventListener('blur', function(){
    setTimeout(function(){
      if (list.style.display === 'none') return;
      list.style.display = 'none';
      var q = inp.value.trim().toLowerCase();
      if (q && q !== label().toLowerCase()) {   // typed a full name without tapping it: take it if it matches exactly
        var hit = Array.prototype.filter.call(sel.options, function(o){ return o.value && o.value !== NEW && o.text.toLowerCase() === q; })[0];
        if (hit) { pick(hit.value); return; }
      }
      inp.value = label();
    }, 150);
  });
  COMBOS.push(sync); sync();
}
function syncCombos(){ (COMBOS || []).forEach(function(f){ f(); }); }
setInterval(syncCombos, 300);

/* ---------- WhatsApp reminder and statement (Summary, one party or transporter) ---------- */
var sCur = null;
function waNum(p){
  var d = String(p || '').replace(/[^0-9]/g, '');
  if (d.length === 11 && d.charAt(0) === '0') d = d.slice(1);
  if (d.length === 10) d = (D.cc || '91') + d;
  return d.length >= 11 ? d : '';
}
function phoneFor(name){
  var r = findRec(sBy, name);
  if (r && r['Phone']) return r['Phone'];
  if (sBy === 'party') {
    var t = D.trips.filter(function(x){ return norm('party', x.party) === norm('party', name) && x.phone; })[0];
    return t ? t.phone : '';
  }
  return '';
}
function periodText(df, dt){ return (df || dt) ? (df ? fmtDate(df) : 'start') + ' – ' + (dt ? fmtDate(dt) : fmtDate(D.today)) : ''; }
// WhatsApp reminder: Hindi first, then the same message in English (same trips, same order, same total).
// Each trip shows what the party recognises: date, route, vehicle, commodity, weight, rate (or "fixed"),
// freight and what is still pending. No internal trip number.
function remindText(name, rows, df, dt, oa){
  var tot = 0; rows.forEach(function(t){ tot += t.remaining; });
  oa = oa || 0;
  var per = periodText(df, dt), nl = String.fromCharCode(10), one = rows.length === 1;
  function rateNum(r){ return Number(r).toLocaleString('en-IN', { maximumFractionDigits: 2 }); }
  function line(t, hi){
    var hasW = t.weight !== '' && t.weight != null, hasR = t.rate !== '' && t.rate != null;
    return '• ' + fmtShort(t.date) + ' · ' + t.from + ' → ' + t.to + ' · ' + t.vehicle +
      (t.commodity ? ' · ' + t.commodity : '') +
      (hasW ? ' · ' + t.weight + (hi ? ' क्विंटल (qtl)' : ' qtl') : '') +
      (t.ftype === 'Fixed amount' ? (hi ? ' · तय भाड़ा (fixed)' : ' · fixed freight') :
        (hasR ? (hi ? ' · दर (rate) ₹' + rateNum(t.rate) + '/क्विंटल' : ' · rate ₹' + rateNum(t.rate) + '/qtl') : '')) +
      (hi ? ' · भाड़ा (freight) ₹' : ' · freight ₹') + inr(t.freight) +
      (hi ? ' · बाकी (pending) ₹' : ' · pending ₹') + inr(t.remaining);
  }
  var hi = 'प्रिय ' + name + ',' + nl +
    (one ? 'इस ट्रिप का भाड़ा बाकी है:' : 'इन ' + rows.length + ' ट्रिप का भाड़ा बाकी है:') + nl +
    rows.map(function(t){ return line(t, true); }).join(nl) + nl +
    (oa ? 'खाते में मिला भुगतान (on account): -₹' + inr(oa) + nl : '') +
    'कुल बकाया: ₹' + inr(tot - oa) + (per ? ' (' + per + ')' : '') + '।' + nl +
    'कृपया जल्द भुगतान करें। धन्यवाद।';
  var en = 'Dear ' + name + ',' + nl +
    (one ? 'Freight is pending for this trip:' : 'Freight is pending for these ' + rows.length + ' trips:') + nl +
    rows.map(function(t){ return line(t, false); }).join(nl) + nl +
    (oa ? 'Less payment received on account: ₹' + inr(oa) + nl : '') +
    'Total pending: ₹' + inr(tot - oa) + (per ? ' (' + per + ')' : '') + '.' + nl +
    'Kindly arrange the payment. Thank you.';
  return hi + nl + nl + en;
}
function sActions(name, rows, df, dt, oa){
  oa = sBy === 'party' ? (oa || 0) : 0;
  var box = $('s_act');
  if (!name || (sBy !== 'party' && sBy !== 'transporter') || norm('party', name) === 'self / own vehicle') { box.style.display = 'none'; sCur = null; return; }
  var min = num('s_min');
  var due = rows.filter(function(t){ return t.status === 'Due' && t.remaining >= min; })
    .sort(function(a, b){ return a.date < b.date ? -1 : a.date > b.date ? 1 : 0; });
  sCur = { name: name, rows: due, df: df, dt: dt, min: min, oa: oa };
  box.style.display = 'flex';
  var a = $('s_wa'), ph = waNum(phoneFor(name));
  if (!due.length) {
    a.className = 'save wa off'; a.removeAttribute('href'); a.textContent = 'No pending trips';
    $('s_actNote').textContent = 'Nothing pending' + (min ? ' above ₹' + inr(min) : '') + ' for ' + name + ' in these dates.';
    $('s_stmtBtn').disabled = true; return;
  }
  $('s_stmtBtn').disabled = false;
  a.className = 'save wa'; a.textContent = 'WhatsApp reminder';
  a.href = 'https://wa.me/' + ph + '?text=' + encodeURIComponent(remindText(name, due, df, dt, oa));
  var tot = 0; due.forEach(function(t){ tot += t.remaining; });
  $('s_actNote').textContent = due.length + ' pending trip' + (due.length === 1 ? '' : 's') + ', ₹' + inr(tot) +
    (oa ? ' less ₹' + inr(oa) + ' paid on account = ₹' + inr(tot - oa) : '') + ' in total. ' +
    (ph ? 'Opens WhatsApp for ' + name + ' (' + ph.slice(-10) + ') with the message in English and Hindi.' :
      'No phone number saved for ' + name + ', so WhatsApp will ask you to pick the contact.');
}
$('s_stmtBtn').onclick = function(){ if (sCur && sCur.rows.length) showStatement(); };
function showStatement(){
  var c = sCur, b = $('s_stmt'); b.innerHTML = '';
  var back = el('button', 'more', '← Back to Summary'); back.type = 'button';
  back.onclick = function(){ b.style.display = 'none'; $('s_main').style.display = 'flex'; window.scrollTo(0, 0); };
  b.appendChild(back);
  b.appendChild(el('h1', '', D.biz + ' — Statement'));
  b.appendChild(el('div', 'sub', (sBy === 'party' ? 'Party: ' : 'Transporter: ') + c.name + (periodText(c.df, c.dt) ? ' · ' + periodText(c.df, c.dt) : '') +
    ' · pending trips only' + (c.min ? ' (above ₹' + inr(c.min) + ')' : '')));
  var st = el('div', 'st'), tot = 0;
  c.rows.forEach(function(t){
    tot += t.remaining;
    var l = el('div', 'sl'), a = el('div');
    a.appendChild(el('b', '', fmtShort(t.date) + ' · ' + t.id));
    a.appendChild(el('small', '', t.from + ' → ' + t.to + ' · ' + t.vehicle));
    a.appendChild(el('small', '', 'Freight ₹' + inr(t.freight) + ' − adv ₹' + inr(t.advance || 0) + ' − paid ₹' + inr(t.paid) + ' − exp ₹' + inr(t.expense)));
    l.appendChild(a); l.appendChild(el('b', '', '₹' + inr(t.remaining)));
    st.appendChild(l);
  });
  var tl = el('div', 'tot'); tl.appendChild(el('span', '', c.oa ? 'Total pending on trips' : 'Total pending')); tl.appendChild(el('span', '', '₹' + inr(tot)));
  st.appendChild(tl);
  if (c.oa) {
    var t2 = el('div', 'sl'); t2.appendChild(el('span', '', 'Less: paid on account')); t2.appendChild(el('span', '', '−₹' + inr(c.oa))); st.appendChild(t2);
    var t3 = el('div', 'tot'); t3.appendChild(el('span', '', 'Net pending')); t3.appendChild(el('span', '', '₹' + inr(tot - c.oa))); st.appendChild(t3);
  }
  b.appendChild(st);
  var row = el('div', 'row');
  var pdf = el('button', 'btn2', 'Save as PDF'); pdf.type = 'button';
  var wa = el('a', 'save wa', 'Send on WhatsApp'); wa.target = '_blank'; wa.rel = 'noopener'; wa.href = $('s_wa').href;
  row.appendChild(pdf); row.appendChild(wa); b.appendChild(row);
  var out = el('div', 'sub'); b.appendChild(out);
  pdf.onclick = function(){
    busy(pdf, true, 'Save as PDF'); out.textContent = 'Making the PDF…';
    google.script.run
      .withSuccessHandler(function(r){
        busy(pdf, false, 'Save as PDF'); out.innerHTML = '';
        out.appendChild(document.createTextNode('Saved in Drive → Transport Tracker → Party Statements. '));
        var l = el('a', '', 'Open PDF'); l.href = r.url; l.target = '_blank'; l.rel = 'noopener'; out.appendChild(l);
        out.appendChild(document.createTextNode(' (then Share → WhatsApp to send the file)'));
      })
      .withFailureHandler(function(e){ busy(pdf, false, 'Save as PDF'); out.textContent = 'PDF not made: ' + errText(e); })
      .makeStatementPdf({ by: sBy, name: c.name, from: c.df, to: c.dt, min: c.min });
  };
  $('s_main').style.display = 'none'; b.style.display = 'flex'; window.scrollTo(0, 0);
}
(function(){ var r = sRange('This month'); $('s_from').value = r[0]; $('s_to').value = r[1]; })();

/* ---------- start ---------- */
$('biz').textContent = D.biz + ' · forms ' + D.ver;
$('whoSave').onclick = function(){
  var v = (D.staff && D.staff.length) ? $('whoSel').value : pc($('whoIn').value);
  if (!v) { $('whoSel').classList.add('bad'); $('whoIn').classList.add('bad'); return; }
  WHO = v; try { localStorage.setItem('tt_name', v); } catch (e) {}
  showWho(); hideMsg();
};
['t_party','t_vehicle','t_driver','t_driver2','t_driver3','t_transporter','p_party','whoSel'].forEach(function(i){ makeCombo(i, false); });
['s_which','l_which'].forEach(function(i){ makeCombo(i, true); });
showWho();
['party','vehicle','driver','transporter'].forEach(buildMaster);
['party','vehicle','driver','transporter'].forEach(refreshType);
fillList('locs', D.locations); fillList('comms', D.commodities);
tripReset(); payReset();
['t_from','t_to','t_via','t_commodity'].forEach(function(i){ $(i).addEventListener('change', function(){ $(i).value = pc($(i).value); }); });
document.querySelectorAll('input,select,textarea').forEach(function(e){ e.addEventListener('input', function(){ e.classList.remove('bad'); }); });
(function(){
  var q = D.qs || {};
  if (D.start === 'summary') {
    if (SBY.some(function(x){ return x[0] === q.by; })) sBy = q.by;
    if ('from' in q) $('s_from').value = q.from;
    if ('to' in q) $('s_to').value = q.to;
    if (q.min) $('s_min').value = q.min;
    sPickInit = q.which || '';
  }
  if (D.start === 'trips') {
    if (FILTERS.indexOf(q.lf) >= 0) listFilter = q.lf;
    if (LBY.some(function(x){ return x[0] === q.lby; })) lBy = q.lby;
    if ('lfrom' in q) $('l_from').value = q.lfrom;
    if ('lto' in q) $('l_to').value = q.lto;
    if (q.q) $('l_q').value = q.q;
    lPickInit = q.lwhich || '';
  }
})();
go(D.start || 'trip', true);
</script></body></html>`;

/* =====================================================================
 * DROPDOWNS (master data → Trip & Payment forms)
 * ===================================================================== */
function refreshDropdowns() {
  if (prop_('APP_ON')) return;   // mobile forms read the master tabs directly
  const lists = {
    party: uniq_(readTable_(S.PARTIES).map(r => r['Party Name'])),
    vehicle: uniq_(readTable_(S.VEHICLES).map(r => normVehicle_(r['Vehicle Number']))),
    driver: uniq_(readTable_(S.DRIVERS).map(r => r['Driver Name'])),
    transporter: ['Self / Own vehicle'].concat(uniq_(readTable_(S.TRANSPORTERS).map(r => r['Transporter Name'])))
  };
  [FORMS.TRIP, FORMS.PAYMENT].forEach(def => {
    const form = FormApp.openById(prop_(def.key));
    const listItems = form.getItems(FormApp.ItemType.LIST);
    def.items.filter(s => s.type === 'dropdown').forEach(s => {
      const item = listItems.find(i => i.getTitle() === s.t);
      if (!item) return;
      const vals = lists[s.source];
      item.asListItem().setChoiceValues(vals.length ? vals : [PLACEHOLDER]);
    });
  });
}

/* =====================================================================
 * REPORTS
 * ===================================================================== */
const PARTY_HEAD = ['Party Name', 'Phone', 'Trips', 'Total Freight', 'Advance Received', 'Payments Received',
  'Expenses', 'Pending Amount', 'Last Trip', 'Oldest Unpaid Since', 'Days Pending', 'Status', 'WhatsApp Reminder'];
const TRIP_SUM_HEAD = ['Trip ID', 'Trip Date', 'Vehicle', 'Driver', 'Party', 'Transporter', 'Route', 'Weight (Quintal)',
  'Freight', 'Advance', 'Paid', 'Expense', 'Remaining', 'Status'];

function buildReports() {
  // Step 1 (short, shares the lock with saving): give IDs, add missing masters and read all data.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let data, master;
  try {
    data = tripData_();
    syncMasters_(data.trips);          // add any party / driver / vehicle / transporter missing from its master
    master = partyMaster_();
  } finally {
    lock.releaseLock();
  }
  // Step 2 (longer): write the report tabs. Uses its own lock, so saving a trip, payment or "+ Add new"
  // never has to wait for the reports to finish.
  const rlock = LockService.getDocumentLock();
  if (rlock) rlock.waitLock(30000);
  try {
    const now = new Date();
    const partyRows = partySummary_(data.trips, data.pays, master, now);
    writePartySummary_(partyRows);
    writeTripSummary_(data.trips);
    if (!prop_('CUSTOM_V3')) {          // one-time switch of the Custom Report to the Trip Summary tab
      setupCustomReport_();
      try {
        const ss = ss_(), ts = ss.getSheetByName(S.TRIP_SUM);
        ss.setActiveSheet(ts);
        ss.moveActiveSheet(Math.min(3, ss.getSheets().length));
        ts.setTabColor('#2e7d32');
      } catch (e) { /* tab order is only cosmetic */ }
      PropertiesService.getScriptProperties().setProperty('CUSTOM_V3', '1');
    }
    if (!prop_('STAFF_TAB')) {          // one-time: Staff tab for office names shown in "Who are you?"
      const st = getSheet_(S.STAFF);
      ensureHeaders_(st, STAFF_HEADERS);
      try { st.setTabColor('#ef6c00'); } catch (e) { /* cosmetic */ }
      PropertiesService.getScriptProperties().setProperty('STAFF_TAB', '1');
    }
    if (!prop_('LISTS_FILLED')) {       // one-time: put places and commodities from older trips into their masters
      addLocations_([].concat.apply([], data.trips.map(t => [t.from, t.via, t.to])));
      addToList_(S.COMMODITIES, COM_HEADERS, 'Commodity', data.trips.map(t => t.commodity));
      PropertiesService.getScriptProperties().setProperty('LISTS_FILLED', '1');
    }
    writeGroup_(S.VEH_SUM, 'Vehicle Number', groupBy_(data.trips, 'vehicle'), true);
    writeGroup_(S.DRV_SUM, 'Driver Name', groupBy_(data.trips, 'driver'), false);
    writeMonthly_(data.trips, data.pays);
    writeDashboard_(partyRows, data.trips, data.pays, now);
    refreshReportFilters_(data.trips);
    if (!prop_('CUSTOM_V5')) {          // one-time: Custom Report with the two-driver and Transporter filters
      setupCustomReport_();
      PropertiesService.getScriptProperties().setProperty('CUSTOM_V5', '1');
      refreshReportFilters_(data.trips);
    }
  } finally {
    if (rlock) rlock.releaseLock();
  }
}

function partySummary_(trips, pays, master, now) {
  const g = {};
  const get = name => {
    const k = key_(name);
    if (!g[k]) g[k] = { name: str_(name), trips: [], received: 0, expense: 0, onAcct: 0 };
    return g[k];
  };
  trips.forEach(t => { if (t.party) get(t.party).trips.push(t); });
  pays.forEach(p => {
    if (!p.party) return;
    const o = get(p.party);
    o.received += p.amount;
    o.expense += p.expense;
    if (!p.linked) o.onAcct += p.amount + p.expense;
  });

  return Object.keys(g).map(k => {
    const p = g[k], m = master[k] || {};
    const freight = sum_(p.trips, 'freight'), adv = sum_(p.trips, 'advance'), bal = sum_(p.trips, 'balance');
    const pending = bal - p.received - p.expense;

    // Trip-wise payments are already in each trip's "remaining"; payments not tied to a trip
    // are applied to the oldest trips first. The first trip still not covered is the oldest unpaid.
    const sorted = p.trips.slice().sort((a, b) => (a.date || 0) - (b.date || 0));
    let left = p.onAcct, oldest = null;
    for (const t of sorted) {
      if (t.remaining <= 0) continue;
      if (left >= t.remaining) { left -= t.remaining; continue; }
      oldest = t.date;
      break;
    }
    const days = pending > 0 && oldest ? Math.floor((now - oldest) / 86400000) : '';
    const credit = num_(m['Credit Days']) || CONFIG.DEFAULT_CREDIT_DAYS;
    const status = pending <= 0 ? 'Cleared' : (days !== '' && days > credit ? 'Overdue' : 'Pending');
    const phone = str_(m['Phone']) || (sorted.slice().reverse().find(t => t.phone) || {}).phone || '';
    const last = sorted.length ? sorted[sorted.length - 1].date : null;
    const wa = pending > 0 ? waLink_(phone, reminderText_(p.name, pending, oldest)) : '';
    return [p.name, phone, p.trips.length, freight, adv, p.received, p.expense, pending,
      last || '', (pending > 0 && oldest) ? oldest : '', days, status, wa];
  }).sort((a, b) => b[7] - a[7]);
}

function writePartySummary_(rows) {
  const sh = writeTable_(S.PARTY_SUM, PARTY_HEAD, rows,
    { 4: CUR, 5: CUR, 6: CUR, 7: CUR, 8: CUR, 9: DATE_F, 10: DATE_F });
  if (!rows.length) return;
  sh.getRange(2, 12, rows.length, 1).setBackgrounds(rows.map(r =>
    [r[11] === 'Overdue' ? '#f8d7da' : r[11] === 'Pending' ? '#fff3cd' : '#d4edda']));
  sh.getRange(2, 8, rows.length, 1).setFontWeight('bold');
}

function writeTripSummary_(trips) {
  const rows = trips.slice()
    .sort((a, b) => (b.date || 0) - (a.date || 0) || (b.id < a.id ? -1 : 1))
    .map(t => [t.id, t.date || '', t.vehicle, t.driver, t.party, t.transporter || 'Self / Own vehicle',
      t.from + (t.via ? ' → ' + t.via : '') + ' → ' + t.to, t.weightRaw, t.freightRaw, t.advanceRaw,
      t.paid || '', t.expense || '', t.freightRaw === '' && !t.paid && !t.expense && t.advanceRaw === '' ? '' : t.remaining, t.status]);
  const sh = writeTable_(S.TRIP_SUM, TRIP_SUM_HEAD, rows, { 2: DATE_F, 9: CUR, 10: CUR, 11: CUR, 12: CUR, 13: CUR });
  if (!rows.length) return;
  sh.getRange(2, 14, rows.length, 1).setBackgrounds(rows.map(r =>
    [r[13] === 'Due' ? '#fff3cd' : r[13] === 'Cleared' ? '#d4edda' : '#e2e3e5']));
  sh.getRange(2, 13, rows.length, 1).setFontWeight('bold');
}

// Groups case-insensitively, so "Deepak", "DEEPAK" and "deepak" are one row (first spelling is shown).
// Groups case-insensitively. For drivers, a trip with two drivers counts under each of them.
function groupBy_(trips, field) {
  const g = {};
  trips.forEach(t => {
    const names = field === 'driver' ? drivers_(t.driver) : [t[field]];
    (names.length ? names : ['(blank)']).forEach(n => {
      const name = n || '(blank)';
      const k = field === 'vehicle' ? name : key_(name);
      if (!g[k]) g[k] = { name: name, trips: 0, weight: 0, freight: 0, adv: 0, bal: 0, last: null };
      const o = g[k];
      o.trips++; o.weight += t.weight; o.freight += t.freight; o.adv += t.advance; o.bal += t.remaining;
      if (t.date && (!o.last || t.date > o.last)) o.last = t.date;
    });
  });
  return Object.keys(g).map(k => g[k]).sort((a, b) => b.bal - a.bal);
}

function writeGroup_(sheet, label, groups, withWeight) {
  const head = [label, 'Trips'].concat(withWeight ? ['Total Weight (Quintal)'] : [])
    .concat(['Total Freight', 'Advance Received', 'Remaining (after payments and expenses)', 'Last Trip']);
  const rows = groups.map(o => [o.name, o.trips].concat(withWeight ? [o.weight] : [])
    .concat([o.freight, o.adv, o.bal, o.last || '']));
  const off = withWeight ? 1 : 0, fm = {};
  fm[3 + off] = CUR; fm[4 + off] = CUR; fm[5 + off] = CUR; fm[6 + off] = DATE_F;
  writeTable_(sheet, head, rows, fm);
}

function writeMonthly_(trips, pays) {
  const tz = CONFIG.TIMEZONE, m = {};
  const get = d => {
    const k = Utilities.formatDate(d, tz, 'yyyy-MM');
    if (!m[k]) {
      const parts = k.split('-');
      m[k] = { k: k, month: new Date(+parts[0], +parts[1] - 1, 15), trips: 0, freight: 0, adv: 0, recv: 0, exp: 0, byParty: {} };
    }
    return m[k];
  };
  trips.forEach(t => {
    if (!t.date) return;
    const o = get(t.date), pk = key_(t.party);
    o.trips++; o.freight += t.freight; o.adv += t.advance;
    if (pk) {
      if (!o.byParty[pk]) o.byParty[pk] = { name: t.party, bal: 0 };
      o.byParty[pk].bal += t.balance;
    }
  });
  pays.forEach(p => { if (p.date) { const o = get(p.date); o.recv += p.amount; o.exp += p.expense; } });

  const rows = Object.keys(m).sort().reverse().map(k => {
    const o = m[k];
    const top = Object.keys(o.byParty).sort((a, b) => o.byParty[b].bal - o.byParty[a].bal)[0];
    return [o.month, o.trips, o.freight, o.adv, o.recv, o.exp, o.freight - o.adv - o.recv - o.exp,
      top ? o.byParty[top].name : '', top ? o.byParty[top].bal : ''];
  });
  writeTable_(S.MONTHLY,
    ['Month', 'Trips', 'Total Freight', 'Advance Received', 'Payments Received', 'Expenses',
      'Net Added to Pending', 'Party with Biggest Balance', 'Their Balance (freight − advance)'],
    rows, { 1: 'mmm yyyy', 3: CUR, 4: CUR, 5: CUR, 6: CUR, 7: CUR, 9: CUR });
}

function writeDashboard_(partyRows, trips, pays, now) {
  const tz = CONFIG.TIMEZONE, cur = Utilities.formatDate(now, tz, 'yyyy-MM');
  const inMonth = d => !!d && Utilities.formatDate(d, tz, 'yyyy-MM') === cur;
  const pending = partyRows.filter(r => r[7] > 0);
  const mTrips = trips.filter(t => inMonth(t.date));
  const mPays = pays.filter(p => inMonth(p.date)).reduce((s, p) => s + p.amount, 0);

  const sh = getSheet_(S.DASH);
  sh.clear();
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getRange('A1').setValue(CONFIG.BUSINESS_NAME + ' — Dashboard').setFontSize(16).setFontWeight('bold');
  sh.getRange('A2').setValue('Updated ' + Utilities.formatDate(now, tz, 'dd-MMM-yyyy HH:mm') + ' · code ' + CODE_VERSION +
    ' · last backup ' + (prop_('LAST_BACKUP') || 'not yet')).setFontColor('#666666');

  const kpi = [
    ['Total pending amount', pending.reduce((s, r) => s + r[7], 0)],
    ['Parties with pending', pending.length],
    ['Overdue parties (past credit days)', partyRows.filter(r => r[11] === 'Overdue').length],
    ['Trips this month', mTrips.length],
    ['Freight booked this month', sum_(mTrips, 'freight')],
    ['Collected this month (advance + payments)', sum_(mTrips, 'advance') + mPays],
    ['Trips with freight not entered yet', trips.filter(t => t.status === 'Freight not entered').length]
  ];
  sh.getRange(4, 1, kpi.length, 2).setValues(kpi);
  sh.getRange(4, 1, kpi.length, 1).setFontWeight('bold');
  sh.getRange(4, 2, kpi.length, 1).setFontSize(12);
  [4, 8, 9].forEach(r => sh.getRange(r, 2).setNumberFormat(CUR));

  let r = 13;
  sh.getRange(r, 1).setValue('Top 10 pending parties').setFontWeight('bold');
  sh.getRange(r + 1, 1, 1, 4).setValues([['Party', 'Pending', 'Days Pending', 'Status']])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  const top = pending.slice(0, 10).map(x => [x[0], x[7], x[10], x[11]]);
  if (top.length) {
    sh.getRange(r + 2, 1, top.length, 4).setValues(top);
    sh.getRange(r + 2, 2, top.length, 1).setNumberFormat(CUR);
    sh.insertChart(sh.newChart()
      .setChartType(Charts.ChartType.BAR)
      .addRange(sh.getRange(r + 1, 1, top.length + 1, 2))
      .setNumHeaders(1)
      .setPosition(3, 6, 0, 0)
      .setOption('title', 'Top pending parties (₹)')
      .setOption('legend', { position: 'none' })
      .setOption('width', 560).setOption('height', 340)
      .build());
  } else {
    sh.getRange(r + 2, 1).setValue('No pending amounts');
  }

  r = r + 3 + Math.max(top.length, 1);
  const exp = expiringDocs_(now);
  sh.getRange(r, 1).setValue('Documents expiring within ' + CONFIG.DOC_EXPIRY_ALERT_DAYS + ' days (negative = already expired)')
    .setFontWeight('bold');
  sh.getRange(r + 1, 1, 1, 4).setValues([['Vehicle / Driver', 'Document', 'Expiry Date', 'Days Left']])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  if (exp.length) {
    sh.getRange(r + 2, 1, exp.length, 4).setValues(exp);
    sh.getRange(r + 2, 3, exp.length, 1).setNumberFormat(DATE_F);
  } else {
    sh.getRange(r + 2, 1).setValue('None');
  }
  sh.setColumnWidth(1, 300);
  sh.autoResizeColumns(2, 3);
}

function refreshReportFilters_(trips) {
  const sh = ss_().getSheetByName(S.CUSTOM);
  if (!sh) return;
  setListValidation_(sh.getRange('B3'), uniq_(trips.map(t => t.party)));
  setListValidation_(sh.getRange('B4'), uniq_(trips.map(t => t.vehicle)));
  setListValidation_(sh.getRange('B5'), uniq_([].concat.apply([], trips.map(t => drivers_(t.driver)))));
  setListValidation_(sh.getRange('B6'), uniq_(trips.map(t => t.transporter || 'Self / Own vehicle')));
}

function setListValidation_(range, values) {
  range.setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['All'].concat(values), true).setAllowInvalid(false).build());
}

function expiringDocs_(now) {
  now = now || new Date();
  const out = [], lim = CONFIG.DOC_EXPIRY_ALERT_DAYS;
  const check = (who, doc, v) => {
    const d = toDate_(v);
    if (!d) return;
    const days = Math.ceil((d - now) / 86400000);
    if (days <= lim) out.push([who, doc, d, days]);
  };
  latestBy_(readTable_(S.VEHICLES), r => normVehicle_(r['Vehicle Number'])).forEach(v =>
    ['Insurance', 'Permit', 'Fitness', 'PUC'].forEach(doc => check(normVehicle_(v['Vehicle Number']), doc, v[doc + ' Expiry'])));
  latestBy_(readTable_(S.DRIVERS), r => key_(r['Driver Name'])).forEach(d =>
    check(str_(d['Driver Name']), 'Driving Licence', d['Licence Expiry']));
  return out.sort((a, b) => a[3] - b[3]);
}

/* =====================================================================
 * REMINDERS
 * ===================================================================== */
function reminderText_(party, pending, oldest) {
  return 'Dear ' + party + ', this is a reminder from ' + CONFIG.BUSINESS_NAME +
    '. Freight payment of Rs. ' + inr_(pending) + ' is pending' +
    (oldest ? ' (oldest unpaid trip dated ' + fmtD_(oldest) + ')' : '') +
    '. Kindly arrange the payment at the earliest. Thank you.';
}

function waLink_(phone, msg) {
  let d = str_(phone).replace(/\D/g, '');
  if (d.length === 11 && d[0] === '0') d = d.slice(1);
  if (d.length === 10) d = CONFIG.COUNTRY_CODE + d;
  if (d.length < 11) return '';
  return '=HYPERLINK("https://wa.me/' + d + '?text=' + encodeURIComponent(msg) + '","Send WhatsApp")';
}

// Menu: emails every pending party that has an Email in Party Master (asks before sending).
function sendPartyReminders() {
  const ui = SpreadsheetApp.getUi();
  buildReports();
  const master = partyMaster_();
  const list = readTable_(S.PARTY_SUM).filter(r =>
    num_(r['Pending Amount']) >= CONFIG.REMINDER_MIN_PENDING && str_((master[key_(r['Party Name'])] || {})['Email']));
  if (!list.length) {
    ui.alert('No parties with a pending amount AND an email address in Party Master.');
    return;
  }
  const preview = list.map(r => r['Party Name'] + ': Rs. ' + inr_(num_(r['Pending Amount']))).join('\n');
  if (ui.alert('Send email reminders to ' + list.length + ' parties?', preview, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  list.forEach(r => {
    const oldest = toDate_(r['Oldest Unpaid Since']);
    MailApp.sendEmail({
      to: str_(master[key_(r['Party Name'])]['Email']),
      subject: 'Payment reminder – ' + CONFIG.BUSINESS_NAME,
      body: reminderText_(r['Party Name'], num_(r['Pending Amount']), oldest),
      name: CONFIG.BUSINESS_NAME
    });
  });
  ui.alert('Sent ' + list.length + ' reminder(s).');
}

function sendDailySummary_() {
  const to = CONFIG.OWNER_EMAIL || Session.getEffectiveUser().getEmail();
  const pending = readTable_(S.PARTY_SUM).filter(r => num_(r['Pending Amount']) > 0);
  const exp = expiringDocs_();
  if (!pending.length && !exp.length) return;
  const total = pending.reduce((s, r) => s + num_(r['Pending Amount']), 0);
  let html = '<h3>' + esc_(CONFIG.BUSINESS_NAME) + ' — pending collections</h3>' +
    '<p>Total pending: <b>Rs. ' + inr_(total) + '</b> from ' + pending.length + ' parties.</p>' +
    htmlTable_(['Party', 'Phone', 'Pending (Rs.)', 'Days', 'Status'],
      pending.slice(0, 15).map(r => [r['Party Name'], r['Phone'], inr_(num_(r['Pending Amount'])), r['Days Pending'], r['Status']]));
  if (exp.length) {
    html += '<h3>Documents expiring within ' + CONFIG.DOC_EXPIRY_ALERT_DAYS + ' days</h3>' +
      htmlTable_(['Vehicle / Driver', 'Document', 'Expiry', 'Days left'], exp.map(x => [x[0], x[1], fmtD_(x[2]), x[3]]));
  }
  html += '<p><a href="' + ss_().getUrl() + '">Open the tracker</a></p>';
  MailApp.sendEmail({ to: to, subject: 'Pending Rs. ' + inr_(total) + ' — daily summary', htmlBody: html });
}

/* =====================================================================
 * BACKUP
 * ===================================================================== */
// Nightly: saves an Excel copy to Drive; on Sundays also emails it to BACKUP_EMAIL.
function backupNow() {
  const ss = ss_();
  const folder = getOrCreateFolder_(CONFIG.BACKUP_FOLDER);
  const now = new Date();
  const stamp = Utilities.formatDate(now, CONFIG.TIMEZONE, 'yyyy-MM-dd_HHmm');
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const blob = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() } })
    .getBlob().setName(ss.getName() + '_' + stamp + '.xlsx');
  folder.createFile(blob);
  PropertiesService.getScriptProperties().setProperty('LAST_BACKUP', Utilities.formatDate(now, CONFIG.TIMEZONE, 'dd-MMM-yyyy HH:mm'));

  const files = [];
  const it = folder.getFiles();
  while (it.hasNext()) files.push(it.next());
  files.sort((a, b) => b.getDateCreated() - a.getDateCreated())
    .slice(CONFIG.KEEP_BACKUPS).forEach(f => f.setTrashed(true));

  if (CONFIG.BACKUP_EMAIL && Utilities.formatDate(now, CONFIG.TIMEZONE, 'EEE') === 'Sun') {
    MailApp.sendEmail({
      to: CONFIG.BACKUP_EMAIL,
      subject: 'Weekly backup – ' + ss.getName() + ' – ' + stamp,
      body: 'Attached is the weekly backup of ' + ss.getName() + '. To restore: upload to Google Drive → right-click → Open with Google Sheets.',
      attachments: [blob]
    });
  }
  return blob.getName();
}

function backupNowUi() {
  alert_('Backup saved', backupNow() + '\nDrive folder: ' + CONFIG.BACKUP_FOLDER);
}

function openBackups() {
  const url = getOrCreateFolder_(CONFIG.BACKUP_FOLDER).getUrl();
  const last = prop_('LAST_BACKUP') || 'not yet';
  const html = '<div style="font-family:Arial,sans-serif;font-size:14px;line-height:1.6">' +
    '<p>A new Excel backup is saved every night at 10 pm. Each file name has its date, for example ' +
    '<b>…_2026-11-11_2200.xlsx</b> is the backup from 11 November. The last ' + CONFIG.KEEP_BACKUPS + ' days are kept.</p>' +
    '<p>Last backup: <b>' + esc_(last) + '</b></p>' +
    '<p><a href="' + url + '" target="_blank">Open the backups folder</a></p>' +
    '<p style="color:#666">To look at a backup: open the folder, double-click the file, then choose "Open with Google Sheets".</p></div>';
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(260), 'Daily backups');
}

/* =====================================================================
 * HELPERS
 * ===================================================================== */
function ss_() { return SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(prop_('SS_ID')); }
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function getSheet_(name) { return ss_().getSheetByName(name) || ss_().insertSheet(name); }

function getOrCreateFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  while (it.hasNext()) {
    const f = it.next();
    if (!f.isTrashed()) return f;   // ignore old folders sitting in Trash
  }
  return DriveApp.createFolder(name);
}

function readTable_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const h = v[0].map(x => String(x).trim());
  // Keep any row that has something in it (rows typed into the sheet have no Timestamp in column A)
  return v.slice(1).filter(r => r.some(x => x !== '' && x !== null)).map(r => {
    const o = {};
    h.forEach((k, i) => { o[k] = r[i]; });
    return o;
  });
}

function writeTable_(name, headers, rows, fmts) {
  const sh = getSheet_(name);
  sh.clear();
  sh.getCharts().forEach(c => sh.removeChart(c));
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff').setWrap(true);
  if (rows.length) {
    sh.getRange(2, 1, rows.length, headers.length).setValues(rows);
    Object.keys(fmts || {}).forEach(c => sh.getRange(2, +c, rows.length, 1).setNumberFormat(fmts[c]));
  }
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, headers.length);
  return sh;
}

function partyMaster_() {
  const m = {};
  readTable_(S.PARTIES).forEach(r => { const k = key_(r['Party Name']); if (k) m[k] = r; });
  return m;
}

function latestBy_(rows, keyFn) {
  const m = {};
  rows.forEach(r => { const k = keyFn(r); if (k) m[k] = r; });
  return Object.keys(m).map(k => m[k]);
}

function uniq_(arr) {
  const seen = {}, out = [];
  arr.forEach(v => {
    const s = str_(v), k = s.toLowerCase();
    if (s && !seen[k]) { seen[k] = 1; out.push(s); }
  });
  return out.sort((a, b) => a.localeCompare(b));
}

function str_(v) { return String(v === null || v === undefined ? '' : v).trim(); }
function key_(v) { return str_(v).toLowerCase().replace(/\s+/g, ' '); }
function normVehicle_(v) { return str_(v).toUpperCase().replace(/[\s-]/g, ''); }
function sum_(arr, f) { return arr.reduce((s, x) => s + (x[f] || 0), 0); }

function num_(v) {
  if (typeof v === 'number') return v;
  const n = parseFloat(str_(v).replace(/[,₹\s]/g, ''));
  return isNaN(n) ? 0 : n;
}

function toDate_(v) {
  if (v instanceof Date) return isNaN(v.getTime()) ? null : v;
  if (!v) return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

function fmtD_(d) { return Utilities.formatDate(d, CONFIG.TIMEZONE, 'dd-MMM-yyyy'); }

// Like inr_ but keeps up to 2 decimals (12.5 qtl stays 12.5, not 13).
function fmtNum_(n) {
  n = Number(n) || 0;
  const neg = n < 0, a = Math.abs(n);
  let whole = Math.floor(a), frac = Math.round((a - whole) * 100);
  if (frac === 100) { whole++; frac = 0; }
  const dec = frac ? '.' + String(frac).padStart(2, '0').replace(/0$/, '') : '';
  return (neg && (whole || frac) ? '-' : '') + inr_(whole) + dec;
}

function inr_(n) {  // Indian digit grouping: 12,34,567
  n = Math.round(n);
  const s = String(Math.abs(n));
  let last3 = s.slice(-3);
  const rest = s.slice(0, -3);
  if (rest) last3 = ',' + last3;
  return (n < 0 ? '-' : '') + rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + last3;
}

function colLetter_(n) {
  let s = '';
  while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
  return s;
}

function esc_(s) { return str_(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }

function htmlTable_(head, rows) {
  const td = 'style="border:1px solid #ccc;padding:4px 8px"';
  return '<table style="border-collapse:collapse">' +
    '<tr>' + head.map(h => '<th ' + td + '>' + esc_(h) + '</th>').join('') + '</tr>' +
    rows.map(r => '<tr>' + r.map(c => '<td ' + td + '>' + esc_(c) + '</td>').join('') + '</tr>').join('') +
    '</table>';
}

// For functions run from Apps Script (setup, upgrades): writes the message to the Execution log and shows a
// small notice in the sheet that closes by itself. It never waits for an OK click, so it can't cause a time-out.
function note_(title, msg) {
  Logger.log(title + ': ' + msg);
  try { ss_().toast(msg, title, 15); } catch (e) { /* sheet not open — the log has the message */ }
}

// For menu items in the sheet: a pop-up with an OK button.
function alert_(title, msg) {
  Logger.log(title + ': ' + msg);
  try { SpreadsheetApp.getUi().alert(title, msg, SpreadsheetApp.getUi().ButtonSet.OK); }
  catch (e) { /* run from the script editor — message is in the Execution log */ }
}
