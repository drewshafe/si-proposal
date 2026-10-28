/**
 * ShipInsure MSA + SOW draft generator — Apps Script Web App backend.
 *
 * Copies the 2027 MSA + SOW template Google Doc, fills in the merge fields (the
 * template carries Earth Breeze sample values in the highlighted spots — those
 * exact strings are what gets replaced), toggles the optional sections (Covered
 * Third-Party Costs, Exhibit B claim categories), inserts any selected optional SOW
 * clauses (Gross Revenue Share, Upfront Incentive, Exclusivity, Software Buyout),
 * clears the highlighting, renames the copy to "<Merchant Name> Draft MSA + SOW",
 * and returns its URL.
 *
 * SETUP (one time):
 * 1. Upload "ShipInsure MSA + SOW Template.docx" to Google Drive, then open it
 *    with Google Docs (right-click > Open with > Google Docs) so it becomes a
 *    native Google Doc, not a .docx file. Copy its file ID from the URL:
 *    https://docs.google.com/document/d/THIS_PART_IS_THE_ID/edit
 * 2. Paste that ID into CONFIG.TEMPLATE_DOC_ID below.
 * 3. Pick any random string and put it in CONFIG.SHARED_SECRET below — this is
 *    what keeps the public Web App URL from being usable by strangers who
 *    might discover it. Put that same string into si-proposal's index.html
 *    wherever MSA_WEBAPP_SECRET is referenced (see the setup note there).
 * 4. (Optional) If you want generated docs to land in a specific Drive folder
 *    instead of your Drive root, put that folder's ID in CONFIG.DEST_FOLDER_ID.
 * 5. Deploy > New deployment > type "Web app". Execute as "Me". Who has
 *    access: "Anyone". Deploy, authorize the requested permissions, and copy
 *    the Web app URL (ends in /exec).
 * 6. Paste that URL into si-proposal/index.html's MSA_WEBAPP_URL constant.
 */

const CONFIG = {
  // 2027 Standard MSA + SOW (frozen copy of "Drew Copy of 2027 Standard MSA + SOW
  // Template"). Its sample values (Earth Breeze, $.98 / 2.5%, 70%, $1,000, AfterShip
  // $2,000 monthly…) are the merge anchors below — don't edit that copy; make a new
  // copy and update the anchors if the template wording changes.
  // Previous (pre-2027) template: 11YmrLdYeTZEkpRkikzvLsZ7CipH_C-ggNGAM8iNVA-Q
  TEMPLATE_DOC_ID: '1UkJ4qXZ7W5v-xKIyxnmsZFN-eJRSy4rBU3xiUolIC1I',
  // Mutual NDA template — upload "Mutual Non-Disclosure Agreement" to Drive, open
  // it as a Google Doc, and paste its file ID here. Give its merge fields the same
  // bracket placeholders the MSA uses ([Client Legal Name], [Client Address Line 1],
  // [Client City, State, ZIP], [Effective Date]) + a trailing signature table.
  MNDA_TEMPLATE_DOC_ID: '18GcgWwXL2UKp-iFt3WkWBzes0xiCv-ATj357VAKYK80',
  SHARED_SECRET: 'SCfP4ZmkdJbf-IY2wtA_3m6FpWq26ZkQ',
  DEST_FOLDER_ID: '', // optional
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (CONFIG.SHARED_SECRET && payload.secret !== CONFIG.SHARED_SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    const docUrl = payload.kind === 'mnda' ? generateMndaDraft(payload) : generateMsaDraft(payload);
    return jsonResponse({ ok: true, docUrl });
  } catch (err) {
    console.error((err && err.stack) || String(err));
    return jsonResponse({ ok: false, error: String(err && err.message || err) });
  }
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function generateMsaDraft(payload) {
  const fields = payload.fields || {};
  const clauses = payload.clauses || {};
  const merchantName = payload.merchantName || fields.clientLegalName || 'Merchant';

  const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
  const destFolder = CONFIG.DEST_FOLDER_ID ? DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID) : null;
  const copyFile = destFolder
    ? templateFile.makeCopy(merchantName + ' Draft MSA + SOW', destFolder)
    : templateFile.makeCopy(merchantName + ' Draft MSA + SOW');

  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody();

  stripOptionalInsertsAppendix(body);
  insertSelectedClauses(body, fields, clauses);
  applyCoreMergeFields(body, fields, clauses);
  applyExhibitB(body, clauses);
  clearHighlights(body);

  doc.saveAndClose();
  return copyFile.getUrl();
}

// Mutual NDA draft — straight template fill (client identity + effective date +
// signature block), mirroring the MSA's client merge fields. Requires
// CONFIG.MNDA_TEMPLATE_DOC_ID to be set (see setup note on CONFIG).
function generateMndaDraft(payload) {
  if (!CONFIG.MNDA_TEMPLATE_DOC_ID) {
    throw new Error('MNDA template not configured — set CONFIG.MNDA_TEMPLATE_DOC_ID.');
  }
  const fields = payload.fields || {};
  const merchantName = payload.merchantName || fields.clientLegalName || 'Merchant';

  const templateFile = DriveApp.getFileById(CONFIG.MNDA_TEMPLATE_DOC_ID);
  const destFolder = CONFIG.DEST_FOLDER_ID ? DriveApp.getFolderById(CONFIG.DEST_FOLDER_ID) : null;
  const copyFile = destFolder
    ? templateFile.makeCopy(merchantName + ' Draft MNDA', destFolder)
    : templateFile.makeCopy(merchantName + ' Draft MNDA');

  const doc = DocumentApp.openById(copyFile.getId());
  const body = doc.getBody();

  // Client identity + effective date — same placeholder convention as the MSA.
  body.replaceText('\\[Client Legal Name\\]', escapeReplacement(val(fields.clientLegalName, '[Client Legal Name]')));
  body.replaceText('\\[Client Address Line 1\\]', escapeReplacement(val(fields.clientAddress1, '[Client Address Line 1]')));
  body.replaceText('\\[Client City, State, ZIP\\]', escapeReplacement(val(fields.clientCityStateZip, '[Client City, State, ZIP]')));
  body.replaceText('\\[Country\\]', escapeReplacement(val(fields.clientCountry, '[Country]')));
  body.replaceText('\\[Effective Date\\]', escapeReplacement(val(fields.effectiveDate, '[Effective Date]')));

  // Signature table (if present) — last table, ShipInsure then Client, same as MSA.
  const tables = body.getTables();
  if (tables.length) {
    const sigTable = tables[tables.length - 1];
    if (sigTable && sigTable.getNumRows() > 0) {
      const row = sigTable.getRow(0);
      fillSignatureCell(row.getCell(0), fields.siSignerName, fields.siSignerTitle);
      if (row.getNumCells() > 1) fillSignatureCell(row.getCell(1), fields.clientSignerName, fields.clientSignerTitle);
    }
  }

  doc.saveAndClose();
  return copyFile.getUrl();
}

// ── Value helpers ──────────────────────────────────────────────────────────

// Returns the provided value, or the original bracketed placeholder text so
// legal can immediately see what's still unfilled, rather than silently
// blanking it out.
function val(v, placeholderText) {
  const s = (v == null ? '' : String(v)).trim();
  return s ? s : placeholderText;
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// ── Appendix removal ────────────────────────────────────────────────────────

function stripOptionalInsertsAppendix(body) {
  const idx = findParagraphIndex(body, 'OPTIONAL INSERTS');
  if (idx === -1) return;
  // Append a fresh blank paragraph first so the document always has a safe
  // trailing paragraph to fall back on — Apps Script refuses to remove
  // whatever would become the document's last paragraph (even via setText),
  // so we never let the removal range include the actual last paragraph.
  body.appendParagraph('');
  const total = body.getNumChildren(); // includes the paragraph just appended
  for (let i = total - 2; i >= idx; i--) {
    body.removeChild(body.getChild(i));
  }
}

function findParagraphIndex(body, needle) {
  const n = body.getNumChildren();
  for (let i = 0; i < n; i++) {
    const child = body.getChild(i);
    if (child.getType && child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      const text = child.asParagraph().getText();
      if (text.indexOf(needle) !== -1) return i;
    }
  }
  return -1;
}

// ── Inserting paragraphs at an anchor, in order ─────────────────────────────

// paragraphs: array of strings, or {text, boldPrefix} to bold a leading label
function insertParagraphsAfter(body, afterIndex, paragraphs) {
  let idx = afterIndex;
  paragraphs.forEach(function (p) {
    idx += 1;
    const text = typeof p === 'string' ? p : p.text;
    const para = body.insertParagraph(idx, text);
    para.setSpacingAfter(8);
    if (typeof p !== 'string' && p.boldPrefix) {
      const end = p.boldPrefix.length - 1;
      if (end >= 0) para.editAsText().setBold(0, end, true);
    }
  });
  return idx;
}

// ── Optional clause content builders ────────────────────────────────────────

function configDeploymentParagraphs(fields, referencesIncentive) {
  const configFormat = val(fields.configFormat, 'dual checkout');
  const basis = referencesIncentive
    ? 'As a condition of the Upfront Merchant Incentive'
    : 'As a condition of the compensation described in this SOW';
  return [
    {
      boldPrefix: 'Configuration Requirement.',
      text: 'Configuration Requirement. ' + basis + ', the ShipInsure offer must be presented in a ' + configFormat +
        ' format at checkout to all of the Client’s customers on one hundred percent (100%) of the Client’s eligible orders ' +
        'across all of the Client’s storefronts and sales channels, with no split, holdback, staged rollout, or other configuration ' +
        'under which any portion of the Client’s orders is not presented the offer (the “Required Configuration”).',
    },
    {
      boldPrefix: 'Full Deployment; Testing.',
      text: 'Full Deployment; Testing. “Full Deployment” means the first date on which the ShipInsure offer is presented in the ' +
        'Required Configuration to all of the Client’s customers on one hundred percent (100%) of the Client’s eligible orders across all ' +
        'of the Client’s storefronts and sales channels. Prior to Full Deployment, the Client may configure, test, and optimize its checkout ' +
        'as it determines, and nothing in this SOW restricts the Client’s testing during that period. From Full Deployment through the end of ' +
        'the Minimum Term, the Client will maintain the Required Configuration. Promptly following Full Deployment, the parties will confirm the ' +
        'Full Deployment date in writing (email being sufficient), and the date so confirmed will be the Full Deployment date for all purposes ' +
        'under this SOW, including the start of the Minimum Term and, if applicable, of each Contract Year.',
    },
  ];
}

function upfrontIncentiveParagraphs(fields, clause, gross) {
  const amount = val(clause.amount, '$_____');
  const minTerm = val(fields.minTermMonths, '[12]');
  const outsideDays = val(clause.outsideDateDays, '[90]');
  const estVol = val(clause.estOrderVolume, '[_____]');
  const estPeriod = val(clause.estOrderVolumePeriod, 'week / month');
  const validationDays = val(clause.validationDays, '[7]');
  const thresholdPct = val(clause.validationThresholdPct, '[80]');
  const repaymentDays = val(clause.repaymentDays, '[30]');

  const paragraphs = [
    {
      boldPrefix: 'Upfront Merchant Incentive.',
      text: 'Upfront Merchant Incentive. The Company will pay the Client a one-time payment of $' + amount +
        ' (the “Upfront Merchant Incentive”), subject to the conditions below. The Client agrees to use the Services for at least ' +
        minTerm + ' months beginning on Full Deployment, as defined below (the “Minimum Term”).',
    },
  ];
  paragraphs.push.apply(paragraphs, configDeploymentParagraphs(fields, true));
  paragraphs.push(
    {
      boldPrefix: 'Outside Date.',
      text: 'Outside Date. If Full Deployment has not occurred by ' + outsideDays + ' days after the Effective Date, either party may ' +
        'terminate this SOW upon written notice to the other, in which case no Upfront Merchant Incentive and no Covered Third-Party Costs ' +
        'will be owed or payable, and neither party will have any further obligation to the other except for obligations intended to survive ' +
        'termination. The parties may extend this date by written agreement (email being sufficient).',
    },
    {
      boldPrefix: 'Estimated Order Volume.',
      text: 'Estimated Order Volume. The Client has represented to the Company that its estimated order volume is ' + estVol +
        ' orders per ' + estPeriod + ' (the “Estimated Order Volume”).',
    },
    {
      boldPrefix: 'Payment Trigger and Volume Validation.',
      text: 'Payment Trigger and Volume Validation. The Company will pay the Upfront Merchant Incentive within ' + validationDays +
        ' days after the Client has been live with the Services in the Required Configuration for one (1) full week beginning on Full ' +
        'Deployment (the “Validation Period”), provided that during the Validation Period the Client’s actual order volume is not less ' +
        'than ' + thresholdPct + '% of the Estimated Order Volume, pro-rated to the length of the Validation Period. No period prior to Full ' +
        'Deployment counts toward the Validation Period. If the Client’s validated volume falls below this threshold, the Company may ' +
        'withhold or proportionately adjust the Upfront Merchant Incentive, or extend the Validation Period until the threshold is met.',
    }
  );
  if (clause.recoupmentApplies !== false && gross) {
    // Gross deals: the incentive is an advance against the Client's Revenue Share.
    paragraphs.push({
      boldPrefix: 'Recoupment of Upfront Merchant Incentive.',
      text: 'Recoupment of Upfront Merchant Incentive. The Upfront Merchant Incentive is an advance against the Client’s Revenue Share. ' +
        'Beginning with the billing period in which the Upfront Merchant Incentive is paid, the Company will withhold the Revenue Share ' +
        'otherwise payable to the Client and apply it against the Upfront Merchant Incentive until the Upfront Merchant Incentive has been ' +
        'recouped in full. The “Unrecouped Balance” means, at any time, the Upfront Merchant Incentive less the aggregate Revenue Share ' +
        'withheld and applied against it. In the billing period in which the Unrecouped Balance reaches zero, the Company will pay the Client ' +
        'the portion of that period’s Revenue Share remaining after the Unrecouped Balance has been fully reduced, and Revenue Share is ' +
        'payable in full for each billing period thereafter.',
    });
  } else if (clause.recoupmentApplies !== false) {
    paragraphs.push({
      boldPrefix: 'Recoupment of Upfront Merchant Incentive.',
      text: 'Recoupment of Upfront Merchant Incentive. Beginning with the billing period in which the Upfront Merchant Incentive is paid, ' +
        'the Company will retain one hundred percent (100%) of Net Revenue, and no Revenue Share is payable to the Client, until the Company ' +
        'has recouped the Upfront Merchant Incentive in full. The “Unrecouped Balance” means, at any time, the Upfront Merchant Incentive ' +
        'less the aggregate Net Revenue applied against it. In each billing period, Net Revenue is applied first to eliminate any carried-forward ' +
        'negative Net Revenue balance, then to reduce the Unrecouped Balance, and only then is Revenue Share payable. In the billing period in ' +
        'which the Unrecouped Balance reaches zero, Revenue Share is payable on the Net Revenue remaining in that period after the Unrecouped ' +
        'Balance has been fully reduced.',
    });
  }
  paragraphs.push({
    boldPrefix: 'Repayment.',
    text: 'Repayment. Upon a Cessation of Use (as defined in this SOW), the Client shall repay the Company the Unrecouped Balance, due within ' +
      repaymentDays + ' days of the Cessation of Use.',
  });
  return paragraphs;
}

function exclusivityParagraphs(clause) {
  const repaymentDays = val(clause.repaymentDays, '[30]');
  return [{
    boldPrefix: 'Exclusivity.',
    text: 'Exclusivity. The Services will be the exclusive package protection service offered on the Client’s storefronts and sales ' +
      'channels during the Minimum Term. If, during the Minimum Term, the Client uses another package protection service, or replaces the ' +
      'Required Configuration with another service, the Client will reimburse the Company for the Revenue Share paid to the Client over the ' +
      'prior twelve (12) months, due within ' + repaymentDays + ' days of written notice.',
  }];
}

// MSA §1.8 "Cessation of Use" was dropped from the 2027 MSA, so any SOW with an
// upfront payment carries the definition itself.
function cessationOfUseParagraph() {
  return {
    boldPrefix: 'Cessation of Use.',
    text: 'Cessation of Use. “Cessation of Use” means the occurrence of any of the following prior to the end of the Minimum Term: ' +
      '(a) the Client terminates this SOW for any reason; (b) the Client ceases using ShipInsure’s dual checkout on its site; or (c) the ' +
      'Client changes away from the Required Configuration or any other checkout or program setup in place upon Full Deployment, ' +
      'including switching off a pre-checked, opt-out, or dual checkout format, in each case without the Company’s prior written consent.',
  };
}

function softwareBuyoutParagraphs(fields, clause, includeConfigBlock) {
  const amount = val(clause.amount, '$_____');
  const paymentDays = val(clause.paymentDays, '[30]');
  const competitorName = val(clause.competitorName, '[name of software/competitor, e.g., Narvar]');
  const partnerClause = (clause.partnerName && String(clause.partnerName).trim())
    ? (' and the Company’s partner, ' + clause.partnerName + ', if applicable')
    : '';
  const minTerm = val(fields.minTermMonths, '[12]');
  const earlyTermDays = val(clause.earlyTermDays, '[15]');

  const paragraphs = [];
  if (includeConfigBlock) paragraphs.push.apply(paragraphs, configDeploymentParagraphs(fields, false));
  paragraphs.push(
    {
      boldPrefix: 'Competitive Buyout.',
      text: 'Competitive Buyout. The Company will pay the Client a one-time payment of $' + amount + ' (the “Buyout Payment”) within ' +
        paymentDays + ' days after the Effective Date, to be applied by the Client toward buying out its existing ' + competitorName +
        ' contract and transitioning to the Services' + partnerClause + '. In return, the Client agrees to use the Services for at least ' +
        minTerm + ' months beginning on Full Deployment (the “Minimum Term”).',
    },
    {
      boldPrefix: 'Early Termination.',
      text: 'Early Termination. If a Cessation of Use (as defined in this SOW) occurs within the Minimum Term, or if the Client determines ' +
        'within the Minimum Term that the Services are not satisfactory, the Client may (or, in the case of a Cessation of Use, shall) repay ' +
        'the full Buyout Payment within ' + earlyTermDays + ' days by written notice to the Company. Upon receipt of repayment, this SOW will ' +
        'terminate and neither party will have further obligations except those intended to survive termination.',
    }
  );
  return paragraphs;
}

// ── Clause insertion orchestration ──────────────────────────────────────────

function insertSelectedClauses(body, fields, clauses) {
  const incentive = clauses.upfrontIncentive;
  const exclusivity = clauses.exclusivity;
  const buyout = clauses.softwareBuyout;
  const minTerm = !!(incentive || buyout);

  // Compensation-clause group — anchored right after the Revenue Share
  // section's Negative Net Revenue paragraph.
  if (incentive || exclusivity || buyout) {
    let cursor = findParagraphIndex(body, 'Negative Net Revenue; Carryforward.');
    if (cursor === -1) cursor = findParagraphIndex(body, 'Revenue Share is calculated and paid');
    if (cursor !== -1) {
      if (incentive) cursor = insertParagraphsAfter(body, cursor, upfrontIncentiveParagraphs(fields, incentive, !!clauses.grossRevShare));
      if (buyout) cursor = insertParagraphsAfter(body, cursor, softwareBuyoutParagraphs(fields, buyout, !incentive));
      if (minTerm) cursor = insertParagraphsAfter(body, cursor, [cessationOfUseParagraph()]);
      if (exclusivity) cursor = insertParagraphsAfter(body, cursor, exclusivityParagraphs(exclusivity));
    }
  }

  // The 2027 SOW is month-to-month with "There is no minimum term." — an upfront
  // payment adds a Minimum Term, so the Term section and MSA §12.1 say so.
  if (minTerm) {
    body.replaceText('this SOW shall continue on a month-to-month basis', 'this SOW shall continue for the Minimum Term set forth below and thereafter on a month-to-month basis');
    body.replaceText('There is no minimum term\\.', 'The Minimum Term begins on Full Deployment, as defined and confirmed in this SOW.');
    body.replaceText('in accordance with Section 12\\.1 of the MSA\\.', 'in accordance with Section 12.1 of the MSA, subject to the minimum-commitment and repayment terms of this SOW.');
    body.replaceText('days’ prior written notice\\.', 'days’ prior written notice, subject to any minimum-commitment or repayment terms set forth in the applicable SOW.');
  }

  // Gross Revenue Share — rewrite the SOW's Revenue Share paragraph and replace the
  // Net-Revenue carryforward paragraph (done after the compensation inserts, which
  // anchor on that paragraph's heading text).
  const recoup = incentive && (incentive.recoupmentApplies !== false);
  if (clauses.grossRevShare) {
    applyGrossRevenueShare(body, fields, clauses.grossRevShare, incentive);
  } else if (recoup) {
    body.replaceText('\\(the “Revenue Share”\\)\\.', '(the “Revenue Share”), subject to the Recoupment provision below.');
  }
}

// ── Gross Revenue Share ─────────────────────────────────────────────────────
// Revenue Share is a % of Gross Premiums (premiums collected, less premium
// refunds and chargebacks) instead of Net Revenue. The Company bears all claims
// cost; there is no negative balance or carryforward. Overrides MSA §§1.7/1.8
// for this SOW (the MSA already lets the SOW control commercial terms).
function grossRevenueShareParagraphs(fields, clause, recoup) {
  const pct = val(fields.revSharePct, '[__]');
  const payDays = val(clause.paymentDays, '[30]');
  return {
    share: 'The Company will pay the Client ' + pct + '% of “Gross Premiums” (the “Revenue Share”)' +
      (recoup ? ', subject to the Recoupment provision below' : '') + '. “Gross Premiums” means the total premiums the Company ' +
      'actually collects from the Client’s customers during a billing period, less any premiums the Company refunds to customers and any ' +
      'chargebacks or payment reversals of premiums during that period. Gross Premiums are not reduced by claims paid, by Covered ' +
      'Third-Party Costs, or by any other cost of providing the Services. Notwithstanding Sections 1.7 and 1.8 of the MSA, the Revenue ' +
      'Share under this SOW is calculated on Gross Premiums rather than Net Revenue, for each billing period, and paid as set forth below.',
    after: [
      {
        boldPrefix: 'Claims Risk; No Carryforward.',
        text: 'Claims Risk; No Carryforward. The Company bears the full cost of all claims paid and refunds issued under this SOW. The ' +
          'Revenue Share for a billing period is not reduced by claims paid in that or any other billing period, no negative balance is ' +
          'carried forward between billing periods, and the Client has no liability for any amount by which claims exceed premiums. The ' +
          'parties acknowledge that the Revenue Share is set in reliance on the premium rate and coverage terms in this SOW, and that any ' +
          'proposed change to the Revenue Share percentage is governed by the Premium Evaluation and Adjustment section below.',
      },
      {
        boldPrefix: 'Revenue Share Statement and Payment.',
        text: 'Revenue Share Statement and Payment. For each billing period, the Company will provide the Client a statement showing the ' +
          'Gross Premiums collected and the Revenue Share earned for that period. The Company will pay the Revenue Share within ' + payDays +
          ' days after the end of the billing period, either by separate payment or as a credit against the Company’s premium invoice to ' +
          'the Client for that period, at the Company’s election. Revenue Share is earned only on premiums the Company actually collects; ' +
          'if a premium on which Revenue Share was paid is later refunded or reversed, the Company may deduct the corresponding Revenue ' +
          'Share from a subsequent payment.',
      },
    ],
  };
}

function applyGrossRevenueShare(body, fields, clause, incentive) {
  const recoup = incentive && (incentive.recoupmentApplies !== false);
  const g = grossRevenueShareParagraphs(fields, clause, recoup);
  const shareIdx = findParagraphIndex(body, 'of “Net Revenue”');
  if (shareIdx !== -1) body.getChild(shareIdx).asParagraph().setText(g.share);
  const carryIdx = findParagraphIndex(body, 'Negative Net Revenue; Carryforward.');
  if (carryIdx !== -1) {
    const para = body.getChild(carryIdx).asParagraph();
    const first = g.after[0];
    para.setText(first.text);
    para.editAsText().setBold(false);
    para.editAsText().setBold(0, first.boldPrefix.length - 1, true);
    insertParagraphsAfter(body, carryIdx, g.after.slice(1));
  } else if (shareIdx !== -1) {
    insertParagraphsAfter(body, shareIdx, g.after);
  }
}

// ── Core template merge fields (2027 template) ──────────────────────────────
// The 2027 template carries Earth Breeze sample values in the highlighted spots;
// each pattern below is that exact sample text (regex-escaped) plus enough
// surrounding words to be unique in the document.

function applyCoreMergeFields(body, fields, clauses) {
  clauses = clauses || {};
  const cur = val(fields.currencySymbol, '$'); // merchant's currency from the calc link
  const quarterly = /quarter/i.test(String(fields.billingCadence || ''));
  const name = val(fields.clientLegalName, '[Client Legal Name]');

  // Client header block (first table, Client column) — address lines + country.
  // Done per-cell because "United States" also appears in ShipInsure's column.
  const tables = body.getTables();
  if (tables.length && tables[0].getNumRows() > 1 && tables[0].getRow(1).getNumCells() > 1) {
    const cell = tables[0].getRow(1).getCell(1);
    cell.replaceText('500 Westover Dr #14160', val(fields.clientAddress1, '[Client Address Line 1]'));
    cell.replaceText('Sanford, NC 27330', val(fields.clientCityStateZip, '[Client City, State, ZIP]'));
    cell.replaceText('United States', val(fields.clientCountry, '[Country]'));
  }

  const replacements = [
    // Client legal name — header, MSA intro, SOW intro, Ongoing Relationship, signature block.
    { pattern: 'Earth Breeze, Inc\\.', value: name },

    // Notice/term periods — each disambiguated by surrounding, doc-unique text.
    { pattern: 'for convenience upon 30 days’', value: 'for convenience upon ' + val(fields.terminationDays, '30') + ' days’' },
    { pattern: 'fails to cure within 30 days after written notice', value: 'fails to cure within ' + val(fields.cureDays, '30') + ' days after written notice' },
    { pattern: 'cannot reach agreement within 30 days of the Company’s notice', value: 'cannot reach agreement within ' + val(fields.pricingReviewDays, '30') + ' days of the Company’s notice' },
    { pattern: 'may terminate this SOW upon 30 days’ written notice to the Client', value: 'may terminate this SOW upon ' + val(fields.pricingReviewDays, '30') + ' days’ written notice to the Client' },

    // Pricing
    {
      pattern: 'charged at the rate of \\$\\.98, or 2\\.5% of the cart subtotal',
      value: 'charged at the rate of ' + cur + val(fields.premiumMin, '[_____]') + ', or ' + val(fields.premiumPct, '[__]') + '% of the cart subtotal',
    },

    // Premium collecting — calendar month / quarter
    {
      pattern: 'generated on the first day of each calendar month for premiums accrued from protected orders in the prior month',
      value: quarterly
        ? 'generated on the first day of each calendar quarter for premiums accrued from protected orders in the prior quarter'
        : 'generated on the first day of each calendar month for premiums accrued from protected orders in the prior month',
    },

    // Revenue share
    { pattern: 'will pay the Client 70% of', value: 'will pay the Client ' + val(fields.revSharePct, '[__]') + '% of' },

    // Claims
    { pattern: 'submitted within 14 days of the tracking delivery date', value: 'submitted within ' + val(fields.claimSubmissionDays, '14') + ' days of the tracking delivery date' },
    { pattern: 'resolve claims within 24 hours', value: 'resolve claims within ' + val(fields.claimResponseTime, '24 hours') },
  ];

  replacements.forEach(function (r) {
    body.replaceText(r.pattern, escapeReplacement(r.value));
  });

  applyCoveredThirdPartyCosts(body, fields, clauses.thirdPartyCosts, !!clauses.grossRevShare);

  // Coverage cap — fill from the Limit $ (orderCap). If none is set, remove the
  // whole "Coverage Cap for High-Value Orders" section rather than leaving the
  // sample amount in a client-facing draft.
  var orderCap = String(val(fields.orderCap, '')).replace(/[^0-9.,]/g, '');
  if (orderCap !== '') {
    body.replaceText('initially capped at \\$1,000\\.', 'initially capped at ' + cur + orderCap + '.');
  } else {
    removeParagraphByText(body, 'Coverage Cap for High-Value Orders');
    removeParagraphByText(body, 'Coverage for any single order is');
  }

  // Signature block — table cells, disambiguated by the "Printed Name:"/"Title:"
  // label within each cell rather than by (row, col), since coordinates would
  // silently break if the template's table layout ever changes.
  const sigTable = tables[tables.length - 1];
  if (sigTable && sigTable.getNumRows() > 0) {
    const row = sigTable.getRow(0);
    fillSignatureCell(row.getCell(0), fields.siSignerName, fields.siSignerTitle);
    fillSignatureCell(row.getCell(1), fields.clientSignerName, fields.clientSignerTitle);
  }
}

// ── Covered Third-Party Costs (2027 SOW section) ────────────────────────────
// Selected: fill the vendor / product / amount into the template's AfterShip
// sample. Not selected: remove the section and the Net Revenue "(b) Covered
// Third-Party Costs" carve-outs so the draft never mentions them.
function applyCoveredThirdPartyCosts(body, fields, tpc, gross) {
  const defIdx = findParagraphIndex(body, 'means the SaaS fees the Company pays');
  if (tpc) {
    if (defIdx === -1) return;
    const para = body.getChild(defIdx).asParagraph();
    const cur = val(fields.currencySymbol, '$');
    const vendor = val(tpc.vendor, '[vendor]');
    const amount = String(tpc.amount || '').trim();
    const fee = amount ? (' (' + cur + amount + ' ' + val(tpc.period, 'monthly') + ')') : '';
    para.replaceText('AfterShip’s Return Platform \\(\\$2,000 monthly\\)', vendor + '’s ' + val(tpc.product, '[product / platform]') + fee);
    para.replaceText('AfterShip', vendor);
    if (gross) {
      para.replaceText('Covered Third-Party Costs are deducted in calculating Net Revenue for the period in which the Company pays them\\.',
        'Covered Third-Party Costs are borne by the Company and do not reduce Gross Premiums or the Revenue Share.');
    }
    return;
  }
  if (defIdx !== -1) {
    // Heading is the paragraph right before the definition.
    const prev = defIdx > 0 ? body.getChild(defIdx - 1) : null;
    body.getChild(defIdx).removeFromParent();
    if (prev && prev.getType() === DocumentApp.ElementType.PARAGRAPH &&
        prev.asParagraph().getText().trim() === 'Covered Third-Party Costs') prev.removeFromParent();
  }
  body.replaceText('minus \\(a\\) the total amount of claims the Company actually pays out \\(and any refunds the Company issues\\) for that period, and \\(b\\) any Covered Third-Party Costs the Company pays or reimburses under this SOW for that period\\.',
    'minus the total amount of claims the Company actually pays out (and any refunds the Company issues) for that period.');
  body.replaceText('claims paid out \\(together with any refunds issued\\) and Covered Third-Party Costs exceed premiums collected',
    'claims paid out (together with any refunds issued) exceed premiums collected');
}

// ── Exhibit B — additional claim categories ─────────────────────────────────
// The template lists Quality Satisfaction Guarantee, Manufacturer Defect and
// Missing Items; unselected ones are removed. Free-text custom claim reasons are
// appended as one more Exhibit B category.
function applyExhibitB(body, clauses) {
  const qg = clauses.qualityGuarantee;
  if (qg) {
    body.replaceText('manufacturing defects within X days of delivery', 'manufacturing defects within ' + val(qg.days, '[__]') + ' days of delivery');
  } else {
    removeParagraphExact(body, 'Quality Satisfaction Guarantee');
    removeParagraphByText(body, 'Provides coverage for accidental damage during normal use');
  }
  if (!clauses.manufacturerDefect) {
    removeParagraphExact(body, 'Manufacturer Defect');
    removeParagraphByText(body, 'Provides coverage for items that arrive defective');
  }
  if (!clauses.missingItems) {
    removeParagraphExact(body, 'Missing Items');
    removeParagraphByText(body, 'Provides coverage when items are missing');
  }

  const reasons = clauses.customClaims && String(clauses.customClaims.reasons || '').trim();
  if (reasons) {
    let anchor = -1;
    ['Provides coverage when items are missing', 'Provides coverage for items that arrive defective',
      'Provides coverage for accidental damage during normal use', 'claims policies in effect on the Effective Date']
      .some(function (n) { anchor = findParagraphIndex(body, n); return anchor !== -1; });
    if (anchor !== -1) {
      const h = body.insertParagraph(anchor + 1, 'Additional Covered Claim Reasons');
      h.setHeading(DocumentApp.ParagraphHeading.HEADING2);
      body.insertParagraph(anchor + 2, 'Provides coverage for the following additional claim reasons: ' + reasons.replace(/[.;\s]+$/, '') + '.')
        .setHeading(DocumentApp.ParagraphHeading.NORMAL);
    }
  }
}

// Template highlights mark the merge spots for the author — never ship them.
function clearHighlights(body) {
  let r = null;
  while ((r = body.findElement(DocumentApp.ElementType.TEXT, r))) {
    const t = r.getElement().asText();
    const len = t.getText().length;
    if (len) t.setBackgroundColor(0, len - 1, null);
  }
}

function fillSignatureCell(cell, name, title) {
  if (!cell) return;
  cell.replaceText('Printed Name: \\[_+\\]', 'Printed Name: ' + escapeReplacement(val(name, '[__________________]')));
  cell.replaceText('Title: \\[_+\\]', 'Title: ' + escapeReplacement(val(title, '[__________________]')));
}

// DocumentApp.Body.replaceText() uses the replacement as a LITERAL string — unlike
// regex substitution APIs, "$" is not special here. The previous version escaped
// "$" -> "\$", which Apps Script then inserted verbatim, producing the stray
// backslashes seen in drafts (e.g. "\$0.98"). Keep the replacement literal.
function escapeReplacement(s) {
  return String(s);
}

// Remove the first paragraph whose text contains `needle` (used to drop an entire
// optional section, e.g. the coverage cap, when it doesn't apply to this deal).
function removeParagraphExact(body, text) {
  const paras = body.getParagraphs();
  for (let i = 0; i < paras.length; i++) {
    if (paras[i].getText().trim() === text) {
      try { paras[i].removeFromParent(); } catch (e) {}
      return true;
    }
  }
  return false;
}

function removeParagraphByText(body, needle) {
  const paras = body.getParagraphs();
  for (let i = 0; i < paras.length; i++) {
    if (paras[i].getText().indexOf(needle) !== -1) {
      try { paras[i].removeFromParent(); } catch (e) {}
      return true;
    }
  }
  return false;
}
