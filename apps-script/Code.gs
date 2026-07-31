/**
 * ShipInsure MSA + SOW draft generator — Apps Script Web App backend.
 *
 * Copies the MSA + SOW template Google Doc, fills in the known merge fields,
 * inserts any selected optional SOW clauses (Upfront Incentive, Exclusivity,
 * Third-Party Cost Coverage, 90-Day Guarantee/Custom Claims, Software Buyout)
 * at the correct point in the document, strips the "OPTIONAL INSERTS" appendix,
 * renames the copy to "<Merchant Name> Draft MSA + SOW", and returns its URL.
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
  TEMPLATE_DOC_ID: 'PUT_YOUR_TEMPLATE_GOOGLE_DOC_ID_HERE',
  SHARED_SECRET: 'PUT_A_RANDOM_SECRET_STRING_HERE',
  DEST_FOLDER_ID: '', // optional
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    if (CONFIG.SHARED_SECRET && payload.secret !== CONFIG.SHARED_SECRET) {
      return jsonResponse({ ok: false, error: 'Unauthorized' });
    }
    const docUrl = generateMsaDraft(payload);
    return jsonResponse({ ok: true, docUrl });
  } catch (err) {
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
  applyCoreMergeFields(body, fields);

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
  const total = body.getNumChildren();
  for (let i = total - 1; i >= idx; i--) {
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

function upfrontIncentiveParagraphs(fields, clause) {
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
  if (clause.recoupmentApplies !== false) {
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
    text: 'Repayment. Upon a Cessation of Use (as defined in the MSA), the Client shall repay the Company the Unrecouped Balance, due within ' +
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

function thirdPartyCostParagraphs(fields, clause) {
  const vendorA = val(clause.vendorA, '[vendor / platform name]');
  const amountA = val(clause.amountA, '$_____');
  const categoryB = val(clause.categoryB, '[cost category]');
  const capBText = (clause.amountB && String(clause.amountB).trim())
    ? ('up to $' + clause.amountB + ' per Contract Year')
    : 'without an annual cap';
  const billingCadence = val(fields.billingCadence, 'monthly');
  const reportingDays = val(clause.reportingDays, '[30]');
  const repaymentDays = val(clause.repaymentDays, '[30]');

  return [
    {
      boldPrefix: 'Covered Third-Party Costs.',
      text: 'Covered Third-Party Costs. As further consideration for the Client’s commitment under this SOW, during the Minimum Term the ' +
        'Company will cover the Client’s third-party costs in the following categories: (a) ' + vendorA + ' fees, up to $' + amountA +
        ' per Contract Year; and (b) ' + categoryB + ', ' + capBText + ' (together, the “Covered Third-Party Costs”). “Contract Year” ' +
        'means each successive twelve (12) month period beginning on Full Deployment.',
    },
    {
      boldPrefix: 'Annual Caps.',
      text: 'Annual Caps. Any amounts stated above as a cap are annual maximums and not fixed or guaranteed payments. The Company is ' +
        'responsible only for Covered Third-Party Costs actually incurred by the Client, up to the applicable cap, and the Client remains ' +
        'solely responsible for any amount exceeding a cap. Unused amounts under a cap do not carry over to a subsequent Contract Year, are ' +
        'not interchangeable between categories, and are not payable to the Client in cash.',
    },
    {
      boldPrefix: 'Reporting and Payment.',
      text: 'Reporting and Payment. The Client will submit to the Company, on a ' + billingCadence + ' basis, invoices, statements, or other ' +
        'reasonable documentation evidencing the Covered Third-Party Costs incurred, and the Company will reimburse the Client within ' +
        reportingDays + ' days after receipt of conforming documentation.',
    },
    {
      boldPrefix: 'No Assumption of Client Agreements.',
      text: 'No Assumption of Client Agreements. The Company’s payment of Covered Third-Party Costs does not make the Company a party to, ' +
        'or otherwise responsible for, the Client’s agreement with ' + vendorA + ', any carrier, or any other third party. The Client remains ' +
        'solely responsible for its obligations under those agreements, including renewal, termination, and any change in pricing. The Client ' +
        'will promptly notify the Company of any renewal, price increase, or other change in the applicable third-party fees or costs. Any ' +
        'amount the Company pays under this section is fixed at the amount stated above and will not increase if the Client’s third-party ' +
        'fees increase.',
    },
    {
      boldPrefix: 'Repayment.',
      text: 'Repayment. Upon a Cessation of Use (as defined in the MSA), the Client shall repay the Company all Covered Third-Party Costs ' +
        'paid or reimbursed by the Company during the twelve (12) months preceding the Cessation of Use, within ' + repaymentDays +
        ' days of the Cessation of Use. This repayment obligation is in addition to, and not in lieu of, any repayment obligation applicable ' +
        'to an Upfront Merchant Incentive or Buyout Payment.',
    },
  ];
}

function customClaimsParagraph(clause) {
  const reasons = val(clause.reasons, '[e.g., wrong item(s) shipped; return to sender; never reached first carrier scan]');
  return {
    boldPrefix: 'Custom Claims Policy.',
    text: 'Custom Claims Policy. In addition to the standard Covered Losses, the Company will cover the following claim reasons for the ' +
      'Client’s customers: ' + reasons + '.',
  };
}

function qualityGuaranteeParagraph(clause) {
  const coverage = val(clause.coverage, '[e.g., accidental damage (photo required); taste/quality guarantee]');
  const days = val(clause.days, '[90]');
  const resolution = val(clause.resolution, 'reship / store credit / refund');
  return {
    boldPrefix: 'Quality Guarantee.',
    text: 'Quality Guarantee. The Company will additionally honor the following merchant quality guarantees: ' + coverage + '. ' +
      'Notwithstanding the standard submission window in the Claim Acceptance Rate section, claims under this Quality Guarantee must be ' +
      'submitted within ' + days + ' days of the tracking delivery date. Resolution for Quality Guarantee claims will be ' + resolution +
      ' at the Company’s discretion in accordance with the Claims Policies.',
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
      text: 'Early Termination. If a Cessation of Use (as defined in the MSA) occurs within the Minimum Term, or if the Client determines ' +
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
  const thirdParty = clauses.thirdPartyCosts;
  const buyout = clauses.softwareBuyout;
  const customClaims = clauses.customClaims;
  const qualityGuarantee = clauses.qualityGuarantee;

  // Compensation-clause group — anchored right after the Revenue Share
  // section's Negative Net Revenue paragraph.
  if (incentive || exclusivity || thirdParty || buyout) {
    let cursor = findParagraphIndex(body, 'Negative Net Revenue; Carryforward.');
    if (cursor === -1) cursor = findParagraphIndex(body, 'Revenue Share is calculated and paid');
    if (cursor !== -1) {
      if (incentive) cursor = insertParagraphsAfter(body, cursor, upfrontIncentiveParagraphs(fields, incentive));
      if (buyout) cursor = insertParagraphsAfter(body, cursor, softwareBuyoutParagraphs(fields, buyout, !incentive));
      if (thirdParty) cursor = insertParagraphsAfter(body, cursor, thirdPartyCostParagraphs(fields, thirdParty));
      if (exclusivity) cursor = insertParagraphsAfter(body, cursor, exclusivityParagraphs(exclusivity));
    }
  }

  // 90-Day Guarantee + Custom Claims — anchored right after the Claim
  // Policies section, before Claim Acceptance Rate.
  if (customClaims || qualityGuarantee) {
    let cursor = findParagraphIndex(body, 'shipinsure-claims-policies');
    if (cursor !== -1) {
      const paras = [];
      if (customClaims) paras.push(customClaimsParagraph(customClaims));
      if (qualityGuarantee) paras.push(qualityGuaranteeParagraph(qualityGuarantee));
      insertParagraphsAfter(body, cursor, paras);
    }
  }
}

// ── Core template merge fields (always present in the base MSA/SOW) ────────

function applyCoreMergeFields(body, fields) {
  const replacements = [
    // Client identity — same value at every occurrence.
    { pattern: '\\[Client Legal Name\\]', value: val(fields.clientLegalName, '[Client Legal Name]') },
    { pattern: '\\[Client Address Line 1\\]', value: val(fields.clientAddress1, '[Client Address Line 1]') },
    { pattern: '\\[Client City, State, ZIP\\]', value: val(fields.clientCityStateZip, '[Client City, State, ZIP]') },
    { pattern: '\\[Country\\]', value: val(fields.clientCountry, '[Country]') },

    // Notice/term periods — each disambiguated by surrounding, doc-unique text.
    {
      pattern: 'at least \\[30\\] days prior to the end of the then-current term',
      value: 'at least ' + val(fields.renewalNoticeDays, '[30]') + ' days prior to the end of the then-current term',
    },
    {
      pattern: 'convenience upon \\[30\\] days’ prior written notice, subject to any minimum-commitment',
      value: 'convenience upon ' + val(fields.terminationDays, '[30]') + ' days’ prior written notice, subject to any minimum-commitment',
    },
    {
      pattern: 'fails to cure within \\[30\\] days after written notice',
      value: 'fails to cure within ' + val(fields.cureDays, '[30]') + ' days after written notice',
    },
    {
      pattern: 'cannot reach agreement within \\[30\\] days of the Company’s notice',
      value: 'cannot reach agreement within ' + val(fields.pricingReviewDays, '[30]') + ' days of the Company’s notice',
    },
    {
      pattern: 'may terminate this SOW upon \\[30\\] days’ written notice to the Client',
      value: 'may terminate this SOW upon ' + val(fields.pricingReviewDays, '[30]') + ' days’ written notice to the Client',
    },

    // Pricing
    {
      pattern: 'charged at the rate of \\[\\$1\\.95\\], or \\[3%\\] of the cart subtotal, whichever is higher',
      value: 'charged at the rate of $' + val(fields.premiumMin, '[$1.95]') + ', or ' + val(fields.premiumPct, '[3%]') +
        '% of the cart subtotal, whichever is higher',
    },
    {
      pattern: 'Premiums start at \\[\\$1\\.95\\] or \\[3%\\] over \\[\\$100\\] of the cart subtotal, whichever is greater',
      value: 'Premiums start at $' + val(fields.premiumMin, '[$1.95]') + ' or ' + val(fields.premiumPct, '[3%]') + '% over $' +
        val(fields.cartThreshold, '[$100]') + ' of the cart subtotal, whichever is greater',
    },

    // Revenue share / billing cadence
    { pattern: '\\[__\\]% of “Net Revenue”', value: val(fields.revSharePct, '[__]') + '% of “Net Revenue”' },
    {
      pattern: 'automatically billed \\[monthly\\], solely for the premiums',
      value: 'automatically billed ' + val(fields.billingCadence, '[monthly]') + ', solely for the premiums',
    },
    {
      pattern: 'calculated and paid on a \\[monthly\\] basis',
      value: 'calculated and paid on a ' + val(fields.billingCadence, '[monthly]') + ' basis',
    },

    // Claims
    { pattern: 'submitted within \\[14\\] days of the tracking delivery date', value: 'submitted within ' + val(fields.claimSubmissionDays, '[14]') + ' days of the tracking delivery date' },
    { pattern: 'resolve claims within \\[24 hours\\]', value: 'resolve claims within ' + val(fields.claimResponseTime, '[24 hours]') },

    // Coverage cap
    { pattern: 'capped at \\[\\$_____\\]\\. Orders exceeding', value: 'capped at $' + val(fields.orderCap, '[$_____]') + '. Orders exceeding' },
  ];

  replacements.forEach(function (r) {
    body.replaceText(r.pattern, escapeReplacement(r.value));
  });

  // Signature block — table cells, disambiguated by the "Printed Name:"/"Title:"
  // label within each cell rather than by (row, col), since coordinates would
  // silently break if the template's table layout ever changes.
  const tables = body.getTables();
  const sigTable = tables[tables.length - 1];
  if (sigTable && sigTable.getNumRows() > 0) {
    const row = sigTable.getRow(0);
    fillSignatureCell(row.getCell(0), fields.siSignerName, fields.siSignerTitle);
    fillSignatureCell(row.getCell(1), fields.clientSignerName, fields.clientSignerTitle);
  }
}

function fillSignatureCell(cell, name, title) {
  if (!cell) return;
  cell.replaceText('Printed Name: \\[_+\\]', 'Printed Name: ' + escapeReplacement(val(name, '[__________________]')));
  cell.replaceText('Title: \\[_+\\]', 'Title: ' + escapeReplacement(val(title, '[__________________]')));
}

// Apps Script's replaceText treats "$" specially in the replacement string
// (like a regex substitution target), so literal dollar signs / backslashes
// in merge values must be escaped before being used as a replacement.
function escapeReplacement(s) {
  return String(s).replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}
