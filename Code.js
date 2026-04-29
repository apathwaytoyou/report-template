// ================================================================
//  CLINICAL TOOLS  ·  Google Apps Script
//  Code.gs
// ================================================================

// ——— CONTROL PANEL SHEET ——————————————————————————————————————————
var CONTROL_PANEL_ID        = '1eQZ5b89RrPQgnQi6OkHMv6_U9w2gFICwFNu7nYhme38';
var CASES_TAB               = 'Cases';
var NOVO_MEASURES_TAB       = 'Measures';
var COL_EHR_ID              = 1;   // B
var COL_ACTIVE_FORMS        = 14;  // O
var COL_COMPLETED_FORMS     = 15;  // P
var COL_CLIENT_NOVO_LINK    = 17;  // R
var COL_CLIENT_NOVO_FORMS   = 18;  // S
var COL_COLLAT_NOVO_LINK    = 19;  // T
var COL_COLLAT_NOVO_FORMS   = 20;  // U

// ——— MENU ————————————————————————————————————————————————————————

function onOpen() {
  DocumentApp.getUi()
    .createMenu('⚕️ Clinical Tools')
    .addItem('⚙️ Setup / Link Case…', 'showSetupSidebar')
    .addSeparator()
    .addItem('🧹 Clean Line Breaks (Selected Text)', 'cleanLineBreaks')
    .addSeparator()
    .addItem('🔄 Replace Pronouns…', 'showPronounDialog')
    .addSeparator()
    .addItem('📋 Diagnostic Snippets…', 'showSnippetSidebar')
    .addItem('🔲 Coventry Grid…', 'showCGSidebar')    
    .addSeparator()
    .addItem('📊 Build Score Table (Score Table)', 'buildScoreTable')
    .addItem('📊 Score Entry…', 'showScoreSidebar')
    .addItem('📋 Measures…', 'showMeasuresSidebar')    
    .addItem('🧹 Tidy Score Table', 'tidyScoreTable')
    .addItem('📋 Generate Data Sources List', 'generateDataSources')
    .addSeparator()
    .addItem('📚 Generate Resource List…', 'openResoListSidebar')
    .addSeparator()
    .addItem('🗑️ Remove Placeholders', 'clearPlaceholders')
    .addSeparator()
    .addItem('✅ Finalize Document', 'finalizeDocument')
    .addToUi();

  // Auto-open setup sidebar if no case is linked yet
  var ehrId = getLinkedEhrId();
  if (!ehrId) showSetupSidebar();
}

// ——— SHARED UTILITIES ————————————————————————————————————————————

function escapeRegex(str) {
  return str.replace(/[-[\]{}()*+?.,\\^$|#\s]/g, '\\$&');
}

/**
 * Returns `replacement` with the same capitalisation as `original`.
 *   "He" → "She"  |  "HE" → "SHE"  |  "he" → "she"
 */
function matchCase(original, replacement) {
  if (!original || original.length === 0) return replacement;
  if (original === original.toUpperCase() && original !== original.toLowerCase()) {
    return replacement.toUpperCase();
  }
  if (original[0] === original[0].toUpperCase() &&
      original[0] !== original[0].toLowerCase()) {
    return replacement.charAt(0).toUpperCase() + replacement.slice(1);
  }
  return replacement.toLowerCase();
}


/**
 * Returns true if the character immediately before the match
 * is a ^ — checks across text run boundaries.
 */
function isPrecededByCaret(element, startOffset) {
  if (startOffset > 0) {
    return element.getText().charAt(startOffset - 1) === '^';
  }
  // startOffset is 0 — check the end of the previous sibling
  var parent = element.getParent();
  if (!parent) return false;
  var index  = parent.getChildIndex(element);
  if (index <= 0) return false;
  var prevSibling = parent.getChild(index - 1);
  if (!prevSibling || prevSibling.getType() !== DocumentApp.ElementType.TEXT) return false;
  var prevText = prevSibling.asText().getText();
  return prevText.length > 0 && prevText.charAt(prevText.length - 1) === '^';
}

// ——— PART 1 : CLEAN LINE BREAKS ——————————————————————————————————

function cleanLineBreaks() {
  var ui  = DocumentApp.getUi();
  var doc = DocumentApp.getActiveDocument();
  var sel = doc.getSelection();

  if (!sel) {
    ui.alert(
      'No Text Selected',
      'Highlight the pasted text first, then run "Clean Line Breaks."',
      ui.ButtonSet.OK
    );
    return;
  }

  var body  = doc.getBody();
  var elems = sel.getRangeElements();

  function parentPara(el) {
    while (el && el.getType() !== DocumentApp.ElementType.PARAGRAPH) {
      if (typeof el.getParent !== 'function') return null;
      el = el.getParent();
    }
    return el;
  }

  var firstPara = parentPara(elems[0].getElement());
  var lastPara  = parentPara(elems[elems.length - 1].getElement());

  if (!firstPara || !lastPara) {
    ui.alert('Selection Not Supported',
      'Please select only regular paragraph text (not a table or image).', ui.ButtonSet.OK);
    return;
  }

  var startIdx = body.getChildIndex(firstPara);
  var endIdx   = body.getChildIndex(lastPara);

  if (startIdx < 0 || endIdx < 0 || startIdx > endIdx) {
    ui.alert('Selection Error',
      'Could not locate the selected paragraphs. Please re-select and try again.', ui.ButtonSet.OK);
    return;
  }

  var paras = [];
  for (var i = startIdx; i <= endIdx; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.PARAGRAPH) continue;
    var txt = child.asParagraph().getText();
    paras.push({ idx: i, text: txt, isEmpty: txt.trim() === '' });
  }

  if (paras.length < 2) {
    ui.alert('Nothing to Clean', 'Only one paragraph found in the selection.', ui.ButtonSet.OK);
    return;
  }

  /*
   * Strategy (mirrors the Word triple-pass method from your resource):
   *   • Empty paragraphs  → treated as "real" paragraph breaks → kept
   *   • Consecutive non-empty paragraphs → merged into one with a space
   *   • Lines ending with a hyphen → hyphen removed, words joined directly
   */
  var groups = [], curr = [];
  paras.forEach(function (p) {
    if (p.isEmpty) {
      if (curr.length) { groups.push({ type: 'text', paras: curr }); curr = []; }
      groups.push({ type: 'empty', paras: [p] });
    } else {
      curr.push(p);
    }
  });
  if (curr.length) groups.push({ type: 'text', paras: curr });

  // Work bottom-up so that removing paragraphs doesn't shift earlier indices
  var mergeCount = 0;
  for (var g = groups.length - 1; g >= 0; g--) {
    var grp = groups[g];
    if (grp.type !== 'text' || grp.paras.length < 2) continue;
    mergeCount++;

    var target = body.getChild(grp.paras[0].idx).asParagraph();

    for (var p = 1; p < grp.paras.length; p++) {
      var src      = grp.paras[p].text.trim();
      if (!src) continue;
      var existing = target.getText();

      if (existing.endsWith('-')) {
        // Hyphenated line-end: remove hyphen and join words directly
        target.editAsText().deleteText(existing.length - 1, existing.length - 1);
      } else if (!existing.endsWith(' ')) {
        target.appendText(' ');
      }
      target.appendText(src);
    }

    for (var p = grp.paras.length - 1; p >= 1; p--) {
      body.getChild(grp.paras[p].idx).removeFromParent();
    }
  }

  if (mergeCount === 0) {
    ui.alert('No Changes Made',
      'No unwanted line breaks were found in the selected text.', ui.ButtonSet.OK);
  } else {
    ui.alert('✅ Done',
      'Merged ' + mergeCount + ' paragraph group(s).\n\n' +
      'Tip: Scroll through to confirm it looks right.\n' +
      'Press Ctrl+Z (Cmd+Z on Mac) to undo if needed.',
      ui.ButtonSet.OK);
  }
}

// ——— PART 2 : PRONOUN REPLACEMENT ————————————————————————————————

/*
 * Pronoun sets — index order: [nominative, accusative, reflexive, indGenitive, depGenitive]
 *
 * Ambiguity note:
 *   Some words serve multiple grammatical roles (e.g. "her" is both accusative
 *   and dependent genitive in the feminine set). When a conflict arises during
 *   replacement, the DEPENDENT GENITIVE form wins — most common in clinical
 *   writing ("her history", "his report"). The highlight step lets you review
 *   before committing.
 */
var PRONOUN_SETS = {
  masculine : ['he',   'him',  'himself',    'his',    'his'  ],
  feminine  : ['she',  'her',  'herself',    'hers',   'her'  ],
  epicene   : ['they', 'them', 'themself',   'theirs', 'their'],
  plural    : ['they', 'them', 'themselves', 'theirs', 'their'],
  neuter    : ['it',   'it',   'itself',     'its',    'its'  ]
};

function showPronounDialog() {
  var html = HtmlService.createHtmlOutputFromFile('PronounSidebar')
    .setTitle('🔄 Replace Pronouns');
  DocumentApp.getUi().showSidebar(html);
}

/** Step 1 (called from dialog) — highlight every FROM pronoun form. Returns count. */
function highlightPronouns(fromKey) {
  var body  = DocumentApp.getActiveDocument().getBody();
  var forms = PRONOUN_SETS[fromKey];

  var seen = {}, unique = [];
  forms.forEach(function (f) {
    var lc = f.toLowerCase();
    if (!seen[lc]) { seen[lc] = true; unique.push(lc); }
  });

  var count = 0;
  unique.forEach(function (form) {
    var pattern = '(?i)\\b' + escapeRegex(form) + '\\b';
    var result  = body.findText(pattern);
    while (result) {
      var el = result.getElement().asText();
      var s  = result.getStartOffset();
      var e  = result.getEndOffsetInclusive();

      // Skip if preceded by ^ (manually or sidebar-marked)
      var preceded = isPrecededByCaret(el, s);
      if (!preceded) {
        el.setBackgroundColor(s, e, '#FFE066');
        count++;
      }
      result = body.findText(pattern, result);
    }
  });
  return count;
}

/** Step 2 (called from dialog) — replace pronouns and clear highlights. Returns { count }. */
function replacePronouns(fromKey, toKey, neoForms) {
  var body      = DocumentApp.getActiveDocument().getBody();
  var fromForms = PRONOUN_SETS[fromKey];
  var toForms   = (toKey === 'neo') ? neoForms : PRONOUN_SETS[toKey];

  var map = {};
  for (var i = 0; i < 5; i++) {
    map[fromForms[i].toLowerCase()] = toForms[i];
  }

  var count = 0;

  Object.keys(map).forEach(function(fromWord) {
    var toWord  = map[fromWord];
    if (!toWord) return;
    var pattern = '(?i)\\b' + escapeRegex(fromWord) + '\\b';

    // Pass 1: collect all replaceable matches into an array
    var matches = [];
    var result  = body.findText(pattern);
    while (result) {
      var el = result.getElement().asText();
      var s  = result.getStartOffset();
      var e  = result.getEndOffsetInclusive();

      var preceded = isPrecededByCaret(el, s);
      if (!preceded) {
        matches.push({
          element : el,
          start   : s,
          end     : e,
          original: el.getText().substring(s, e + 1)
        });
      }
      result = body.findText(pattern, result);
    }

    // Pass 2: replace bottom-to-top so offsets stay valid
    for (var m = matches.length - 1; m >= 0; m--) {
      var match = matches[m];
      var repl  = matchCase(match.original, toWord);
      match.element.deleteText(match.start, match.end);
      match.element.insertText(match.start, repl);
      var newE = match.start + repl.length - 1;
      if (newE >= match.start) {
        match.element.setBackgroundColor(match.start, newE, null);
      }
      count++;
    }
  });

  return { count: count };
}

/** Clears all yellow highlights (called by "Clear Highlights" button). */
function clearHighlightedPronouns(fromKey) {
  var body  = DocumentApp.getActiveDocument().getBody();
  var forms = PRONOUN_SETS[fromKey];

  var seen = {}, unique = [];
  forms.forEach(function(f) {
    var lc = f.toLowerCase();
    if (!seen[lc]) { seen[lc] = true; unique.push(lc); }
  });

  unique.forEach(function(form) {
    var pattern = '(?i)\\b' + escapeRegex(form) + '\\b';
    var result  = body.findText(pattern);
    while (result) {
      var el = result.getElement().asText();
      el.setBackgroundColor(result.getStartOffset(), result.getEndOffsetInclusive(), null);
      result = body.findText(pattern, result);
    }
  });
}

/**
 * Fixes verb agreement after "they" before pronoun replacement runs.
 * Targets common irregular verbs immediately following "they".
 * Only meaningful when replacing they → he/she.
 */
function fixVerbsAfterThey() {
  var body = DocumentApp.getActiveDocument().getBody();

  // Pairs: [pattern to find after "they ", replacement verb]
  var verbPairs = [
    ['don\'t',   'doesn\'t'],
    ['haven\'t', 'hasn\'t' ],
    ['weren\'t', 'wasn\'t' ],
    ['aren\'t',  'isn\'t'  ],
    ['don\'t',   'doesn\'t'],
    ['have',     'has'     ],
    ['were',     'was'     ],
    ['are',      'is'      ],
    ['do',       'does'    ]
  ];

  var count = 0;

  verbPairs.forEach(function(pair) {
    var fromVerb = pair[0];
    var toVerb   = pair[1];
    // Match "they [verb]" with word boundary, case-insensitive
    var pattern  = '(?i)\\bthey\\s+' + escapeRegex(fromVerb) + '\\b';

    var result = body.findText(pattern);
    while (result) {
      var el   = result.getElement().asText();
      var s    = result.getStartOffset();
      var e    = result.getEndOffsetInclusive();
      var full = el.getText().substring(s, e + 1);

      // Skip ^-marked instances
      var preceded = isPrecededByCaret(el, s);
      if (preceded) {
        result = body.findText(pattern, result);
        continue;
      }

      // The verb starts after "they " — find where it begins
      var verbStart = s + full.toLowerCase().indexOf(fromVerb.toLowerCase());
      var verbEnd   = verbStart + fromVerb.length - 1;

      // Match case of original verb
      var original = el.getText().substring(verbStart, verbEnd + 1);
      var replacement = matchCase(original, toVerb);

      el.deleteText(verbStart, verbEnd);
      el.insertText(verbStart, replacement);
      count++;

      // Restart search — document changed
      result = body.findText(pattern);
    }
  });

  return { count: count };
}

/**
 * Scrolls the document cursor to the nth instance of the
 * target pronoun (she/he/they etc), so the user can review
 * the following verb manually. Returns total match count.
 */
function scrollToNthPronoun(toKey, index) {
  var body  = DocumentApp.getActiveDocument().getBody();
  var forms = PRONOUN_SETS[toKey];

  var seen = {}, unique = [];
  forms.forEach(function(f) {
    var lc = f.toLowerCase();
    if (!seen[lc]) { seen[lc] = true; unique.push(lc); }
  });

  var matches = [];
  unique.forEach(function(form) {
    var pattern = '(?i)\\b' + escapeRegex(form) + '\\b';
    var result  = body.findText(pattern);
    while (result) {
      var el = result.getElement();
      var s  = result.getStartOffset();
      if (!isPrecededByCaret(el.asText(), s)) {
        matches.push({ element: el, offset: s });
      }
      result = body.findText(pattern, result);
    }
  });

  if (matches.length === 0) return { total: 0 };

  var i   = index % matches.length;
  var pos = DocumentApp.getActiveDocument()
              .newPosition(matches[i].element, matches[i].offset);
  DocumentApp.getActiveDocument().setCursor(pos);

  return { total: matches.length, current: i };
}


/**
 * Highlights all instances of the TO pronoun set in orange
 * so the user can review verb agreement after replacement.
 */
function highlightToPronouns(toKey) {
  var body  = DocumentApp.getActiveDocument().getBody();
  var forms = PRONOUN_SETS[toKey];

  var seen = {}, unique = [];
  forms.forEach(function(f) {
    var lc = f.toLowerCase();
    if (!seen[lc]) { seen[lc] = true; unique.push(lc); }
  });

  var count = 0;
  unique.forEach(function(form) {
    var pattern = '(?i)\\b' + escapeRegex(form) + '\\b';
    var result  = body.findText(pattern);
    while (result) {
      var el = result.getElement().asText();
      var s  = result.getStartOffset();
      var e  = result.getEndOffsetInclusive();
      if (!isPrecededByCaret(el, s)) {
        el.setBackgroundColor(s, e, '#FFD580');
        count++;
      }
      result = body.findText(pattern, result);
    }
  });
  return { count: count };
}


function skipSelectedPronoun() {
  var sel = DocumentApp.getActiveDocument().getSelection();
  if (!sel) return { success: false, message: 'No text selected in document.' };

  var elems = sel.getRangeElements();
  if (!elems || elems.length === 0) return { success: false, message: 'Could not read selection.' };

  var re = elems[0];
  var el = re.getElement();

  if (el.getType() !== DocumentApp.ElementType.TEXT) {
    return { success: false, message: 'Please select or click within a pronoun.' };
  }

  el = el.asText();
  var s = re.isPartial() ? re.getStartOffset() : 0;
  var e = re.isPartial() ? re.getEndOffsetInclusive() : el.getText().length - 1;

  // Clear yellow highlight from this instance
  if (e >= s) el.setBackgroundColor(s, e, null);

  // Prepend ^ to mark as skip
  el.insertText(s, '^');

  return { success: true };
}

// ——— PART 3 : SNIPPET SIDEBAR ————————————————————————————————————

function showSnippetSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SnippetSidebar')
    .setTitle('📋 Diagnostic Snippets')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Finds the table cell containing {{slotId}} and appends bullet
 * list items — one per non-empty line in bulletTexts array.
 * The placeholder is NOT removed (so it stays findable for future
 * appends). Call clearPlaceholders() when finalising the report.
 */
function pushSnippet(slotId, bulletTexts) {
  if (!bulletTexts || bulletTexts.length === 0) {
    return { success: false, message: 'Nothing to push.' };
  }

  var body        = DocumentApp.getActiveDocument().getBody();
  var placeholder = '{{' + slotId + '}}';
  var found       = body.findText(escapeRegex(placeholder));

  if (!found) {
    return {
      success : false,
      message : 'Placeholder ' + placeholder + ' not found in document.'
    };
  }

  // Walk up the element tree to reach the TableCell
  var node = found.getElement().getParent(); // usually Paragraph
  var cell = null;

  if (node.getType() === DocumentApp.ElementType.TABLE_CELL) {
    cell = node.asTableCell();
  } else {
    var up = node.getParent();
    if (up && up.getType() === DocumentApp.ElementType.TABLE_CELL) {
      cell = up.asTableCell();
    }
  }

  if (!cell) {
    return {
      success : false,
      message : 'Could not locate table cell for ' + placeholder + '.'
    };
  }

  var pushed = 0;
  bulletTexts.forEach(function (line) {
    var trimmed = line.trim();
    if (trimmed) {
      var item = cell.appendListItem(trimmed);
      item.setGlyphType(DocumentApp.GlyphType.BULLET);
      pushed++;
    }
  });

  return {
    success : true,
    message : 'Pushed ' + pushed + ' bullet(s) to ' + slotId + '.'
  };
}

/**
 * Pushes multiple slots in one call.
 * snippetsObj shape: { 'ASD-A1': ['line1', 'line2'], 'ADHD-A1a': ['line3'] }
 */
function pushAllSnippets(snippetsObj) {
  var results = [];
  Object.keys(snippetsObj).forEach(function (slotId) {
    var lines = snippetsObj[slotId];
    if (lines && lines.length > 0) {
      var r    = pushSnippet(slotId, lines);
      r.slotId = slotId;
      results.push(r);
    }
  });
  return results;
}

/**
 * Strips every {{...}} placeholder from the document.
 * Run this when the report is finalised.
 */
function clearPlaceholders() {
  var body  = DocumentApp.getActiveDocument().getBody();
  var count = 0;

  // Strip {{placeholders}}
  var pattern = '\\{\\{[^}]+\\}\\}';
  var result  = body.findText(pattern);
  while (result) {
    var el = result.getElement().asText();
    el.deleteText(result.getStartOffset(), result.getEndOffsetInclusive());
    count++;
    result = body.findText(pattern);
  }

  // Strip ^ skip markers
  var caretCount = 0;
  result = body.findText('\\^');
  while (result) {
    var el = result.getElement().asText();
    el.deleteText(result.getStartOffset(), result.getEndOffsetInclusive());
    caretCount++;
    result = body.findText('\\^');
  }

  DocumentApp.getUi().alert(
    '✅ Done',
    'Removed ' + count + ' placeholder(s) and ' + caretCount + ' skip marker(s) from the document.',
    DocumentApp.getUi().ButtonSet.OK
  );
}

// ——— PART 4 : SCORE TABLE —————————————————————————————

/*
 * Single source of truth for all scored measures.
 * To add a measure: append one object to the correct group.
 * Nothing else needs to change — the table builder and sidebar
 * both generate themselves from this array automatically.
 *
 * Fields:
 *   group  — display group: 'General' | 'ADHD' | 'Autism' | 'Other-Report'
 *   abbr   — short label used as the row anchor for score writing
 *   full   — full measure name
 */
var SCORE_MEASURES = [
  // ── General ──────────────────────────────────────────────────────
  { group: 'General',      abbr: 'WSAS',        full: 'Work and Social Adjustment Scale' },
  { group: 'General',      abbr: 'GAD-7',       full: 'General Anxiety Disorder 7-item version' },
  { group: 'General',      abbr: 'PHQ-9',       full: 'Patient Health Questionnaire 9-item version' },
  { group: 'General',      abbr: 'MAIA-2',      full: 'Multidimensional Assessment of Interoceptive Awareness' },
  { group: 'General',      abbr: 'DERS',        full: 'Difficulties in Emotion Regulation Scale' },
  // ── ADHD ─────────────────────────────────────────────────────────
  { group: 'ADHD',         abbr: 'WFIRS-S',     full: 'Weiss Functional Impairment Rating Scale' },
  { group: 'ADHD',         abbr: 'BDEFS-LF-S',  full: 'Barkley Deficits in Executive Functioning – Long Form (Self-report)' },
  { group: 'ADHD',         abbr: 'ESQ-R',       full: 'Executive Skills Questionnaire-Revised' },
  { group: 'ADHD',         abbr: 'AAMM',        full: 'Adult ADHD Masking Measure' },
  { group: 'ADHD',         abbr: 'WURS-25',     full: 'Wender Utah Rating Scale 25-item version' },
  { group: 'ADHD',         abbr: 'ASRS',        full: 'Adult ADHD Self Report Scale, version 1.1' },
  // ── Autism ───────────────────────────────────────────────────────
  { group: 'Autism',       abbr: 'CAT-Q',       full: 'Camouflaging Autistic Traits Questionnaire' },
  { group: 'Autism',       abbr: 'AQ',          full: 'Autism Quotient' },
  { group: 'Autism',       abbr: 'RAADS-R',     full: 'Ritvo Autism Asperger Diagnostic Scale-Revised' },
  { group: 'Autism',       abbr: 'CATI',        full: 'Comprehensive Autistic Trait Inventory' },
  // ── Other-Report ─────────────────────────────────────────────────
  { group: 'Other-Report', abbr: 'BDEFS-LF-OR', full: 'Barkley Deficits in Executive Functioning – Long Form (Other-report)' }
];

// Group display order and colours
var GROUP_META = {
  'General'     : { label: 'General — Self-Report',  color: '#e8eaf6' },
  'ADHD'        : { label: 'ADHD — Self-Report',      color: '#e3f2fd' },
  'Autism'      : { label: 'Autism — Self-Report',      color: '#f3e5f5' },
  'Other-Report': { label: 'Other-Report',            color: '#fce4ec' }
};
var GROUP_ORDER = ['General', 'ADHD', 'Autism', 'Other-Report'];

/** Called from menu — inserts the score table at {{SCORE-TABLE}}. */
function buildScoreTable() {
  var ui   = DocumentApp.getUi();
  var body = DocumentApp.getActiveDocument().getBody();

  var found = body.findText(escapeRegex('{{SCORE-TABLE}}'));
  if (!found) {
    ui.alert('Placeholder not found',
      '{{SCORE-TABLE}} was not found in the document.\n\n' +
      'Add it to where you want the table to appear, then try again.',
      ui.ButtonSet.OK);
    return;
  }

  // Walk up to the containing paragraph
  var el = found.getElement();
  while (el.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    el = el.getParent();
  }
  var para = el.asParagraph();

  // Clear the placeholder text from the paragraph
  para.editAsText().setText('');

  var paraIndex = body.getChildIndex(para);
  var insertAt  = paraIndex + 1;

  // ── Build table rows ──────────────────────────────────────────────
  // Header row first
  var allRows = [['Measure', 'Name', 'Score', 'Level / Percentile', 'Notes']];

  GROUP_ORDER.forEach(function (group) {
    // Group subheader row (we mark it with a special sentinel in col 0)
    allRows.push(['__GROUP__' + group, '', '', '', '']);
    SCORE_MEASURES.forEach(function (m) {
      if (m.group === group) {
        allRows.push([m.abbr, m.full, '', '', '']);
      }
    });
  });

  var table = body.insertTable(insertAt, allRows);

  // ── Style the table ───────────────────────────────────────────────
  // Column widths (approx, in points — 1 inch = 72 pt)
  // Total usable width ~450pt for a standard doc
  var colWidths = [72, 144, 50, 90, 108];

  // Header row styling
  var headerRow = table.getRow(0);
  for (var c = 0; c < 5; c++) {
    var cell = headerRow.getCell(c);
    cell.setBackgroundColor('#37474f');
    cell.editAsText().setBold(true);
    try { table.setColumnWidth(c, colWidths[c]); } catch(e) {}
  }

  // Walk remaining rows and style group headers vs data rows
  var rowIndex = 1;
  GROUP_ORDER.forEach(function (group) {
    var meta     = GROUP_META[group];
    var groupRow = table.getRow(rowIndex);
    var cell0    = groupRow.getCell(0);

    // Replace sentinel text with real label
    cell0.editAsText().setText(meta.label);
    cell0.editAsText().setBold(true);

    // Merge visually by colouring all cells in the group header row
    for (var c = 0; c < 5; c++) {
      groupRow.getCell(c).setBackgroundColor(meta.color);
    }
    rowIndex++;

    SCORE_MEASURES.forEach(function (m) {
      if (m.group === group) {
        // Abbr cell — bold
        table.getRow(rowIndex).getCell(0).editAsText().setBold(true);
        rowIndex++;
      }
    });
  });

  ui.alert('✅ Score table built',
    'The score table has been inserted.\n\n' +
    'Open "📊 Score Entry…" from the menu to fill in scores per client.',
    ui.ButtonSet.OK);
}

/** Returns the SCORE_MEASURES array to the sidebar. */
function getScoreMeasures() {
  return SCORE_MEASURES;
}

/**
 * Writes score data into the score table.
 * scoresObj shape: { 'WSAS': { score: '14', severity: 'Moderate', notes: '...' }, ... }
 */
function saveScores(scoresObj) {
  var body    = DocumentApp.getActiveDocument().getBody();
  var results = [];

  Object.keys(scoresObj).forEach(function (abbr) {
    var data    = scoresObj[abbr];
    var pattern = '^' + escapeRegex(abbr) + '$';
    var found   = body.findText(pattern);

    if (!found) {
      results.push({ abbr: abbr, success: false, message: 'Row not found.' });
      return;
    }

    // Walk up to TableCell → TableRow
    var node = found.getElement().getParent();
    while (node && node.getType() !== DocumentApp.ElementType.TABLE_ROW) {
      node = node.getParent();
    }
    if (!node) {
      results.push({ abbr: abbr, success: false, message: 'Could not locate table row.' });
      return;
    }

    var row = node.asTableRow();
    // Columns: 0=abbr, 1=full name, 2=score, 3=severity, 4=notes
    if (data.score    !== undefined) row.getCell(2).editAsText().setText(data.score);
    if (data.severity !== undefined) row.getCell(3).editAsText().setText(data.severity);
    if (data.notes    !== undefined) row.getCell(4).editAsText().setText(data.notes);

    results.push({ abbr: abbr, success: true });
  });
  return results;
}

function markMeasuresCompleted(abbrs) {
  var ehrId = getLinkedEhrId();
  if (!ehrId) return;

  var sheet = SpreadsheetApp.openById(CONTROL_PANEL_ID)
                .getSheetByName(CASES_TAB);
  if (!sheet) return;

  var data      = sheet.getDataRange().getValues();
  var header    = data[0];
  var ehrIdx    = header.indexOf('EhrID');
  var compIdx   = header.indexOf('CompletedForms');
  if (ehrIdx < 0 || compIdx < 0) return;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ehrIdx]).trim() !== String(ehrId).trim()) continue;

    var existing = String(data[i][compIdx] || '').trim();
    var completed = existing
      ? new Set(existing.split(',').map(function(s){ return s.trim(); }).filter(Boolean))
      : new Set();

    abbrs.forEach(function(a){ completed.add(a); });
    sheet.getRange(i + 1, compIdx + 1).setValue([...completed].join(', '));
    return;
  }
}

// ——— PART 4b : GENERATE DATA SOURCES LIST ————————————————————————

/**
 * Reads the score table to find which measures have scores entered,
 * then inserts a formatted list at {{DATA-SOURCES}}.
 * Only measures with a non-empty Score cell are included.
 */
function generateDataSources() {
  var ui   = DocumentApp.getUi();
  var body = DocumentApp.getActiveDocument().getBody();

  // ── Step 1: Find which measures have scores in the table ──────────
  var scoredAbbrs = {};
  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    var table = child.asTable();
    for (var r = 1; r < table.getNumRows(); r++) {
      var row = table.getRow(r);
      if (row.getNumCells() < 3) continue;
      var abbrText  = row.getCell(0).getText().trim();
      var scoreText = row.getCell(2).getText().trim();
      if (abbrText && scoreText) {
        scoredAbbrs[abbrText] = true;
      }
    }
  }

  if (Object.keys(scoredAbbrs).length === 0) {
    ui.alert(
      'No scores found',
      'No scores have been entered in the score table yet.\n\n' +
      'Enter scores using "📊 Score Entry…" first, then generate the list.',
      ui.ButtonSet.OK
    );
    return;
  }

  // ── Step 2: Find the placeholder ──────────────────────────────────
  var found = body.findText(escapeRegex('{{DATA-SOURCES}}'));
  if (!found) {
    ui.alert(
      'Placeholder not found',
      '{{DATA-SOURCES}} was not found in the document.\n\n' +
      'Add it where you want the data sources list to appear, then try again.',
      ui.ButtonSet.OK
    );
    return;
  }

  // Walk up to containing paragraph
  var el = found.getElement();
  while (el.getType() !== DocumentApp.ElementType.PARAGRAPH) {
    el = el.getParent();
  }
  var anchorPara  = el.asParagraph();
  var anchorIndex = body.getChildIndex(anchorPara);

  // Clear the placeholder text
  anchorPara.editAsText().setText('');

  // ── Step 3: Build paragraph list ──────────────────────────────────
  // Each entry: { text, bold, italic }
  var paras = [];
  var selfReportGroups = ['General', 'ADHD', 'Autism'];
  var hasSelfReport = selfReportGroups.some(function (group) {
    return SCORE_MEASURES.some(function (m) {
      return m.group === group && scoredAbbrs[m.abbr];
    });
  });

  if (hasSelfReport) {
    paras.push({ text: 'Self-Report', bold: true, italic: false });

    selfReportGroups.forEach(function (group) {
      var groupMeasures = SCORE_MEASURES.filter(function (m) {
        return m.group === group && scoredAbbrs[m.abbr];
      });
      if (groupMeasures.length === 0) return;

      paras.push({ text: group + ':', bold: true, italic: false });

      groupMeasures.forEach(function (m) {
        paras.push({
          text   : m.full + ' (' + m.abbr + ')',
          bold   : false,
          italic : true
        });
      });
    });
  }

  // Other-Report section
  var orMeasures = SCORE_MEASURES.filter(function (m) {
    return m.group === 'Other-Report' && scoredAbbrs[m.abbr];
  });
  if (orMeasures.length > 0) {
    paras.push({ text: 'Other-Report', bold: true, italic: false });
    orMeasures.forEach(function (m) {
      paras.push({
        text   : m.full + ' (' + m.abbr + ')',
        bold   : false,
        italic : true
      });
    });
  }

  // ── Step 4: Insert paragraphs into document ───────────────────────
  paras.forEach(function (p, idx) {
    var newPara = body.insertParagraph(anchorIndex + 1 + idx, p.text);
    var txt     = newPara.editAsText();
    txt.setBold(p.bold);
    txt.setItalic(p.italic);
  });

  ui.alert(
    '✅ Data sources list generated',
    'Inserted ' + paras.length + ' line(s) at {{DATA-SOURCES}}.\n\n' +
    'Only measures with scores entered in the score table are included.\n\n' +
    'Tip: Run "🗑️ Remove Placeholders" when the report is finalised.',
    ui.ButtonSet.OK
  );
}

function showScoreSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('ScoreSidebar')
    .setTitle('📊 Score Entry')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

// ——— PART 4c : TIDY SCORE TABLE ——————————————————————————————————

function tidyScoreTable() {
  var body        = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();
  var scoreTable  = null;

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    var t = child.asTable();
    if (t.getNumRows() > 0 &&
        t.getRow(0).getNumCells() > 0 &&
        t.getRow(0).getCell(0).getText().trim() === 'Measure') {
      scoreTable = t;
      break;
    }
  }

  if (!scoreTable) return;

  var knownAbbrs = {};
  SCORE_MEASURES.forEach(function(m) { knownAbbrs[m.abbr] = true; });

  // Build set of group label texts for identification
  var groupLabels = {};
  GROUP_ORDER.forEach(function(g) {
    groupLabels[GROUP_META[g].label] = g;
  });

  // Pass 1: Remove unscored measure rows (bottom-up)
  for (var r = scoreTable.getNumRows() - 1; r >= 1; r--) {
    var row   = scoreTable.getRow(r);
    if (row.getNumCells() < 5) continue;
    var cell0 = row.getCell(0).getText().trim();
    if (!knownAbbrs[cell0]) continue;

    var score    = row.getCell(2).getText().trim();
    var severity = row.getCell(3).getText().trim();
    var notes    = row.getCell(4).getText().trim();
    if (!score && !severity && !notes) {
      scoreTable.removeRow(r);
    }
  }

  // Pass 2: Remove orphaned group header rows (bottom-up)
  for (var r = scoreTable.getNumRows() - 1; r >= 1; r--) {
    var row   = scoreTable.getRow(r);
    if (row.getNumCells() < 1) continue;
    var cell0 = row.getCell(0).getText().trim();

    // Is this a group header row?
    if (!(cell0 in groupLabels)) continue;

    // Check if next row is another group header, the column header, or doesn't exist
    var nextRow = r + 1;
    var isOrphaned = nextRow >= scoreTable.getNumRows();
    if (!isOrphaned) {
      var nextCell0 = scoreTable.getRow(nextRow).getCell(0).getText().trim();
      isOrphaned = (nextCell0 in groupLabels);
    }
    if (isOrphaned) scoreTable.removeRow(r);
  }

  // Pass 3: Re-style remaining group header rows
  for (var r = 1; r < scoreTable.getNumRows(); r++) {
    var row   = scoreTable.getRow(r);
    if (row.getNumCells() < 5) continue;
    var cell0 = row.getCell(0).getText().trim();

    if (!(cell0 in groupLabels)) continue;

    var groupKey = groupLabels[cell0];
    var color    = GROUP_META[groupKey].color;

    for (var c = 0; c < 5; c++) {
      var cell = row.getCell(c);
      cell.setBackgroundColor(color);
      cell.editAsText().setBold(true).setFontSize(11);
    }

    // Merge visual appearance — make label span full width visually
    row.getCell(0).editAsText().setText(cell0);
  }
}

// ——— FINALIZE DOCUMENT ———————————————————————————————————————————

function finalizeDocument() {
  var ui = DocumentApp.getUi();
  var response = ui.alert(
    '⚠️ Finalize Document',
    'This will:\n\n' +
    '• Tidy the score table (remove unscored rows)\n' +
    '• Strip all ^ pronoun skip markers\n' +
    '• Remove all {{placeholders}}\n' +
    '• Remove white/near-white text backgrounds\n' +
    '• Normalize near-black font colors to true black\n' +
    '• Clear resource insertion history\n\n' +
    'This cannot be undone. Continue?',
    ui.ButtonSet.OK_CANCEL
  );

  if (response !== ui.Button.OK) return;

  // Step 1: Tidy score table
  tidyScoreTable();

  // Step 2: Strip ^ markers and {{placeholders}}
  var body         = DocumentApp.getActiveDocument().getBody();
  var caretCount   = 0;
  var placeholderCount = 0;

  var result = body.findText('\\^');
  while (result) {
    var el = result.getElement().asText();
    el.deleteText(result.getStartOffset(), result.getEndOffsetInclusive());
    caretCount++;
    result = body.findText('\\^');
  }

  result = body.findText('\\{\\{[^}]+\\}\\}');
  while (result) {
    var el = result.getElement().asText();
    el.deleteText(result.getStartOffset(), result.getEndOffsetInclusive());
    placeholderCount++;
    result = body.findText('\\{\\{[^}]+\\}\\}');
  }

  // Step 2b: Clear inserted-resources tracking for this document
  PropertiesService.getScriptProperties().deleteAllProperties();

  // Step 3: Strip white/near-white backgrounds and normalize dark fonts
  var bgCount   = 0;
  var fontCount = 0;

  var numChildren = body.getNumChildren();
  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var para = child.asParagraph();
      var text = para.editAsText();
      var len  = para.getText().length;
      if (len === 0) continue;

      for (var c = 0; c < len; c++) {
        // Background color
        var bg = text.getBackgroundColor(c);
        if (bg && (isNearWhite(bg) || isScriptHighlight(bg))) {
          text.setBackgroundColor(c, c, null);
          bgCount++;
        }

        // Font color
        var fc = text.getForegroundColor(c);
        if (fc && isNearBlackNonBlue(fc)) {
          text.setForegroundColor(c, c, '#000000');
          fontCount++;
        }
                
      }
    }
  }


  ui.alert(
    '✅ Document Finalized',
    'Completed:\n\n' +
    '• Score table tidied\n' +
    '• ' + caretCount + ' skip marker(s) removed\n' +
    '• ' + placeholderCount + ' placeholder(s) removed\n' +
    '• ' + bgCount + ' white background(s) cleared\n' +
    '• ' + fontCount + ' font color(s) normalized to black\n' +
    '• Resource insertion history cleared',
    ui.ButtonSet.OK
  );
}

/**
 * Returns true if a hex color is white or near-white.
 * Threshold: all RGB channels >= 240.
 */
function isNearWhite(hex) {
  if (!hex || hex.charAt(0) !== '#' || hex.length < 7) return false;
  var r = parseInt(hex.slice(1,3), 16);
  var g = parseInt(hex.slice(3,5), 16);
  var b = parseInt(hex.slice(5,7), 16);
  return r >= 240 && g >= 240 && b >= 240;
}

/**
 * Returns true if a hex color is near-black but not true black
 * and not a blue/link color.
 * Near-black: all channels <= 80, not #000000.
 * Blue: R < G or R < B by significant margin — skip those.
 */
function isNearBlackNonBlue(hex) {
  if (!hex || hex.charAt(0) !== '#' || hex.length < 7) return false;
  if (hex.toLowerCase() === '#000000') return false;
  var r = parseInt(hex.slice(1,3), 16);
  var g = parseInt(hex.slice(3,5), 16);
  var b = parseInt(hex.slice(5,7), 16);
  // Skip blues — blue channel dominant or link-style colors
  if (b > r + 30 || b > g + 30) return false;
  // Near-black: all channels dark
  return r <= 80 && g <= 80 && b <= 80;
}

/**
 * Returns true if the color is one of the script's own
 * highlight colors (pronoun yellow or verb-review orange).
 */
function isScriptHighlight(hex) {
  if (!hex) return false;
  var lower = hex.toLowerCase();
  return lower === '#ffe066' || lower === '#ffd580';
}

/**
 * Moves the cursor to the score table,
 * which causes the document view to scroll to it.
 */
function scrollToScoreTable() {
  var body       = DocumentApp.getActiveDocument().getBody();
  var numChildren = body.getNumChildren();

  for (var i = 0; i < numChildren; i++) {
    var child = body.getChild(i);
    if (child.getType() !== DocumentApp.ElementType.TABLE) continue;
    var t = child.asTable();
    if (t.getNumRows() > 0 &&
        t.getRow(0).getNumCells() > 0 &&
        t.getRow(0).getCell(0).getText().trim() === 'Measure') {
      var cell   = t.getRow(0).getCell(0);
      var pos    = DocumentApp.getActiveDocument()
                    .newPosition(cell.getChild(0), 0);
      DocumentApp.getActiveDocument().setCursor(pos);
      return { success: true };
    }
  }
  return { success: false };
}


// ——— PART 5 : CASE LINKING ———————————————————————————————————————

/**
 * Stores the EHR ID in the document's Properties so it persists
 * across sessions. Called from SetupSidebar when the user confirms.
 */
function linkCase(ehrId) {
  PropertiesService.getDocumentProperties().setProperty('EHR_ID', ehrId.trim());
  return { success: true };
}

/**
 * Returns the stored EHR ID, or null if not yet set.
 */
function getLinkedEhrId() {
  return PropertiesService.getDocumentProperties().getProperty('EHR_ID');
}

/**
 * Looks up the case row by EHR ID and returns ActiveForms and an alias.
 * Returns { found, alias, activeForms } where activeForms is an array.
 */
function lookupCase(ehrId) {
  var sheet = SpreadsheetApp.openById(CONTROL_PANEL_ID)
                .getSheetByName(CASES_TAB);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowEhr = String(data[i][COL_EHR_ID]).trim();
    if (rowEhr === String(ehrId).trim()) {
      var alias       = String(data[i][0]).trim();
      var activeRaw   = String(data[i][COL_ACTIVE_FORMS]).trim();
      var activeForms = activeRaw
        ? activeRaw.split(',').map(function(s){ return s.trim(); }).filter(Boolean)
        : [];
      var clientNovoForms = String(data[i][COL_CLIENT_NOVO_FORMS]).trim();
      var collatNovoForms = String(data[i][COL_COLLAT_NOVO_FORMS]).trim();
      var clientNovoList  = clientNovoForms
        ? clientNovoForms.split(',').map(function(s){ return s.trim(); }).filter(Boolean)
        : [];
      var collatNovoList  = collatNovoForms
        ? collatNovoForms.split(',').map(function(s){ return s.trim(); }).filter(Boolean)
        : [];

      return {
        found           : true,
        alias           : alias,
        activeForms     : activeForms,
        clientNovoForms : clientNovoList,
        collatNovoForms : collatNovoList
      };
    }
  }
  return { found: false };
}

/**
 * Reads the Measures tab and returns an array of
 * measure objects in the same shape as SCORE_MEASURES.
 * Skips rows with empty Abbr.
 */
function getNovoPsychMeasures() {
  var sheet = SpreadsheetApp.openById(CONTROL_PANEL_ID)
                .getSheetByName(NOVO_MEASURES_TAB);
  if (!sheet) return [];

  var data     = sheet.getDataRange().getValues();
  var measures = [];
  var validGroups = ['General', 'ADHD', 'Autism', 'Sensory', 'Other-Report'];

  for (var i = 1; i < data.length; i++) {
    var abbr  = String(data[i][0]).trim();
    var full  = String(data[i][1]).trim();
    var group = String(data[i][2]).trim();
    if (!abbr) continue;
    if (validGroups.indexOf(group) === -1) group = 'General';
    measures.push({ group: group, abbr: abbr, full: full });
  }
  return measures;
}

/**
 * Appends newly completed measure abbrs to the CompletedForms cell
 * for the linked case. Deduplicates — won't add if already present.
 */
function markMeasuresCompleted(completedAbbrs) {
  var ehrId = getLinkedEhrId();
  if (!ehrId) return { success: false, message: 'No case linked to this document.' };

  var sheet = SpreadsheetApp.openById(CONTROL_PANEL_ID)
                .getSheetByName(CASES_TAB);
  var data  = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    var rowEhr = String(data[i][COL_EHR_ID]).trim();
    if (rowEhr !== String(ehrId).trim()) continue;

    var existing = String(data[i][COL_COMPLETED_FORMS]).trim();
    var already  = existing
      ? existing.split(',').map(function(s){ return s.trim(); }).filter(Boolean)
      : [];

    // Add only abbrs not already present
    var toAdd = completedAbbrs.filter(function(a){
      return already.indexOf(a) === -1;
    });

    if (toAdd.length === 0) return { success: true, added: 0 };

    var updated = already.concat(toAdd).join(', ');
    // COL_COMPLETED_FORMS is 0-indexed, sheet column is +1
    sheet.getRange(i + 1, COL_COMPLETED_FORMS + 1).setValue(updated);
    return { success: true, added: toAdd.length };
  }

  return { success: false, message: 'EHR ID not found in Cases sheet.' };
}

/**
 * Returns only the SCORE_MEASURES entries whose abbr appears in
 * ActiveForms for the linked case. Falls back to all measures if
 * no case is linked or ActiveForms is empty.
 */
/**
 * Returns the list of measures to show in the Score Entry sidebar.
 * Merges SCORE_MEASURES with NovoPsych measures from the sheet,
 * deduplicating by abbr (case-insensitive, typo-tolerant).
 * Filters to only ActiveForms + NovoPsych forms for the linked case.
 */
function getActiveMeasures() {
  var ehrId = getLinkedEhrId();

  // Get NovoPsych measures from sheet
  var novoMeasures = getNovoPsychMeasures();

  // Build a normalized lookup of existing SCORE_MEASURES abbrs
  var existingAbbrs = {};
  SCORE_MEASURES.forEach(function(m) {
    existingAbbrs[m.abbr.toLowerCase().replace(/[^a-z0-9]/g, '')] = true;
  });

  // Merge in NovoPsych measures that aren't already in SCORE_MEASURES
  var merged = SCORE_MEASURES.slice();
  novoMeasures.forEach(function(nm) {
    var normalized = nm.abbr.toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!existingAbbrs[normalized]) {
      merged.push(nm);
    }
  });

  // If no case linked, return full merged list
  if (!ehrId) return merged;

  var caseData = lookupCase(ehrId);
  if (!caseData.found) return merged;

  // Combine ActiveForms + clientNovoForms + collatNovoForms
  var allActive = caseData.activeForms
    .concat(caseData.clientNovoForms)
    .concat(caseData.collatNovoForms);

  if (allActive.length === 0) return merged;

  // Normalize active forms for fuzzy matching
  var normalizedActive = allActive.map(function(a) {
    return a.toLowerCase().replace(/[^a-z0-9]/g, '');
  });

  return merged.filter(function(m) {
    var normalized = m.abbr.toLowerCase().replace(/[^a-z0-9]/g, '');
    return normalizedActive.indexOf(normalized) !== -1;
  });
}

/**
 * Opens the setup sidebar. Called from menu and from onOpen
 * if no case is linked yet.
 */
function showSetupSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('SetupSidebar')
    .setTitle('⚙️ Setup')
    .setWidth(300);
  DocumentApp.getUi().showSidebar(html);
}

function showMeasuresSidebar() {
  var ehrId = getLinkedEhrId();
  if (!ehrId) {
    DocumentApp.getUi().alert('No case linked. Please run Setup / Link Case first.');
    return;
  }

  var sheet = SpreadsheetApp.openById(CONTROL_PANEL_ID).getSheetByName(CASES_TAB);
  var data  = sheet.getDataRange().getValues();
  var header = data[0];
  var ehrIdx       = header.indexOf('EhrID');
  var clientIdx    = header.indexOf('Client_NovoForms');
  var collatIdx    = header.indexOf('Collat_NovoForms');
  var completedIdx = header.indexOf('CompletedForms');

  var clientForms = '', collatForms = '', completedForms = '';
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ehrIdx]).trim() === String(ehrId).trim()) {
      clientForms    = String(data[i][clientIdx]    || '');
      collatForms    = String(data[i][collatIdx]    || '');
      completedForms = String(data[i][completedIdx] || '');
      break;
    }
  }

  var measures = getNovoPsychMeasures();

  var html = HtmlService.createHtmlOutput(
    buildMeasuresSidebarHtml_(ehrId, measures, clientForms, collatForms, completedForms)
  ).setTitle('Measures').setWidth(360);
  DocumentApp.getUi().showSidebar(html);
}

function buildMeasuresSidebarHtml_(ehrId, measures, clientForms, collatForms, completedForms) {
  var navy   = '#172B36';
  var green  = '#4A9B6F';
  var pink   = '#E7AA98';
  var ltblue = '#DAE9F1';
  var amber  = '#C17A3A';

  var clientSet    = new Set(clientForms.split(',').map(function(s){ return s.trim(); }).filter(Boolean));
  var collatSet    = new Set(collatForms.split(',').map(function(s){ return s.trim(); }).filter(Boolean));
  var completedSet = new Set(completedForms.split(',').map(function(s){ return s.trim(); }).filter(Boolean));

  var groupOrder = ['General', 'ADHD', 'Autism', 'Sensory', 'Other-Report'];
  var grouped = {};
  groupOrder.forEach(function(g){ grouped[g] = []; });
  measures.forEach(function(m){
    var g = m.group && grouped[m.group] !== undefined ? m.group : 'General';
    grouped[g].push(m);
  });

  var tableHtml =
    '<table style="width:100%;border-collapse:collapse;font-size:12px;">' +
    '<tr>' +
    '<th style="text-align:left;padding:4px 0;">Measure</th>' +
    '<th style="text-align:center;width:44px;">Client</th>' +
    '<th style="text-align:center;width:52px;">Collat</th>' +
    '<th style="text-align:center;width:52px;">Done</th>' +
    '</tr>';

  groupOrder.forEach(function(g){
    if (grouped[g].length === 0) return;
    tableHtml +=
      '<tr><td colspan="4" style="padding:6px 0 2px 0;font-size:10px;font-weight:bold;' +
      'text-transform:uppercase;letter-spacing:.5px;color:#555;border-bottom:1px solid #ccc;">' +
      g + '</td></tr>';
    grouped[g].forEach(function(m){
      var cChk    = clientSet.has(m.abbr)    ? ' checked' : '';
      var oChk    = collatSet.has(m.abbr)    ? ' checked' : '';
      var doneChk = completedSet.has(m.abbr) ? ' checked' : '';
      tableHtml +=
        '<tr style="border-bottom:1px solid #f0f0f0;">' +
        '<td style="padding:5px 0;" title="' + m.full + '">' + m.abbr + '</td>' +
        '<td style="text-align:center;">' +
          '<input type="checkbox" name="novo_client" value="' + m.abbr + '"' + cChk +
          ' style="accent-color:' + navy + ';">' +
        '</td>' +
        '<td style="text-align:center;">' +
          '<input type="checkbox" name="novo_collat" value="' + m.abbr + '"' + oChk +
          ' style="accent-color:' + navy + ';">' +
        '</td>' +
        '<td style="text-align:center;">' +
          '<input type="checkbox" name="novo_done" value="' + m.abbr + '"' + doneChk +
          ' style="accent-color:' + green + ';">' +
        '</td>' +
        '</tr>';
    });
  });
  tableHtml += '</table>';

  var css =
    'body{font-family:Arial,sans-serif;font-size:13px;margin:0;padding:12px;background:#fff;color:' + navy + ';}' +
    'h3{margin:0 0 4px 0;font-size:15px;color:' + navy + ';}' +
    '.sub{font-size:11px;color:#666;margin-bottom:14px;}' +
    '.btn{display:block;width:100%;margin-top:18px;padding:10px;background:' + navy + ';color:#fff;' +
      'border:none;border-radius:8px;font-size:14px;cursor:pointer;}' +
    '.btn:hover{background:#0f1e27;}' +
    '.status{margin-top:10px;font-size:13px;color:' + green + ';text-align:center;min-height:18px;}';

  var safeId = ehrId.replace(/"/g, '&quot;');

  return '<!DOCTYPE html><html><head><base target="_top"><style>' + css + '</style></head><body>' +
    '<h3>Measures</h3>' +
    '<div class="sub">EhrID: ' + safeId + '</div>' +
    tableHtml +
    '<button class="btn" onclick="saveMeasures()">Save Changes</button>' +
    '<div class="status" id="status"></div>' +
    '<script>' +
    'function saveMeasures() {' +
    '  var clientForms  = Array.from(document.querySelectorAll("input[name=novo_client]:checked")).map(function(b){ return b.value; });' +
    '  var collatForms  = Array.from(document.querySelectorAll("input[name=novo_collat]:checked")).map(function(b){ return b.value; });' +
    '  var doneForms    = Array.from(document.querySelectorAll("input[name=novo_done]:checked")).map(function(b){ return b.value; });' +
    '  document.getElementById("status").textContent = "Saving\u2026";' +
    '  google.script.run' +
    '    .withSuccessHandler(function(msg){' +
    '      document.getElementById("status").textContent = msg;' +
    '    })' +
    '    .withFailureHandler(function(err){' +
    '      document.getElementById("status").textContent = "Error: " + err.message;' +
    '      document.getElementById("status").style.color = "red";' +
    '    })' +
    '    .saveMeasuresFromReport("' + safeId + '", clientForms, collatForms, doneForms);' +
    '}' +
    '<\/script></body></html>';
}

function saveMeasuresFromReport(ehrId, clientForms, collatForms, doneForms) {
  var sheet = SpreadsheetApp.openById(CONTROL_PANEL_ID).getSheetByName(CASES_TAB);
  var data  = sheet.getDataRange().getValues();
  var header = data[0];
  var ehrIdx       = header.indexOf('EhrID');
  var clientIdx    = header.indexOf('Client_NovoForms');
  var collatIdx    = header.indexOf('Collat_NovoForms');
  var completedIdx = header.indexOf('CompletedForms');

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][ehrIdx]).trim() !== String(ehrId).trim()) continue;
    var row = i + 1;
    sheet.getRange(row, clientIdx + 1).setValue(clientForms.join(', '));
    sheet.getRange(row, collatIdx + 1).setValue(collatForms.join(', '));

    // Merge doneForms into existing CompletedForms without overwriting
    var existing = String(data[i][completedIdx] || '').trim();
    var completed = existing
      ? new Set(existing.split(',').map(function(s){ return s.trim(); }).filter(Boolean))
      : new Set();
    doneForms.forEach(function(a){ completed.add(a); });
    sheet.getRange(row, completedIdx + 1).setValue([...completed].join(', '));
    return 'Saved \u2713';
  }
  return 'Error: case not found.';
}

function showCGSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('CGSidebar')
    .setTitle('Coventry Grid — Lived Experience')
    .setWidth(320);
  DocumentApp.getUi().showSidebar(html);
}

function pushCGSnippet(slotId, text) {
  var body        = DocumentApp.getActiveDocument().getBody();
  var placeholder = '{{' + slotId + '}}';
  var found       = body.findText(escapeRegex(placeholder));
  if (!found) throw new Error('Placeholder ' + placeholder + ' not found in document.');

  var cell = found.getElement().getParent();
  while (cell && cell.getType() !== DocumentApp.ElementType.TABLE_CELL) {
    cell = cell.getParent();
  }
  if (!cell) throw new Error('Placeholder is not inside a table cell.');

  var lines = text.split('\n').map(function(l){ return l.trim(); }).filter(Boolean);

  // Find and remove the placeholder paragraph
  var tc        = cell.asTableCell();
  var numParas  = tc.getNumChildren();
  for (var i = 0; i < numParas; i++) {
    var child = tc.getChild(i);
    if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
      var pt = child.asParagraph().getText();
      if (pt.indexOf(placeholder) !== -1) {
        child.asParagraph().editAsText().setText('');
        break;
      }
    }
  }

  // Append each line as a bullet in the Specifically: style
  lines.forEach(function(line) {
    var para = tc.appendParagraph(line);
    para.setGlyphType(DocumentApp.GlyphType.BULLET);
    para.editAsText().setBold(false);
  });
}

function diagnosePlaceholder() {
  const doc  = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  
  const result = body.findText('\\{\\{RESOURCES\\}\\}');
  if (!result) {
    Logger.log('Placeholder not found');
    return;
  }
  
  const element = result.getElement();
  const parent  = element.getParent();
  const grandparent = parent.getParent();
  
  Logger.log('Element type: ' + element.getType());
  Logger.log('Parent type: ' + parent.getType());
  Logger.log('Grandparent type: ' + grandparent.getType());
  Logger.log('Grandparent is Body?: ' + (grandparent.getType() === DocumentApp.ElementType.BODY_SECTION));
  
  // Try to see if it's in a section
  try {
    const section = parent.getParent();
    Logger.log('Section type: ' + section.getType());
  } catch(e) {
    Logger.log('No section element');
  }
}