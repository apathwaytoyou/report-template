// ============================================================
//  Report Generator  •  A Pathway to You
// ============================================================

const SHEET_ID   = '173Gsuzo6RdzGSjUPlWDg3JwXe8PWY6XblyMzlNJK3zc';
const SHEET_NAME = 'full_list';

const COL = {
  title:      0,
  desc:       1,
  link:       2,
  format:     3,
  population: 4,
  cost:       5,
  pages:      6
};

const PLACEHOLDER = '{{RESOURCES}}';
const PROPS_KEY   = 'insertedLinks_v1';

function splitT_(str) {
  return String(str).split(',').map(s => s.trim()).filter(Boolean);
}

function getResourceData_() {
  const ss    = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_NAME);
  const data  = sheet.getDataRange().getValues();

  return data
    .slice(1)
    .filter(row => row[COL.title])
    .map(row => ({
      title:      String(row[COL.title]).trim(),
      desc:       String(row[COL.desc]).trim(),
      link:       String(row[COL.link]).trim(),
      format:     String(row[COL.format]).trim(),
      population: String(row[COL.population]).trim(),
      cost:       String(row[COL.cost]).trim(),
      pages:      String(row[COL.pages]).trim()
    }));
}

// ── Menu ──────────────────────────────────────────────────────

function openResoListSidebar() {
  const html = HtmlService.createTemplateFromFile('ResoListSidebar')
    .evaluate()
    .setTitle('Resource List Generator');
  DocumentApp.getUi().showSidebar(html);
}

// ── Called by sidebar on load ─────────────────────────────────
function getSidebarData() {
  const ss     = SpreadsheetApp.openById(SHEET_ID);
  const config = ss.getSheetByName('config');
  const data   = config.getDataRange().getValues();

  const populations = {};  // { group: [tags] }
  const formats     = [];
  const costs       = [];
  const pages       = [];

  // Skip header row
  data.slice(1).forEach(row => {
    const type  = String(row[0]).trim();
    const group = String(row[1]).trim();
    const value = String(row[2]).trim();

    if (!value) return;

    if (type === 'population') {
      if (!populations[group]) populations[group] = [];
      populations[group].push(value);
    } else if (type === 'format') {
      formats.push(value);
    } else if (type === 'cost') {
      costs.push(value);
    } else if (type === 'page') {
      pages.push(value);
    }
  });

  return {
    populations: populations,
    formats:     formats,
    costs:       costs,
    pages:       pages
  };
}

// ── Reset tracking for current document ──────────────────────
function resetTracking_() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROPS_KEY);
  return { status: 'success', message: 'Tracking reset — all resources available for insertion.' };
}

// ── Main insert function ──────────────────────────────────────
function insertResources(filters) {
  const resources = getResourceData_();

  // 1. Apply filters
  const matched = resources.filter(r => {
    const rPops = splitT_(r.population).map(s => s.toLowerCase());
    const rFmts = splitT_(r.format).map(s => s.toLowerCase());
    const rCost = r.cost.trim().toLowerCase();

    const popOk  = !filters.populations.length ||
                   filters.populations.some(p => rPops.includes(p.toLowerCase()));
    const fmtOk  = !filters.formats.length ||
                   filters.formats.some(f => rFmts.includes(f.toLowerCase()));
    const costOk = !filters.costs.length ||
                   filters.costs.map(c => c.toLowerCase()).includes(rCost);

    return popOk && fmtOk && costOk;
  });

  if (!matched.length) {
    return { status: 'empty', message: 'No resources matched your filters.' };
  }

  // 2. Remove already-inserted resources
  const props       = PropertiesService.getScriptProperties();
  const insertedSet = new Set(JSON.parse(props.getProperty(PROPS_KEY) || '[]'));
  const toInsert    = matched.filter(r => !insertedSet.has(r.link));

  if (!toInsert.length) {
    return { status: 'empty', message: 'All matching resources are already in the report.' };
  }

  // 3. Build structure: H3 (pop term) → styled para (format) → [resources]
  const h3Terms = [...filters.populations].sort();

  const structure = {};
  h3Terms.forEach(pop => {
    const popLo  = pop.toLowerCase();
    const forPop = toInsert.filter(r =>
      splitT_(r.population).map(s => s.toLowerCase()).includes(popLo)
    );
    if (!forPop.length) return;

    structure[pop] = {};
    forPop.forEach(r => {
      splitT_(r.format).forEach(fmt => {
        if (!structure[pop][fmt])             structure[pop][fmt] = [];
        if (!structure[pop][fmt].includes(r)) structure[pop][fmt].push(r);
      });
    });
  });

  if (!Object.keys(structure).length) {
    return { status: 'empty', message: 'No new resources to insert for the selected populations.' };
  }

  // 4. Find placeholder
  const doc  = DocumentApp.getActiveDocument();
  const body = doc.getBody();

  const placeholderSearch = body.findText('\\{\\{RESOURCES\\}\\}');
  if (!placeholderSearch) {
    return {
      status:  'error',
      message: PLACEHOLDER + ' not found in the document. Please add it where you want resources inserted.'
    };
  }

  const placeholderElement = placeholderSearch.getElement();
  const placeholderPara    = placeholderElement.getParent().asParagraph();

  // 5. Insert content BEFORE the placeholder (it acts as an anchor)
  Object.keys(structure).sort().forEach(pop => {
    // Top-level: H3 (population)
    const h3 = body.insertParagraph(body.getChildIndex(placeholderPara), pop);
    h3.setHeading(DocumentApp.ParagraphHeading.HEADING3);

    let isFirstSub = true;

    Object.keys(structure[pop]).sort().forEach(fmt => {
      // Sub-level: normal paragraph, ALL-CAPS, bold, underlined, centered
      const sub = body.insertParagraph(body.getChildIndex(placeholderPara), fmt.toUpperCase());
      sub.setHeading(DocumentApp.ParagraphHeading.NORMAL);
      sub.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      sub.editAsText().setBold(true).setUnderline(true);

      // Add extra space above, except for the first sub under each population
      if (!isFirstSub) {
        sub.setSpacingBefore(12);  // 12 points ≈ one blank line
      }
      isFirstSub = false;

      structure[pop][fmt].forEach(r => {
        const fullText = r.title
          + (r.desc   ? ' \u2014 ' + r.desc   : '')
          + (r.cost   ? ' ['       + r.cost   + ']' : '')
          + (r.format ? ' ('       + r.format + ')' : '');

        // Resources as bullets, explicitly non-bold, non-underlined
        const item = body.insertListItem(body.getChildIndex(placeholderPara), fullText);
        item.setGlyphType(DocumentApp.GlyphType.BULLET);
        item.editAsText().setBold(false).setUnderline(false);

        if (r.link && r.link.startsWith('http')) {
          item.editAsText().setLinkUrl(0, r.title.length - 1, r.link);
        }
      });
    });
  });

  // 6. Track inserted links
  toInsert.forEach(r => { if (r.link) insertedSet.add(r.link); });
  props.setProperty(PROPS_KEY, JSON.stringify([...insertedSet]));

  // 7. Set cursor at the placeholder (where resources were just inserted)
  const position = doc.newPosition(placeholderPara, 0);
  doc.setCursor(position);

  doc.saveAndClose();

  return {
    status:  'success',
    message: '✓ Added ' + toInsert.length + ' resource'
             + (toInsert.length === 1 ? '' : 's') + ' to your report.'
  };
}

function scrollToHeading(headingText) {
  const doc  = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const result = body.findText(headingText);
  if (!result) return;

  const element = result.getElement().getParent();
  const position = doc.newPosition(element, 0);
  doc.setCursor(position);
}

// ── Diagnostic / utility functions ───────────────────────────

function diagnostic1_resetTracking() {
  const props = PropertiesService.getScriptProperties();
  const before = props.getProperty('insertedLinks_v1') || '[]';
  Logger.log('BEFORE reset, tracked links: ' + JSON.parse(before).length);
  props.deleteProperty('insertedLinks_v1');
  Logger.log('Tracking has been reset.');
}

function diagnostic2_checkPlaceholder() {
  const doc  = DocumentApp.getActiveDocument();
  const body = doc.getBody();
  const result = body.findText('\\{\\{RESOURCES\\}\\}');
  if (!result) {
    Logger.log('PLACEHOLDER NOT FOUND IN DOC');
    return;
  }
  const parent = result.getElement().getParent();
  const grandparent = parent.getParent();
  Logger.log('Parent type: ' + parent.getType());
  Logger.log('Grandparent type: ' + grandparent.getType());
  Logger.log('Is in body? ' + (grandparent.getType() === DocumentApp.ElementType.BODY_SECTION));
}

function diagnostic3_runInsert() {
  const filters = { populations: ['ADHD'], formats: [], costs: [] };
  const result = insertResources(filters);
  Logger.log('Result: ' + JSON.stringify(result));
}