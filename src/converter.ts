/* global Word */

export const ARABIC_TO_THAI_MAP: { [key: string]: string } = {
  '0': '๐', '1': '๑', '2': '๒', '3': '๓', '4': '๔',
  '5': '๕', '6': '๖', '7': '๗', '8': '๘', '9': '๙',
};

const MONTHS_MAP: { [key: string]: string } = {
  'January': 'มกราคม', 'February': 'กุมภาพันธ์', 'March': 'มีนาคม',
  'April': 'เมษายน', 'May': 'พฤษภาคม', 'June': 'มิถุนายน',
  'July': 'กรกฎาคม', 'August': 'สิงหาคม', 'September': 'กันยายน',
  'October': 'ตุลาคม', 'November': 'พฤศจิกายน', 'December': 'ธันวาคม',
  'Jan': 'ม.ค.', 'Feb': 'ก.พ.', 'Mar': 'มี.ค.', 'Apr': 'เม.ย.', 'Jun': 'มิ.ย.',
  'Jul': 'ก.ค.', 'Aug': 'ส.ค.', 'Sep': 'ก.ย.', 'Oct': 'ต.ค.', 'Nov': 'พ.ย.', 'Dec': 'ธ.ค.'
};

function toThaiDigits(text: string): string {
  return text.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
}

/**
 * Collects the parts of a conversion that failed, so the UI can report them
 * instead of claiming full success.
 */
function warn(warnings: string[], label: string, error: unknown) {
  console.warn(label, error);
  if (warnings.some((w) => w.startsWith(label))) return;
  // OfficeExtension.Error carries a code (e.g. ApiNotFound) and where it failed
  const err = error as { code?: string; message?: string; debugInfo?: { errorLocation?: string } };
  const detail = [err?.code, err?.debugInfo?.errorLocation].filter(Boolean).join(" @ ") || err?.message || String(error);
  warnings.push(`${label} [${detail}]`);
}

/**
 * Converts Arabic numerals to Thai numerals in a string.
 */
export function convertText(text: string, useSmartIgnore: boolean): string {
  let processedText = text;
  if (useSmartIgnore) {
    const dateRegex = /\b(\d{1,2})?\s?(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s?(\d{1,2})?,?\s(20\d{2})\b/gi;
    processedText = processedText.replace(dateRegex, (_, d1, month, d2, year) => {
      const thaiMonth = MONTHS_MAP[month] || month;
      const beYear = parseInt(year) + 543;
      const day = d1 || d2 || "";
      return `${day} ${thaiMonth} ${beYear}`.trim();
    });
    const smartRegex = /(?<![a-zA-Z0-9])[0-9]+(?![a-zA-Z0-9])/g;
    return processedText.replace(smartRegex, (match: string) => {
      return match.split('').map((char: string) => ARABIC_TO_THAI_MAP[char] || char).join('');
    });
  } else {
    return processedText.replace(/[0-9]/g, (match: string) => ARABIC_TO_THAI_MAP[match] || match);
  }
}

/**
 * Decides, for each digit run in `text` (in order), whether it should be converted.
 * Smart Ignore is applied here in JavaScript instead of with a Word wildcard like
 * [a-zA-Z0-9]{1,}, which Word 16.114 on Mac answers with an empty result.
 */
function digitRunPlan(text: string, useSmartIgnore: boolean): boolean[] {
  const runs = [...text.matchAll(/[0-9]+/g)];
  if (!useSmartIgnore) return runs.map(() => true);
  const convertible = new Set([...text.matchAll(/(?<![a-zA-Z0-9])[0-9]+(?![a-zA-Z0-9])/g)].map((m) => m.index));
  return runs.map((m) => convertible.has(m.index));
}

/**
 * Converts the digits in each target (paragraph or range) whose `text` is already loaded.
 * Word's "[0-9]{1,}" search returns the same maximal digit runs, in the same order,
 * as /[0-9]+/g over the text, so the plan lines up with the search results.
 * Mimics VBA Find.Execute Replace:=wdReplaceAll
 */
async function replaceDigitsIn(targets: (Word.Paragraph | Word.Range)[], useSmartIgnore: boolean, context: Word.RequestContext, warnings: string[]) {
  const jobs: { plan: boolean[]; results: Word.RangeCollection }[] = [];
  for (const target of targets) {
    const plan = digitRunPlan(target.text ?? "", useSmartIgnore);
    if (!plan.some(Boolean)) continue;
    const results = target.search("[0-9]{1,}", { matchWildcards: true });
    results.load("items/text");
    jobs.push({ plan, results });
  }
  if (jobs.length === 0) return;
  await context.sync();

  // Queue every replacement, then sync once
  for (const { plan, results } of jobs) {
    if (!results.items || results.items.length !== plan.length) {
      warn(warnings, "บางย่อหน้า (Some paragraphs)", new Error(`expected ${plan.length} digit runs, Word found ${results.items?.length}`));
      continue;
    }
    results.items.forEach((digits, i) => {
      if (plan[i]) digits.insertText(toThaiDigits(digits.text), "Replace");
    });
  }
  await context.sync();
}

async function replaceDigitsInRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext, warnings: string[]) {
  range.load("text");
  await context.sync();
  await replaceDigitsIn([range], useSmartIgnore, context, warnings);
}

/**
 * Processes a "Story" (Body, Header, Footer, Shape Body)
 */
async function processStory(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext, warnings: string[]) {
  if (!body) return;

  // 1. Convert digits paragraph by paragraph. Done first so that a failure in the
  // field or shape steps below can't stop the main text from converting.
  try {
    const paragraphs = body.paragraphs;
    paragraphs.load("items/text");
    await context.sync();
    await replaceDigitsIn(paragraphs.items, useSmartIgnore, context, warnings);
  } catch (e) {
    warn(warnings, "ตัวเลขในเนื้อหา (Text)", e);
  }

  // 2. Handle Fields (Page Numbers, Total Pages)
  // VBA approach: Update field codes with \* ThaiArabic
  try {
    // Load the bare list first: Word 16.114 on Mac answers it, but can return null
    // for "items/type,items/code", so only ask for details when there are fields
    const fields = body.fields;
    fields.load("items");
    await context.sync();
    if (fields.items.length > 0) {
      fields.load("items/type,items/code");
      await context.sync();
    }

    for (let i = 0; i < fields.items.length; i++) {
      const field = fields.items[i];
      const type = field.type.toString().toLowerCase();
      // Page = 33, NumPages = 26
      if (type === "page" || type === "numpages") {
        try {
          if (!field.code.includes("ThaiArabic")) {
            field.code = field.code.trimEnd() + " \\* ThaiArabic ";
          }
          field.updateResult();
          // Office.js only reports API errors on sync, so sync here to reach the fallback
          await context.sync();
        } catch (e) {
          // If we can't update code (WordApi < 1.5), fallback to flattening
          try {
            const res = field.result;
            res.load("text");
            await context.sync();
            if (res.text) {
              const thai = convertText(res.text, false);
              field.unlink();
              res.insertText(thai, "Replace");
              await context.sync();
            }
          } catch (err) {
            warn(warnings, "เลขหน้า (Page fields)", err);
          }
        }
      }
    }
  } catch (e) {
    warn(warnings, "ฟิลด์ (Fields)", e);
  }

  // 3. Recurse into Shapes (Textboxes) in this story
  try {
    const shapes = body.shapes;
    shapes.load("items");
    await context.sync();
    for (let i = 0; i < shapes.items.length; i++) {
      const shape = shapes.items[i];
      try {
        const sBody = (shape as any).body;
        sBody.load("type");
        await context.sync();
        await processStory(sBody, useSmartIgnore, context, warnings);
      } catch (e) {
        // Fallback for shapes without a full 'body' property
        try {
          const tFrame = (shape as any).textFrame;
          tFrame.load("hasText");
          await context.sync();
          if (tFrame.hasText) {
            await replaceDigitsInRange(tFrame.textRange, useSmartIgnore, context, warnings);
          }
        } catch (err) {
          warn(warnings, "กล่องข้อความ (Text boxes)", err);
        }
      }
    }
  } catch (e) {
    warn(warnings, "รูปร่าง (Shapes)", e);
  }
}

/**
 * Main Conversion Functions
 */
export async function convertSelection(useSmartIgnore: boolean): Promise<string[]> {
  const warnings: string[] = [];
  await Word.run(async (context: Word.RequestContext) => {
    await replaceDigitsInRange(context.document.getSelection(), useSmartIgnore, context, warnings);
  });
  return warnings;
}

export async function convertMainBody(useSmartIgnore: boolean): Promise<string[]> {
  const warnings: string[] = [];
  await Word.run(async (context: Word.RequestContext) => {
    await processStory(context.document.body, useSmartIgnore, context, warnings);
    await context.sync();
  });
  return warnings;
}

/**
 * THE ULTIMATE CONVERTER (Inspired by VBA)
 */
export async function flattenAdvancedElements(useSmartIgnore: boolean): Promise<string[]> {
  const warnings: string[] = [];
  await Word.run(async (context: Word.RequestContext) => {
    // 1. FREEZE AUTOMATIC LISTS (ActiveDocument.ConvertNumbersToText)
    try {
      (context.document as any).convertNumbersToText("AllNumbers");
      await context.sync();
    } catch (e) {
      // Manual fallback if API not supported
      const body = context.document.body;
      const paragraphs = body.paragraphs;
      paragraphs.load("items/isListItem,items/listItem");
      await context.sync();
      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        if (para.isListItem && para.listItem) {
          try {
            para.listItem.load("listString");
            await context.sync();
            const label = para.listItem.listString;
            para.detachFromList();
            para.insertText(convertText(label, false) + " ", "Start");
          } catch (err) {
            warn(warnings, "เลขหัวข้อ (List numbers)", err);
          }
        }
      }
    }

    // 2. PROCESS ALL STORY RANGES
    // A. Main Body
    await processStory(context.document.body, useSmartIgnore, context, warnings);

    // B. Headers and Footers in all sections
    const sections = context.document.sections;
    sections.load("items");
    await context.sync();

    for (let i = 0; i < sections.items.length; i++) {
      const section = sections.items[i];
      const hfTypes: Word.HeaderFooterType[] = [
        Word.HeaderFooterType.primary,
        Word.HeaderFooterType.firstPage,
        Word.HeaderFooterType.evenPages
      ];
      for (const type of hfTypes) {
        try { await processStory(section.getHeader(type), useSmartIgnore, context, warnings); } catch (e) { warn(warnings, "หัวกระดาษ (Headers)", e); }
        try { await processStory(section.getFooter(type), useSmartIgnore, context, warnings); } catch (e) { warn(warnings, "ท้ายกระดาษ (Footers)", e); }
      }
    }

    await context.sync();
  });
  return warnings;
}

/**
 * Read-only probes that show which Word API calls return data on this host.
 * Each probe runs in its own Word.run so one failure doesn't hide the others.
 */
export async function runDiagnostics(): Promise<string[]> {
  const describe = (v: unknown) => (v === undefined ? "undefined" : v === null ? "null" : JSON.stringify(v).slice(0, 40));
  const probes: [string, (context: Word.RequestContext) => Promise<unknown>][] = [
    ["body.text", async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return body.text?.length;
    }],
    ["paragraphs (items)", async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();
      return paragraphs.items?.length;
    }],
    ["paragraphs (text)", async (context) => {
      const paragraphs = context.document.body.paragraphs;
      paragraphs.load("text");
      await context.sync();
      return paragraphs.items?.map((p) => p.text.length).length;
    }],
    ["first paragraph", async (context) => {
      const first = context.document.body.paragraphs.getFirst();
      first.load("text");
      await context.sync();
      return first.text;
    }],
    ["search \"1\"", async (context) => {
      const results = context.document.body.search("1");
      results.load("items");
      await context.sync();
      return results.items?.length;
    }],
    ["search [0-9]{1,}", async (context) => {
      const results = context.document.body.getRange().search("[0-9]{1,}", { matchWildcards: true });
      results.load("items/text");
      await context.sync();
      return results.items?.length;
    }],
    ["selection.text", async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return selection.text?.length;
    }],
    ["fields", async (context) => {
      const fields = context.document.body.fields;
      fields.load("items");
      await context.sync();
      return fields.items?.length;
    }],
    // The probes below mirror exactly what processStory() does
    ["fields (type,code)", async (context) => {
      const fields = context.document.body.fields;
      fields.load("items/type,items/code");
      await context.sync();
      return fields.items?.length;
    }],
    ["search [a-zA-Z0-9]{1,}", async (context) => {
      const results = context.document.body.getRange().search("[a-zA-Z0-9]{1,}", { matchWildcards: true });
      results.load("items/text");
      await context.sync();
      return results.items?.length;
    }],
    ["shapes", async (context) => {
      const shapes = context.document.body.shapes;
      shapes.load("items");
      await context.sync();
      return shapes.items?.length;
    }],
    ["same run: fields(type,code) → search [0-9]", async (context) => {
      const body = context.document.body;
      const fields = body.fields;
      fields.load("items/type,items/code");
      await context.sync();
      const results = body.getRange().search("[0-9]{1,}", { matchWildcards: true });
      results.load("items/text");
      await context.sync();
      return `${describe(fields.items?.length)} → ${describe(results.items?.length)}`;
    }],
    ["same run: search [0-9] → shapes", async (context) => {
      const body = context.document.body;
      const results = body.getRange().search("[0-9]{1,}", { matchWildcards: true });
      results.load("items/text");
      await context.sync();
      const shapes = body.shapes;
      shapes.load("items");
      await context.sync();
      return `${describe(results.items?.length)} → ${describe(shapes.items?.length)}`;
    }],
  ];

  const lines: string[] = [];
  for (const [name, probe] of probes) {
    try {
      lines.push(`${name}: ${describe(await Word.run(probe))}`);
    } catch (e) {
      const err = e as { code?: string; message?: string };
      lines.push(`${name}: ERROR ${err?.code ?? ""} ${err?.message ?? String(e)}`.trim());
    }
  }
  for (const set of ["WordApi 1.1", "WordApi 1.5", "WordApiDesktop 1.1", "WordApiDesktop 1.2"]) {
    const [setName, version] = set.split(" ");
    lines.push(`${set}: ${Office.context.requirements.isSetSupported(setName, version) ? "yes" : "no"}`);
  }
  return lines;
}
