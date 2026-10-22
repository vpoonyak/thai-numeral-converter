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
 * Performs Global Replacement for digits 0-9 in a range.
 * Mimics VBA Find.Execute Replace:=wdReplaceAll
 */
async function replaceAllDigitsInRange(range: Word.Range, useSmartIgnore: boolean, context: Word.RequestContext) {
  // Search for all digit blocks (or alphanumeric words, so Smart Ignore can skip spin9, 9arm, ...)
  const searchPattern = useSmartIgnore ? "[a-zA-Z0-9]{1,}" : "[0-9]{1,}";
  const results = range.search(searchPattern, { matchWildcards: true });
  results.load("items/text");
  await context.sync();

  // Queue every replacement, then sync once
  for (const blockRange of results.items) {
    if (useSmartIgnore && !/^[0-9]+$/.test(blockRange.text)) continue;
    blockRange.insertText(toThaiDigits(blockRange.text), "Replace");
  }
  await context.sync();
}

/**
 * Processes a "Story" (Body, Header, Footer, Shape Body)
 */
async function processStory(body: Word.Body, useSmartIgnore: boolean, context: Word.RequestContext, warnings: string[]) {
  if (!body) return;

  // 1. Handle Fields (Page Numbers, Total Pages)
  // VBA approach: Update field codes with \* ThaiArabic
  try {
    const fields = body.fields;
    fields.load("items/type,items/code");
    await context.sync();

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

  // 2. Perform Global Digit Replacement in this story
  try {
    await replaceAllDigitsInRange(body.getRange(), useSmartIgnore, context);
  } catch (e) {
    warn(warnings, "ตัวเลขในเนื้อหา (Text)", e);
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
            await replaceAllDigitsInRange(tFrame.textRange, useSmartIgnore, context);
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
  await Word.run(async (context: Word.RequestContext) => {
    const selection = context.document.getSelection();
    await replaceAllDigitsInRange(selection, useSmartIgnore, context);
  });
  return [];
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
