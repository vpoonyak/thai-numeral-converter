import { useState } from 'react';
import {
  CompoundButton,
  DefaultButton,
  Dialog,
  DialogFooter,
  DialogType,
  MessageBar,
  MessageBarType,
  PrimaryButton,
  Spinner,
  SpinnerSize,
  Toggle,
} from '@fluentui/react';
import { convertSelection, convertMainBody, flattenAdvancedElements } from './converter';
import { Brand, Footer } from './Brand';

type Status = { message: string; detail?: string; type: MessageBarType };

// Host details to include when something fails, so bug reports say which Word/API level ran
function diagnostics(): string {
  try {
    const d = Office.context.diagnostics;
    const api = ['1.5', '1.4', '1.3', '1.2', '1.1'].find((v) => Office.context.requirements.isSetSupported('WordApi', v));
    return `Word ${d?.version ?? '?'} (${d?.platform ?? '?'}) · WordApi ${api ?? 'none'} · ${location.origin}`;
  } catch {
    return 'diagnostics unavailable';
  }
}

const flattenDialog = {
  type: DialogType.normal,
  title: 'แปลงแบบ Flatten?',
  subText: 'เลขหน้าและเลขหัวข้ออัตโนมัติจะกลายเป็นข้อความถาวร และจะไม่นับเลขใหม่อัตโนมัติอีก (กด ⌘Z / Ctrl+Z เพื่อย้อนกลับได้ทันทีหลังแปลง)',
};

function TaskPane() {
  const [useSmartIgnore, setUseSmartIgnore] = useState(true);
  const [status, setStatus] = useState<Status | null>(null);
  const [isProcessing, setIsProcessing] = useState(false);
  const [confirmFlatten, setConfirmFlatten] = useState(false);

  const handleAction = async (action: () => Promise<string[]>) => {
    try {
      setStatus(null);
      setIsProcessing(true);
      const warnings = await action();
      if (warnings.length > 0) {
        setStatus({
          message: 'แปลงเสร็จแล้ว แต่บางส่วนแปลงไม่สำเร็จ',
          detail: [...warnings, diagnostics()].join('\n'),
          type: MessageBarType.warning,
        });
      } else {
        setStatus({ message: 'แปลงเรียบร้อยแล้ว', type: MessageBarType.success });
      }
    } catch (error) {
      console.error(error);
      setStatus({
        message: 'แปลงไม่สำเร็จ ลองเลือกข้อความใหม่แล้วกดอีกครั้ง',
        detail: error instanceof Error ? error.message : String(error),
        type: MessageBarType.error,
      });
    } finally {
      setIsProcessing(false);
    }
  };

  return (
    <main className="pane">
      <Brand />

      <section className="section" aria-labelledby="options-title">
        <h2 id="options-title" className="section-title">ตัวเลือก</h2>
        <div className="option">
          <div>
            <div className="option-label" id="smart-label">Smart Ignore</div>
            <div className="muted" id="smart-desc">
              ข้ามเลขที่ติดกับอักษรอังกฤษ เช่น spin9, 9arm, www.site123.com
            </div>
          </div>
          <Toggle
            aria-labelledby="smart-label"
            aria-describedby="smart-desc"
            checked={useSmartIgnore}
            onChange={(_, checked) => setUseSmartIgnore(!!checked)}
            disabled={isProcessing}
            styles={{ root: { marginBottom: 0 } }}
          />
        </div>
      </section>

      <section className="section actions" aria-labelledby="convert-title">
        <h2 id="convert-title" className="section-title">แปลงเป็นเลขไทย</h2>
        <CompoundButton
          primary
          text="แปลงเนื้อหาหลัก"
          secondaryText="ตัวเลขในเนื้อหา ไม่รวมหัว/ท้ายกระดาษ"
          onClick={() => handleAction(() => convertMainBody(useSmartIgnore))}
          disabled={isProcessing}
        />
        <CompoundButton
          text="แปลงเฉพาะที่เลือก"
          secondaryText="เฉพาะข้อความที่ไฮไลต์ไว้"
          onClick={() => handleAction(() => convertSelection(useSmartIgnore))}
          disabled={isProcessing}
        />
        <CompoundButton
          text="แปลงรวมเลขหน้าและเลขหัวข้อแบบ Flatten"
          secondaryText="ทั้งเอกสาร รวมหัว/ท้ายกระดาษ เลขหน้า และเลขหัวข้อ"
          onClick={() => setConfirmFlatten(true)}
          disabled={isProcessing}
        />
      </section>

      <div className="status" aria-live="polite">
        {isProcessing && <Spinner size={SpinnerSize.medium} label="กำลังแปลง..." labelPosition="right" />}
        {status && !isProcessing && (
          <MessageBar messageBarType={status.type} isMultiline onDismiss={() => setStatus(null)} dismissButtonAriaLabel="ปิด">
            <strong>{status.message}</strong>
            {status.detail && <div className="muted" style={{ whiteSpace: 'pre-wrap' }}>{status.detail}</div>}
          </MessageBar>
        )}
      </div>

      <Footer />

      <Dialog
        hidden={!confirmFlatten}
        onDismiss={() => setConfirmFlatten(false)}
        dialogContentProps={flattenDialog}
        modalProps={{ isBlocking: false }}
      >
        <DialogFooter>
          <PrimaryButton
            text="แปลงแบบ Flatten"
            onClick={() => {
              setConfirmFlatten(false);
              handleAction(() => flattenAdvancedElements(useSmartIgnore));
            }}
          />
          <DefaultButton text="ยกเลิก" onClick={() => setConfirmFlatten(false)} />
        </DialogFooter>
      </Dialog>
    </main>
  );
}

export default TaskPane;
