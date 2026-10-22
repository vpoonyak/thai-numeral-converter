import { useState } from 'react';
import { DefaultButton, Pivot, PivotItem, PrimaryButton, TextField, Toggle } from '@fluentui/react';
import { convertText } from './converter';
import { Brand, Footer } from './Brand';

const MANIFEST_URL = `${import.meta.env.BASE_URL}manifest.xml`;

function ManifestLink() {
  return (
    <>
      <a href={MANIFEST_URL}>ดาวน์โหลดไฟล์ manifest.xml</a>
      <div className="note">คลิกขวาที่ลิงก์ แล้วเลือก “Save Link As…” (บันทึกลิงก์เป็น…)</div>
    </>
  );
}

function Landing() {
  const [sample, setSample] = useState('หนังสือเลขที่ 123/2567 ลงวันที่ 5 มกราคม 2567 อ้างอิง spin9 และ www.site123.com');
  const [smart, setSmart] = useState(true);

  return (
    <main className="site">
      <Brand />

      <section className="hero">
        <h2>แปลงเลขอารบิกเป็นเลขไทยใน Word ในคลิกเดียว</h2>
        <p>
          ส่วนเสริม (Add-in) สำหรับ Microsoft Word ที่แปลงตัวเลขทั้งเอกสาร รวมถึงเลขหน้าและเลขหัวข้อ
          โดยข้ามเลขในคำภาษาอังกฤษและ URL ให้อัตโนมัติ และคงรูปแบบตัวอักษรเดิมไว้
        </p>
        <div className="hero-actions">
          <PrimaryButton text="วิธีติดตั้ง" href="#install" />
          <DefaultButton text="ลองแปลงตัวอย่าง" href="#demo" />
        </div>
      </section>

      <section aria-labelledby="features-title">
        <h3 id="features-title">ทำอะไรได้บ้าง</h3>
        <ul className="features">
          <li>
            <strong>Smart Ignore</strong>
            <span className="muted">ข้ามเลขในคำอังกฤษ เช่น spin9, iPhone 15 และ URL/อีเมล</span>
          </li>
          <li>
            <strong>Flatten ทั้งเอกสาร</strong>
            <span className="muted">หัว/ท้ายกระดาษ เลขหน้า และเลขหัวข้ออัตโนมัติ (1.1 → ๑.๑)</span>
          </li>
          <li>
            <strong>คงรูปแบบเดิม</strong>
            <span className="muted">ไม่เปลี่ยนฟอนต์ สี ตัวหนา หรือขนาดตัวอักษร</span>
          </li>
        </ul>
      </section>

      <section id="demo" className="demo" aria-labelledby="demo-title">
        <h3 id="demo-title">ลองแปลงตัวอย่าง</h3>
        <TextField
          label="ข้อความ"
          multiline
          autoAdjustHeight
          value={sample}
          onChange={(_, v) => setSample(v ?? '')}
        />
        <Toggle label="Smart Ignore" inlineLabel checked={smart} onChange={(_, c) => setSmart(!!c)} />
        <div>
          <div className="option-label" id="demo-result">ผลลัพธ์</div>
          <output className="demo-output" aria-labelledby="demo-result" style={{ display: 'block' }}>
            {convertText(sample, smart)}
          </output>
        </div>
      </section>

      <section id="install" aria-labelledby="install-title">
        <h3 id="install-title">วิธีติดตั้ง</h3>
        <Pivot aria-label="เลือกระบบปฏิบัติการ">
          <PivotItem headerText="macOS">
            <ol className="steps">
              <li>
                <ManifestLink />
              </li>
              <li>
                คัดลอกไฟล์ไปไว้ที่ <code>~/Library/Containers/com.microsoft.Word/Data/Documents/wef/</code>
                <div className="note">ถ้ายังไม่มีโฟลเดอร์ <code>wef</code> ให้สร้างขึ้นใหม่</div>
              </li>
              <li>ปิด Word ให้สนิท (⌘Q) แล้วเปิดใหม่</li>
              <li>
                ไปที่ <strong>Insert → Add-ins → My Add-ins</strong> เลือกหัวข้อ <strong>Developer Add-ins</strong>{' '}
                แล้วกด <strong>IT๙ Converter</strong>
              </li>
            </ol>
          </PivotItem>
          <PivotItem headerText="Windows">
            <ol className="steps">
              <li>
                สร้างโฟลเดอร์ เช่น <code>C:\Addins</code> แล้วนำไฟล์ manifest.xml ไปวางไว้
                <div style={{ marginTop: 4 }}><ManifestLink /></div>
              </li>
              <li>
                คลิกขวาที่โฟลเดอร์ → <strong>Properties → Sharing → Share</strong> เลือกชื่อผู้ใช้ของคุณ แล้วคัดลอก
                Network Path (เช่น <code>\\ชื่อคอมพิวเตอร์\Addins</code>)
              </li>
              <li>
                ใน Word ไปที่ <strong>File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs</strong>{' '}
                วาง Network Path ในช่อง Catalog URL กด <strong>Add Catalog</strong> ติ๊ก <strong>Show in Menu</strong>{' '}
                แล้วเปิด Word ใหม่
              </li>
              <li>
                ไปที่ <strong>Insert → My Add-ins</strong> แถบ <strong>Shared Folder</strong> แล้วกด{' '}
                <strong>IT๙ Converter</strong>
              </li>
            </ol>
          </PivotItem>
        </Pivot>
      </section>

      <Footer />
    </main>
  );
}

export default Landing;
