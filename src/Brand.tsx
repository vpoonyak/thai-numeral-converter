const BASE = import.meta.env.BASE_URL;
const VERSION = __APP_VERSION__;

export function Brand() {
  return (
    <div className="brand">
      <img src={`${BASE}icon-80.png`} alt="" />
      <div>
        <h1>IT๙ Converter</h1>
        <div className="version">v{VERSION}</div>
      </div>
    </div>
  );
}

export function Footer() {
  return (
    <footer className="footer">
      <a href={`${BASE}privacy.html`} target="_blank" rel="noreferrer">Privacy Policy</a>
      <a href={`${BASE}eula.html`} target="_blank" rel="noreferrer">Terms of Use</a>
      <a href="https://github.com/vpoonyak/thai-numeral-converter" target="_blank" rel="noreferrer">GitHub</a>
      <span>v{VERSION}</span>
    </footer>
  );
}
