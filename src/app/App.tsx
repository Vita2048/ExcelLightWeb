export function App() {
  return (
    <div className="app-shell">
      <header className="title-bar"><span className="brand-mark" aria-hidden="true">E</span>ExcelLight Web</header>
      <main>
        <p className="eyebrow">Browser workspace</p>
        <h1>Your spreadsheets. On your device.</h1>
        <p className="intro">The web port starts here: a dark workspace for spreadsheets and local Python analysis.</p>
        <section aria-labelledby="status-title" className="status-panel">
          <h2 id="status-title">Application scaffold</h2>
          <p>Spreadsheet editing, file handling, and Python execution are planned and are not available yet.</p>
          <p>The implementation roadmap is in the project’s <code>docs/Implementation-Plan.md</code>.</p>
        </section>
      </main>
      <footer>ExcelLight Web · Development scaffold</footer>
    </div>
  );
}
