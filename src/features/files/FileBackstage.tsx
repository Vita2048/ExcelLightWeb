import type { LocalWorkspaceRecord } from '../../persistence/local';

export type FileAction = 'new' | 'example' | 'open' | 'save' | 'saveAs' | 'export' | 'close' | 'options';

export function FileBackstage({ open, busy, recent, activeWorkspaceId, storageError,
  onAction, onBack, onRecent, onRemoveRecent }: {
  readonly open: boolean;
  readonly busy: boolean;
  readonly recent: readonly LocalWorkspaceRecord[];
  readonly activeWorkspaceId: string;
  readonly storageError?: string;
  readonly onAction: (action: FileAction) => void;
  readonly onBack: () => void;
  readonly onRecent: (record: LocalWorkspaceRecord) => void;
  readonly onRemoveRecent: (record: LocalWorkspaceRecord) => void;
}) {
  if (!open) return null;
  const commands: readonly [FileAction, string][] = [
    ['new', 'New'], ['open', 'Open'], ['save', 'Save'], ['saveAs', 'Save As'],
    ['export', 'Export'], ['close', 'Close'],
  ];
  return <div className="file-backstage" role="dialog" aria-modal="true" aria-label="File Home">
    <nav className="file-backstage-nav" aria-label="File commands">
      <button type="button" className="file-back" autoFocus onClick={onBack}>‹ <span>Back</span></button>
      <button type="button" className="file-home-nav" aria-current="page">Home</button>
      <div className="file-command-list">
        {commands.map(([action, label]) => <button key={action} type="button" disabled={busy}
          onClick={() => onAction(action)}>{label}</button>)}
      </div>
      <button type="button" className="file-options" disabled={busy} onClick={() => onAction('options')}>Options</button>
    </nav>
    <section className="file-home-page" aria-label="Home">
      <h1>Home</h1>
      <h2>New</h2>
      <div className="file-new-gallery">
        <button type="button" className="file-new-card" aria-label="Blank workbook" disabled={busy} onClick={() => onAction('new')}>
          <span className="file-template-preview blank-workbook-preview" aria-hidden="true">
            <BlankWorkbookPreview />
          </span>
          <span className="file-new-card-label">Blank workbook</span>
        </button>
        <button type="button" className="file-new-card" aria-label="Example: Support SLA Dashboard" disabled={busy}
          onClick={() => onAction('example')}>
          <span className="file-template-preview sla-dashboard-preview" aria-hidden="true">
            <SupportSlaDashboardPreview />
          </span>
          <span className="file-new-card-label">Example: Support SLA Dashboard</span>
        </button>
      </div>
      <h2 className="recent-heading">Recent</h2>
      {storageError && <p className="file-storage-warning" role="status">{storageError}</p>}
      {recent.length === 0 ? <p className="file-empty-recent">No recent workbooks</p> :
        <ul className="file-recent-list">
          {recent.map(record => <li key={record.id}>
            <div className="file-recent-entry">
              <button type="button" className="file-recent-open" disabled={busy} aria-label={'Open recent ' + record.name}
                onClick={() => onRecent(record)}>
                <span className="file-recent-name">{record.name}</span>
                <span className="file-recent-location">{record.fileHandle ? `Linked file · ${record.fileHandle.name}` : 'Saved in this browser'}</span>
              </button>
              <button type="button" className="file-recent-remove" disabled={busy || record.id === activeWorkspaceId}
                aria-label={'Remove ' + record.name + ' from this browser'}
                title={record.id === activeWorkspaceId ? 'Close this workbook before removing its browser copy' : 'Remove only the browser-local copy'}
                onClick={() => onRemoveRecent(record)}>Remove</button>
            </div>
          </li>)}
        </ul>}
    </section>
  </div>;
}

function BlankWorkbookPreview() {
  return <svg viewBox="0 0 560 320" role="presentation" focusable="false">
    <rect width="560" height="320" fill="#f1f1f1" />
    <rect width="560" height="38" fill="#d8d8d8" />
    <rect y="38" width="46" height="282" fill="#e6e6e6" />
    <path d="M46 0v320 M148 0v320 M251 0v320 M354 0v320 M457 0v320 M0 38h560 M0 73h560 M0 108h560 M0 143h560 M0 178h560 M0 213h560 M0 248h560 M0 283h560 M0 318h560"
      fill="none" stroke="#aaa" strokeWidth="1" />
    <g fill="#333" fontFamily="Arial, sans-serif" fontSize="14" textAnchor="middle">
      <text x="97" y="25">A</text><text x="200" y="25">B</text><text x="303" y="25">C</text><text x="406" y="25">D</text>
    </g>
    <g fill="#555" fontFamily="Arial, sans-serif" fontSize="12" textAnchor="middle">
      <text x="23" y="61">1</text><text x="23" y="96">2</text><text x="23" y="131">3</text><text x="23" y="166">4</text>
      <text x="23" y="201">5</text><text x="23" y="236">6</text><text x="23" y="271">7</text><text x="23" y="306">8</text>
    </g>
  </svg>;
}

function SupportSlaDashboardPreview() {
  return <svg viewBox="0 0 560 320" role="presentation" focusable="false">
    <rect width="560" height="320" fill="#f3f7fa" />
    <rect x="12" y="12" width="372" height="68" rx="3" fill="#142d46" />
    <rect x="396" y="12" width="152" height="68" rx="3" fill="#142d46" />
    <g fontFamily="Arial, sans-serif">
      <text x="24" y="31" fill="#85d8cb" fontSize="10" fontWeight="700">SERVICE OPERATIONS / 01</text>
      <text x="24" y="55" fill="#fff" fontSize="21" fontWeight="700">Support ticket SLA</text>
      <text x="24" y="70" fill="#ccdfef" fontSize="8">JAN–JUN 2026 · FIRST-RESPONSE PERFORMANCE</text>
      <text x="408" y="31" fill="#85d8cb" fontSize="9" fontWeight="700">SCENARIO INPUTS / 02</text>
      <text x="408" y="53" fill="#fff" fontSize="13" fontWeight="700">Response targets</text>
    </g>
    <g>
      <rect x="12" y="92" width="84" height="84" rx="3" fill="#fff" />
      <rect x="104" y="92" width="84" height="84" rx="3" fill="#fff" />
      <rect x="196" y="92" width="84" height="84" rx="3" fill="#fff" />
      <rect x="288" y="92" width="96" height="84" rx="3" fill="#fff" />
      <rect x="408" y="92" width="140" height="142" rx="3" fill="#fff" />
    </g>
    <g fontFamily="Arial, sans-serif">
      <g fill="#425c73" fontSize="8" fontWeight="700">
        <text x="20" y="110">TICKETS OPENED</text><text x="112" y="110">BREACH RATE</text>
        <text x="204" y="110">OPEN AT SNAPSHOT</text><text x="296" y="110">P90 RESOLUTION</text>
      </g>
      <g fontSize="21" fontWeight="700">
        <text x="20" y="143" fill="#142d46">610</text><text x="112" y="143" fill="#087f81">24.8%</text>
        <text x="204" y="143" fill="#142d46">8</text><text x="296" y="143" fill="#142d46">77 h</text>
      </g>
      <g fill="#647b8f" fontSize="8">
        <text x="20" y="160">In six-month cohort</text><text x="112" y="160">Of assessed tickets</text>
        <text x="204" y="160">Awaiting resolution</text><text x="296" y="160">Resolved tickets</text>
      </g>
      <rect x="408" y="92" width="140" height="23" rx="3" fill="#142d46" />
      <g fill="#fff" fontSize="8" fontWeight="700">
        <text x="416" y="107">PRIORITY</text><text x="458" y="107">TARGET HRS</text><text x="505" y="107">EXAMPLE</text>
      </g>
      <g fontSize="8">
        <rect x="408" y="115" width="140" height="28" fill="#fff" />
        <rect x="408" y="143" width="140" height="28" fill="#eaf1f7" />
        <rect x="408" y="171" width="140" height="28" fill="#fff" />
        <rect x="408" y="199" width="140" height="28" fill="#eaf1f7" />
        <g fill="#142d46" fontWeight="700">
          <text x="416" y="133">P1</text><text x="416" y="161">P2</text><text x="416" y="189">P3</text><text x="416" y="217">P4</text>
        </g>
        <g fill="#087f81" fontWeight="700">
          <text x="463" y="133">1</text><text x="463" y="161">4</text><text x="463" y="189">12</text><text x="463" y="217">24</text>
        </g>
        <g fill="#425c73">
          <text x="486" y="133">Outage</text><text x="486" y="161">Blocked</text><text x="486" y="189">Degraded</text><text x="486" y="217">Question</text>
        </g>
      </g>
      <rect x="12" y="192" width="372" height="104" rx="3" fill="#fff" />
      <text x="24" y="216" fill="#142d46" fontSize="11" fontWeight="700">How the SLA is scored</text>
      <text x="24" y="238" fill="#425c73" fontSize="9">Breach · First reply exceeds its target or passes the deadline.</text>
      <text x="24" y="257" fill="#425c73" fontSize="9">Pending tickets with time remaining are excluded from the rate.</text>
      <text x="24" y="276" fill="#425c73" fontSize="9">P90 resolution includes tickets resolved by the snapshot.</text>
    </g>
  </svg>;
}
