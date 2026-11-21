import type { LocalWorkspaceRecord } from '../../persistence/local';

export type FileAction = 'new' | 'open' | 'save' | 'saveAs' | 'export' | 'close' | 'options';

export function FileBackstage({ open, busy, recent, activeWorkspaceId, status, recoveryStatus, storageError,
  onAction, onBack, onRecent, onRemoveRecent }: {
  readonly open: boolean;
  readonly busy: boolean;
  readonly recent: readonly LocalWorkspaceRecord[];
  readonly activeWorkspaceId: string;
  readonly status: string;
  readonly recoveryStatus: string;
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
      <button type="button" className="blank-workbook-command" disabled={busy} onClick={() => onAction('new')}>
        Blank workbook
      </button>
      <h2 className="recent-heading">Recent</h2>
      {status && <p className="file-current-status" role="status">{status}</p>}
      {recoveryStatus && <p className="file-recovery-status" role="status">{recoveryStatus}</p>}
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
