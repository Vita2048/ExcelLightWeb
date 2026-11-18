export function DialogActions({ confirmLabel = 'OK', onCancel }: {
  confirmLabel?: string;
  onCancel: () => void;
}) {
  return <div className="dialog-actions">
    <button type="submit">{confirmLabel}</button>
    <button type="button" onClick={onCancel}>Cancel</button>
  </div>;
}
