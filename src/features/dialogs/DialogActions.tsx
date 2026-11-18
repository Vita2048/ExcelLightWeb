export function DialogActions({ confirmLabel = 'OK', disabled = false, onCancel }: {
  confirmLabel?: string;
  disabled?: boolean;
  onCancel: () => void;
}) {
  return <div className="dialog-actions">
    <button type="submit" disabled={disabled}>{confirmLabel}</button>
    <button type="button" onClick={onCancel}>Cancel</button>
  </div>;
}
