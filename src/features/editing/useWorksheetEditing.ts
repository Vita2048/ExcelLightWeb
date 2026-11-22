import { useRef, useState } from 'react';
import { autoSum, captureClipboard, clear, clipboardText, editorEntryText, expandMerges, fill, parseEntry, pasteBlock, pasteText,
  type Aggregate, type ClearMode, type ClipboardBlock, type EditRange, type FillDirection, type PasteMode,
} from '../../application/editing-commands';
import type { WorkbookSession } from '../../application/workbook-session';
import type { WorkbookSnapshot } from '../../application/workbook-protocol';
import { equalDto } from '../../core/workbook-delta';
import { documentLimits, findReferencedCell, type CellDocument, type CellReference, type SheetDocument } from '../../core/workbook';
import { formulaAddress, quoteSheet, type ReferenceBox } from '../../core/formula-references';
import { requireDocument } from '../../core/workbook-validation';
import { presentFormula } from '../formulas/formula-presentation';
import { referenceSlot, type TextSelection } from '../formulas/reference-pick';

export type EditSurface = 'inline' | 'formula';
export type AfterEditKey = 'Enter' | 'Tab' | 'ArrowUp' | 'ArrowDown' | 'ArrowLeft' | 'ArrowRight';
export type EditAction =
  | { kind: 'copy' | 'cut' }
  | { kind: 'paste'; mode: PasteMode }
  | { kind: 'clear'; mode: ClearMode }
  | { kind: 'fill'; direction: FillDirection }
  | { kind: 'autosum'; aggregate: Aggregate };

interface Draft {
  readonly surface: EditSurface;
  readonly sessionId: string;
  readonly target: CellReference;
  readonly original: CellDocument | undefined;
  readonly text: string;
  readonly direct: boolean;
  readonly functionDraft: boolean;
  readonly pick?: TextSelection & { readonly sheetId: string; readonly area: ReferenceBox };
  readonly selectText?: TextSelection;
}

interface LocalClipboard {
  readonly token: string;
  readonly block: ClipboardBlock;
  readonly text: string;
}
interface ManualPaste {
  readonly basis: WorkbookSnapshot;
  readonly range: EditRange;
  readonly mode: PasteMode;
}

const clipboardType = 'application/x-excellight-cells';

export function useWorksheetEditing(
  session: WorkbookSession, changed: (compound: boolean) => void, separator: ',' | ';', locale: string,
  makeId: () => string, focusGrid: () => void, focusEditor: (surface: EditSurface) => void,
) {
  const liveDraft = useRef<Draft | null>(null);
  const composing = useRef(false);
  const [draft, renderDraft] = useState<Draft | null>(null);
  const [status, setStatus] = useState('');
  const localClipboard = useRef<LocalClipboard | null>(null);
  const [manualPaste, setManualPaste] = useState<ManualPaste | null>(null);
  const intent = useRef(0);

  function showDraft(next: Draft | null) {
    liveDraft.current = next;
    renderDraft(next);
  }

  function editorText(cell: CellDocument | undefined) {
    const source = editorEntryText(cell, locale);
    return cell?.content.kind === 'formula' ? presentFormula(source, ',', separator) : source;
  }

  function begin(target: CellReference, surface: EditSurface, replacement?: string, functionDraft = false) {
    const existing = liveDraft.current;
    if (existing) {
      if (surface !== existing.surface) {
        const { selectText: _selection, ...rest } = existing;
        showDraft({ ...rest, surface });
      }
      return;
    }
    intent.current++;
    setStatus('');
    const original = findReferencedCell(session.active, target);
    showDraft({ surface, sessionId: session.sessionId, target, original, text: replacement ?? editorText(original),
      direct: surface === 'inline' && replacement !== undefined && !functionDraft, functionDraft });
  }

  function update(text: string) {
    const current = liveDraft.current;
    if (current && current.text !== text) {
      const { pick: _pick, selectText: _selection, ...rest } = current;
      showDraft({ ...rest, text });
    }
  }

  function pickSlot(selection: TextSelection) {
    const current = liveDraft.current;
    if (!current || composing.current || current.original?.style?.numberFormat === '@') return;
    const last = current.pick;
    return referenceSlot(current.text, last && selection.start === last.end && selection.end === last.end ? last : selection);
  }

  function beginReferencePick(selection: TextSelection, sheet: SheetDocument) {
    const slot = pickSlot(selection);
    const original = liveDraft.current;
    if (!slot || !original) return;
    let latest = original;
    return {
      update(area: ReferenceBox) {
        if (liveDraft.current !== latest) return;
        const range = expandMerges({ sheetId: sheet.id, row: area.top, column: area.left,
          rows: area.bottom - area.top + 1, columns: area.right - area.left + 1 }, sheet);
        const bottom = range.row + range.rows - 1, right = range.column + range.columns - 1;
        const first = formulaAddress(range);
        const last = range.rows === 1 && range.columns === 1 ? '' : ':' + formulaAddress({ row: bottom, column: right });
        const reference = (sheet.id === original.target.sheetId ? '' : quoteSheet(sheet.name)) + first + last;
        const text = original.text.slice(0, slot.start) + reference + original.text.slice(slot.end);
        if (text.length > documentLimits.formulaLength || latest.text === text && latest !== original) return;
        const end = slot.start + reference.length;
        latest = { ...original, text, pick: { start: slot.start, end, sheetId: sheet.id,
          area: { top: range.row, left: range.column, bottom, right } }, selectText: { start: end, end } };
        showDraft(latest);
      },
      cancel() {
        if (liveDraft.current === latest) showDraft({ ...original, selectText: { ...selection } });
      },
    };
  }

  function showCrossSheetDraft(selection: TextSelection) {
    const current = liveDraft.current;
    if (current?.surface === 'inline') showDraft({ ...current, surface: 'formula', direct: false, selectText: selection });
  }

  function reject(error: unknown) {
    setStatus(error instanceof Error ? error.message : 'Worksheet command failed.');
  }

  function commit(): boolean {
    const current = liveDraft.current;
    if (!current) return true;
    if (composing.current) return false;
    try {
      if (current.text !== editorText(current.original)) {
        const basis = session.capture();
        const actual = findReferencedCell(basis.document, current.target);
        requireDocument(current.sessionId === basis.sessionId && actual?.id === current.original?.id &&
          equalDto(actual?.content, current.original?.content) && equalDto(actual?.style, current.original?.style),
        'The cell changed while you were editing it.');
        const outcome = session.execute({ kind: 'set-cells', label: 'Edit cell', edits: [{
          target: current.target, content: parseEntry(actual?.style?.numberFormat === '@' ? current.text :
            presentFormula(current.text, separator, ','), actual?.style, locale, { makeId, existing: actual?.content }),
        }] });
        if (outcome.changed) changed(false);
      }
      showDraft(null);
      setStatus('');
      return true;
    } catch (error) {
      reject(error);
      queueMicrotask(() => focusEditor(current.surface));
      return false;
    }
  }

  function cancel() {
    composing.current = false;
    showDraft(null);
    setStatus('');
    focusGrid();
  }

  function moveHistory(undo: boolean) {
    if (!commit()) return;
    intent.current++;
    if (undo ? session.undo() : session.redo()) changed(true);
    setStatus('');
    focusGrid();
  }

  function onlyRange(ranges: readonly EditRange[]): EditRange {
    requireDocument(ranges.length === 1, 'This command requires one contiguous selection.');
    return ranges[0]!;
  }

  function paste(basis: WorkbookSnapshot, range: EditRange, mode: PasteMode, rich: LocalClipboard | null, text?: string) {
    const operation = rich ? pasteBlock(basis, range, rich.block, mode, makeId) :
      pasteText(basis, range, text ?? '', mode, locale, makeId);
    const outcome = session.execute(operation);
    if (rich?.block.cut && localClipboard.current === rich) localClipboard.current = null;
    if (outcome.changed) changed(outcome.delta.kind !== 'cells' || outcome.delta.changes.length > 1);
  }

  function closeManualPaste() {
    setManualPaste(null);
    setStatus('');
    requestAnimationFrame(focusGrid);
  }

  function submitManualPaste(text: string) {
    if (!manualPaste) return;
    try {
      requireDocument(session.isCurrentAuthored(manualPaste.basis), 'The worksheet changed before the paste completed.');
      paste(session.capture(), manualPaste.range, manualPaste.mode, null, text);
      closeManualPaste();
    } catch (error) { reject(error); }
  }

  async function run(action: EditAction, ranges: readonly EditRange[], event?: DataTransfer): Promise<CellReference | undefined> {
    if (!commit()) return;
    const serial = ++intent.current;
    const basis = session.capture();
    setStatus('');
    try {
      if (action.kind === 'cut' || action.kind === 'copy') {
        const block = captureClipboard(basis, onlyRange(ranges), action.kind === 'cut');
        const text = clipboardText(block);
        const token = crypto.randomUUID();
        localClipboard.current = { text, block, token };
        if (block.cut) setStatus('Cut selection. Paste to move it.');
        if (event) {
          event.setData('text/plain', text);
          try { event.setData(clipboardType, token); } catch { /* Plain text is already available. */ }
        } else {
          focusGrid();
          try { await navigator.clipboard.writeText(text); }
          catch {
            if (serial === intent.current) setStatus(block.cut ? 'Cut selection. Use the ribbon to paste it.' :
              'Copied to this workbook. System clipboard access is unavailable.');
          }
        }
      } else if (action.kind === 'paste') {
        const range = onlyRange(ranges);
        const rich = localClipboard.current;
        if (event) {
          const text = event.getData('text/plain');
          const token = event.getData(clipboardType);
          // Browsers may strip custom formats. An exact plain-text match preserves the in-page block.
          if (rich && (token === rich.token || !rich.block.cut && !token &&
              text.replaceAll('\r\n', '\n') === rich.text.replaceAll('\r\n', '\n'))) {
            paste(basis, range, action.mode, rich);
          } else if (text || Array.from(event.types).includes('text/plain')) {
            paste(basis, range, action.mode, null, text);
          } else requireDocument(false, 'The clipboard contains no worksheet text.');
        } else {
          let text: string | undefined;
          let denied = false;
          // Base the completion on the original selection and revision.
          try {
            if (rich) {
              // Firefox and WebKit may show a browser paste prompt for programmatic reads.
              const permission = await navigator.permissions.query({ name: 'clipboard-read' as PermissionName });
              if (permission.state !== 'granted') throw new Error('System clipboard permission is unavailable.');
            }
            text = await navigator.clipboard.readText();
          }
          catch { denied = true; }
          if (serial !== intent.current || !session.isCurrentAuthored(basis)) return;
          if (rich && (denied || text?.replaceAll('\r\n', '\n') === rich.text.replaceAll('\r\n', '\n'))) {
            paste(session.capture(), range, action.mode, rich);
          } else {
            if (denied || text === undefined) setManualPaste({ basis, range, mode: action.mode });
            else {
              requireDocument(text !== '', 'The system clipboard contains no worksheet text.');
              paste(session.capture(), range, action.mode, null, text);
            }
          }
        }
      } else if (action.kind === 'clear') {
        if (session.execute(clear(basis, ranges, action.mode)).changed) changed(true);
      } else if (action.kind === 'fill') {
        if (session.execute(fill(basis, onlyRange(ranges), action.direction, makeId)).changed) changed(true);
      } else if (action.kind === 'autosum') {
        const result = autoSum(basis, onlyRange(ranges), action.aggregate);
        if (session.execute(result.command).changed) changed(false);
        return result.target;
      }
    } catch (error) { reject(error); }
    finally { if (serial === intent.current) focusGrid(); }
  }

  return {
    draft, status, begin, update, commit, cancel, moveHistory, run, manualPaste, closeManualPaste, submitManualPaste,
    pickSlot, beginReferencePick, showCrossSheetDraft,
    blur: (surface: EditSurface) => { if (liveDraft.current?.surface === surface) commit(); },
    editorText, reject, report: (message: string) => setStatus(message), clearStatus: () => setStatus(''),
    composing: () => composing.current,
    compose: (active: boolean) => { composing.current = active; },
    keepCaret: () => { if (liveDraft.current?.direct) showDraft({ ...liveDraft.current, direct: false }); },
    selectionChanged: () => { intent.current++; },
    clearCut: () => { if (localClipboard.current?.block.cut) localClipboard.current = null; },
  };
}
