import { useEffect, useRef, useState } from 'react';
import { totalFunctions } from '../../application/table-commands';

interface TotalRowDropdownProps {
  readonly currentCode: number;
  readonly position: { readonly x: number; readonly y: number; readonly size: number };
  readonly onSelect: (code: number) => void;
}

export function TotalRowDropdown({ currentCode, position, onSelect }: TotalRowDropdownProps) {
  const [open, setOpen] = useState(false);
  const containerRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    if (!open) return;
    function handleClickOutside(event: MouseEvent) {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setOpen(false);
      }
    }
    document.addEventListener('pointerdown', handleClickOutside);
    return () => document.removeEventListener('pointerdown', handleClickOutside);
  }, [open]);

  return (
    <div
      ref={containerRef}
      className="total-button-container"
      style={{
        position: 'absolute',
        left: position.x,
        top: position.y,
        width: position.size,
        height: position.size,
        zIndex: 3,
      }}
    >
      <button
        type="button"
        className="total-button"
        aria-label="Table total function options"
        title="Table total function options"
        onClick={() => setOpen(prev => !prev)}
        style={{ width: '100%', height: '100%' }}
      >
        <svg width="10" height="6" viewBox="0 0 10 6" fill="none" stroke="currentColor" strokeWidth="1.5">
          <path d="M1 1L5 5L9 1" strokeLinecap="round" strokeLinejoin="round" />
        </svg>
      </button>

      {open && (
        <div className="total-dropdown-menu" role="menu">
          {totalFunctions.map(([label, code]) => {
            const isChecked = currentCode === code;
            return (
              <button
                key={code}
                type="button"
                role="menuitemcheckbox"
                aria-checked={isChecked}
                className={'total-dropdown-item' + (isChecked ? ' checked' : '')}
                onClick={() => {
                  onSelect(code);
                  setOpen(false);
                }}
              >
                <span className="total-dropdown-check">{isChecked ? '✓' : ''}</span>
                <span className="total-dropdown-label">{label}</span>
              </button>
            );
          })}
        </div>
      )}
    </div>
  );
}
