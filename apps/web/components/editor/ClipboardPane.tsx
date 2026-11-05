export function ClipboardPane({
  items,
  disabled,
  onPaste,
  onClear,
  onRead,
  error,
}: {
  items: string[];
  disabled: boolean;
  onPaste: (text: string) => void;
  onClear: () => void;
  onRead: () => void;
  error: string;
}) {
  return (
    <>
      <div className="clipboard-pane-actions">
        <button onClick={onRead}>Read Clipboard</button>
        <button
          disabled={disabled || !items.length}
          onClick={() => onPaste([...items].reverse().join("\n"))}
        >
          Paste All
        </button>
        <button disabled={!items.length} onClick={onClear}>
          Clear All
        </button>
      </div>
      <p className="pane-description">
        Click an item to paste it. Clipboard text is collected while this pane
        is open.
      </p>
      {error && (
        <p className="pane-description" role="status">
          {error}
        </p>
      )}
      <div className="clipboard-pane-items">
        {items.map((item, index) => (
          <button
            key={index}
            disabled={disabled}
            onClick={() => onPaste(item)}
            title={item}
          >
            <span dir="auto">{item}</span>
          </button>
        ))}
      </div>
      {!items.length && <p className="pane-description">Clipboard is empty.</p>}
    </>
  );
}
