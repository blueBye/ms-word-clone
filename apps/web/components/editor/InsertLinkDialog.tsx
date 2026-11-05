import type { EditCommand, Paragraph } from "@local/docx-ir-engine/core";
import { paragraphText } from "../../lib/editor";
import { Dialog } from "./Controls";
export function InsertLinkDialog({
  selected,
  range,
  onClose,
  execute,
}: {
  selected: Paragraph;
  range: { start: number; end: number };
  onClose: () => void;
  execute: (command: EditCommand) => Promise<boolean>;
}) {
  return (
    <Dialog title="Insert Hyperlink" onClose={onClose}>
      <form
        onSubmit={async (event) => {
          event.preventDefault();
          const data = new FormData(event.currentTarget);
          if (
            await execute({
              type: "link",
              id: selected.id!,
              ...range,
              text: String(data.get("text")),
              href: String(data.get("href")).trim(),
            })
          )
            onClose();
        }}
      >
        <label className="field">
          Text to display
          <input
            name="text"
            defaultValue={paragraphText(selected).slice(range.start, range.end)}
            required
            autoFocus
          />
        </label>
        <label className="field">
          Address
          <input name="href" placeholder="https://example.com" required />
        </label>
        <div className="dialog-actions">
          <button type="button" className="secondary-button" onClick={onClose}>
            Cancel
          </button>
          <button className="primary-button">OK</button>
        </div>
      </form>
    </Dialog>
  );
}
