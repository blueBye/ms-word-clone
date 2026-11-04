import { useState } from "react";
import type { SearchDirection, SearchOptions } from "../../lib/search";

export interface SearchProps {
  query: string;
  replacement: string;
  options: SearchOptions;
  onFind: (
    query: string,
    options: SearchOptions,
    direction: SearchDirection,
  ) => Promise<string>;
  onReplace: (
    query: string,
    replacement: string,
    options: SearchOptions,
    all: boolean,
  ) => Promise<string>;
}

export function SearchForm({
  query,
  replacement,
  options,
  replace,
  onFind,
  onReplace,
}: SearchProps & { replace: boolean }) {
  const [value, setValue] = useState(query),
    [withText, setWithText] = useState(replacement);
  const [settings, setSettings] = useState(options),
    [direction, setDirection] = useState<SearchDirection>("down");
  const [more, setMore] = useState(false),
    [busy, setBusy] = useState(false),
    [message, setMessage] = useState("");
  const run = async (all?: boolean) => {
    setBusy(true);
    try {
      setMessage(
        all === undefined
          ? await onFind(value, settings, direction)
          : await onReplace(value, withText, settings, all),
      );
    } catch (error) {
      setMessage(
        error instanceof Error
          ? error.message
          : "Search could not be completed.",
      );
    } finally {
      setBusy(false);
    }
  };
  return (
    <form
      className="advanced-search"
      onSubmit={(event) => {
        event.preventDefault();
        void run();
      }}
    >
      <label className="field">
        Find what
        <input
          value={value}
          onChange={(e) => {
            setValue(e.target.value);
            setMessage("");
          }}
          required
          autoFocus
          dir="auto"
        />
      </label>
      {replace && (
        <label className="field">
          Replace with
          <input
            value={withText}
            onChange={(e) => setWithText(e.target.value)}
            dir="auto"
          />
        </label>
      )}
      <div className="field-row">
        <label className="field">
          Search
          <select
            value={direction}
            onChange={(e) => setDirection(e.target.value as SearchDirection)}
          >
            <option value="down">Down</option>
            <option value="up">Up</option>
          </select>
        </label>
        <button
          type="button"
          className="secondary-button"
          aria-expanded={more}
          onClick={() => setMore(!more)}
        >
          {more ? "Less «" : "More »"}
        </button>
      </div>
      {more && (
        <fieldset className="search-options">
          <legend>Search options</legend>
          {(
            [
              ["matchCase", "Match case"],
              ["wholeWords", "Find whole words only"],
              ["prefix", "Match prefix"],
              ["suffix", "Match suffix"],
              ["ignorePunctuation", "Ignore punctuation characters"],
              ["ignoreWhitespace", "Ignore white-space characters"],
            ] as const
          ).map(([key, label]) => (
            <label className="check-field" key={key}>
              <input
                type="checkbox"
                checked={!!settings[key]}
                onChange={(e) =>
                  setSettings({ ...settings, [key]: e.target.checked })
                }
              />
              {label}
            </label>
          ))}
        </fieldset>
      )}
      <p className="pane-description">
        Searches text across runs, tables, headers and footers. Search wraps at
        the end of the document.
      </p>
      <div className="dialog-actions">
        <button className="primary-button" disabled={busy || !value}>
          Find Next
        </button>
        {replace && (
          <>
            <button
              type="button"
              className="secondary-button"
              disabled={busy || !value}
              onClick={() => void run(false)}
            >
              Replace
            </button>
            <button
              type="button"
              className="secondary-button"
              disabled={busy || !value}
              onClick={() => void run(true)}
            >
              Replace All
            </button>
          </>
        )}
      </div>
      <p className="pane-description" role="status" aria-live="polite">
        {message}
      </p>
    </form>
  );
}
