'use client';

import { useEffect, useRef, useState } from 'react';
import {
  getParagraphs,
  getTextRuns,
  importDocx,
  type DocumentIR,
} from '@local/docx-ir-engine';
import { DocumentRenderer } from '@local/docx-ir-engine/react';

const MAX_FILE_BYTES = 50 * 1024 * 1024;

export function DocxWorkspace() {
  const [document, setDocument] = useState<DocumentIR | null>(null);
  const [filename, setFilename] = useState('Sample document.docx');
  const [busy, setBusy] = useState(true);
  const [error, setError] = useState('');
  const [fontFamily, setFontFamily] = useState('');
  const [fontStatus, setFontStatus] = useState('');
  const fileInput = useRef<HTMLInputElement>(null);
  const request = useRef(0);
  const fonts = useRef<FontFace[]>([]);

  async function open(read: () => Promise<ArrayBuffer>, name: string) {
    const current = ++request.current;
    setBusy(true);
    setError('');
    setDocument(null);
    setFilename(name);
    try {
      const result = await importDocx(await read(), {
        maxUncompressedBytes: MAX_FILE_BYTES,
      });
      if (current === request.current) setDocument(result);
    } catch (cause) {
      if (current === request.current)
        setError(
          cause instanceof Error
            ? cause.message
            : 'The file could not be opened.',
        );
    } finally {
      if (current === request.current) setBusy(false);
    }
  }

  function openSample() {
    return open(async () => {
      const response = await fetch('/sample.docx');
      if (!response.ok)
        throw new Error(
          'The sample could not be loaded. You can still open your own DOCX file.',
        );
      return response.arrayBuffer();
    }, 'Sample document.docx');
  }

  useEffect(() => {
    void openSample();
    return () => {
      request.current++;
      for (const font of fonts.current) window.document.fonts.delete(font);
    };
  }, []);

  async function loadFont(file: File) {
    if (!fontFamily.trim()) {
      setFontStatus('Enter the font family name first.');
      return;
    }
    setFontStatus('Loading font…');
    try {
      const font = await new FontFace(
        fontFamily.trim(),
        await file.arrayBuffer(),
      ).load();
      window.document.fonts.add(font);
      fonts.current.push(font);
      setFontStatus(`“${font.family}” is ready to use in this document.`);
    } catch {
      setFontStatus(
        'Could not load this font. Try a TTF, OTF, WOFF or WOFF2 file.',
      );
    }
  }

  const runs = document ? getTextRuns(document.blocks) : [];
  const paragraphs = document ? getParagraphs(document.blocks) : [];
  const storyRuns = document
    ? [
        ...Object.values(document.headers ?? {}),
        ...Object.values(document.footers ?? {}),
      ].flatMap((blocks) => getTextRuns(blocks ?? []))
    : [];
  const families = [
    ...new Set(
      [...runs, ...storyRuns]
        .flatMap((run) => [
          run.properties.fontFamily,
          ...Object.values(run.properties.fontSlots ?? {}),
          run.properties.complexScript?.fontFamily,
        ])
        .filter((name): name is string => !!name),
    ),
  ];

  return (
    <div className="workspace">
      <header className="app-header">
        <a href="/" className="brand" aria-label="Docs home">
          <span className="document-icon" aria-hidden="true">
            D
          </span>
          <span>
            Docs<span className="brand-note">Your document workspace</span>
          </span>
        </a>
        <div className="header-actions">
          <button
            className="button button-quiet"
            disabled={busy}
            onClick={() => void openSample()}
          >
            Try sample
          </button>
          <button
            className="button button-primary"
            disabled={busy}
            onClick={() => fileInput.current?.click()}
          >
            <span aria-hidden="true">↑</span> Open DOCX
          </button>
          <input
            ref={fileInput}
            hidden
            type="file"
            accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            aria-label="Open a DOCX file"
            onChange={(event) => {
              const file = event.target.files?.[0];
              if (file)
                void open(async () => {
                  if (!/\.docx$/i.test(file.name))
                    throw new Error('Please choose a .docx file.');
                  if (file.size > MAX_FILE_BYTES)
                    throw new Error('Please choose a file smaller than 50 MB.');
                  return file.arrayBuffer();
                }, file.name);
              event.target.value = '';
            }}
          />
        </div>
      </header>

      <main>
        <div className="document-toolbar">
          <div className="document-title">
            <h1>{filename}</h1>
            <span className="preview-badge">Preview</span>
          </div>
          <span className="local-note">
            <span aria-hidden="true">●</span> Files stay on your device
          </span>
        </div>
        <div className="support-note">
          Preview includes tables, lists, images, inherited formatting and
          RTL/LTR text. Word pagination and floating layouts may differ. Editing
          and saving DOCX are not available yet.
        </div>
        {document?.diagnostics?.length ? (
          <details className="compatibility-notes">
            <summary>
              Document compatibility notes ({document.diagnostics.length})
            </summary>
            <ul>
              {document.diagnostics.map((note, index) => (
                <li key={index}>{note.message}</li>
              ))}
            </ul>
          </details>
        ) : null}

        <details className="font-settings">
          <summary>Custom fonts</summary>
          <div className="font-controls">
            <div className="font-help">
              <strong>Use your document’s font</strong>
              <p>
                {families.length
                  ? `Font names in this document: ${families.join(', ')}.`
                  : 'Enter the exact font family name used in your document.'}{' '}
                Load a regular font face for this browser session.
              </p>
            </div>
            <label>
              Font family name
              <input
                value={fontFamily}
                onChange={(event) => setFontFamily(event.target.value)}
                placeholder="e.g. B Nazanin"
              />
            </label>
            <label>
              Font file
              <input
                type="file"
                accept=".ttf,.otf,.woff,.woff2"
                onChange={(event) => {
                  const file = event.target.files?.[0];
                  if (file) void loadFont(file);
                  event.target.value = '';
                }}
              />
            </label>
          </div>
          {fontStatus && (
            <p className="font-status" role="status">
              {fontStatus}
            </p>
          )}
        </details>

        {error && (
          <div className="error" role="alert">
            <strong>Couldn’t open this document</strong>
            <p>{error}</p>
            <button
              className="button button-quiet"
              onClick={() => fileInput.current?.click()}
            >
              Choose another file
            </button>
          </div>
        )}

        <div className="document-canvas" aria-busy={busy}>
          <article
            className={`document-paper${document ? ' has-document' : ''}`}
            dir="ltr"
            aria-label="Document preview"
          >
            {busy ? (
              <div className="empty-state" role="status">
                <span className="loading-dot" />
                Opening your document…
              </div>
            ) : document ? (
              document.blocks.length ? (
                <DocumentRenderer document={document} pageLayout />
              ) : (
                <div className="empty-state">
                  <h2>No supported text found</h2>
                  <p>
                    This file may contain tables, images or other content the
                    engine cannot display yet.
                  </p>
                </div>
              )
            ) : (
              <div className="empty-state">
                <h2>Open a document to get started</h2>
                <p>Choose a DOCX file using the button above.</p>
              </div>
            )}
          </article>
          <footer className="document-footer">
            <span>
              {document
                ? `${paragraphs.length} paragraphs · ${runs.length} text runs`
                : 'DOCX workspace'}
            </span>
            <a href="/sample.docx" download>
              Download sample
            </a>
          </footer>
          {document && (
            <details className="structure">
              <summary>Inspect document structure</summary>
              <pre>{JSON.stringify(document, null, 2)}</pre>
            </details>
          )}
        </div>
      </main>
    </div>
  );
}
