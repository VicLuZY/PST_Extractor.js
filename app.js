/**
 * PST Extractor - Browser app for GitHub Pages.
 * Pure static — no build. Libraries loaded from CDN.
 */
import * as PST from 'https://esm.sh/pst-parser@0.0.7';
import JSZip from 'https://esm.sh/jszip@3.10.1';
import { extractPst, buildOutput } from './extraction.js';

const statusEl = document.getElementById('status');
const fileInput = document.getElementById('pst_files');
const form = document.getElementById('extract-form');
const btn = document.getElementById('submit-btn');

function setStatus(msg, isError = false) {
  statusEl.textContent = msg;
  statusEl.className = 'flash ' + (isError ? 'error' : '');
  statusEl.style.display = msg ? 'block' : 'none';
}

function setProgress(msg) {
  statusEl.textContent = msg;
  statusEl.className = 'flash';
  statusEl.style.display = 'block';
}

function formatErrorDetails(err) {
  const name = err?.name || 'Error';
  const message = err?.message || String(err);
  const parts = [`${name}: ${message}`];
  if (err?.cause) parts.push(`Cause: ${err.cause?.message || String(err.cause)}`);
  if (err?.stack) parts.push(`Stack:
${err.stack}`);
  return parts.join('\n');
}

form.addEventListener('submit', async (e) => {
  e.preventDefault();
  const files = fileInput.files;
  if (!files || !files.length) {
    setStatus('Please select one or more PST files.', true);
    return;
  }
  btn.disabled = true;
  setProgress('Extracting… This may take a few minutes for large files.');

  const zip = new JSZip();
  const summary = { pst_files: [], failed_files: [], total_emails: 0, total_attachments: 0, total_teams: 0 };
  const stamp = new Date().toISOString().replace(/[-:]/g, '').slice(0, 15);
  const rootName = `extraction_${stamp}`;

  try {
    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      const name = (file.name || 'pst').replace(/\.pst$/i, '') || 'pst';
      setProgress(`Processing ${i + 1}/${files.length}: ${file.name}…`);
      try {
        const buffer = await file.arrayBuffer();
        const { messages, attachments } = await extractPst(buffer, name, PST);
        const basePath = `${rootName}/${name}`;
        const stats = buildOutput(messages, attachments, zip, basePath);
        summary.pst_files.push({ name, ...stats });
        summary.total_emails += stats.emails;
        summary.total_attachments += stats.attachments;
        summary.total_teams += stats.teams_messages;
      } catch (err) {
        const reason = err && err.message ? err.message : String(err);
        const details = formatErrorDetails(err);
        summary.failed_files.push({ name, reason, details });
        console.warn(`[PST extractor] Failed file ${file.name}:`, err);
      }
    }

    if (!summary.pst_files.length) {
      const details = summary.failed_files
        .map((f, idx) => `${idx + 1}. ${f.name}: ${f.details || f.reason}`)
        .join('\n\n');
      throw new Error(`All files failed to extract.\n\n${details}`);
    }

    zip.file(`${rootName}/summary.json`, JSON.stringify(summary, null, 2));
    const blob = await zip.generateAsync({ type: 'blob' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = `pst_extraction_${summary.pst_files[0]?.name || 'output'}.zip`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(a.href);
    setStatus(summary.failed_files.length
      ? `Done with warnings: extracted ${summary.total_emails} emails, ${summary.total_attachments} attachments. Failed files: ${summary.failed_files.map(x => x.name).join(', ')}`
      : `Done! Extracted ${summary.total_emails} emails, ${summary.total_attachments} attachments.`);
  } catch (err) {
    setStatus('Extraction failed:\n' + formatErrorDetails(err), true);
  } finally {
    btn.disabled = false;
  }
});
