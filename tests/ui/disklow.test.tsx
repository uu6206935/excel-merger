/**
 * Low-disk notice (final round): GET /api/status diskLow / diskFreeMB / diskMessage / outputDiskLow ->
 * helpers (diskNotice, sameServerStatus ...), the reducer's server/refresh, the controller's refresh when the export
 * dialog opens, and the markup of the header chip and of the export dialog (server-side render, no DOM needed).
 * The dynamic behaviour in a browser (notice appears / disappears, default engine follows) is tests/ui/disklow-check.mjs.
 */
import { renderToStaticMarkup } from 'react-dom/server';
import { describe, expect, it } from 'vitest';
import type { ServerStatus } from '../../src/api';
import { AppHeader } from '../../src/components/AppHeader';
import { ExportDialog, type ExportDialogProps } from '../../src/components/ExportDialog';
import { createController } from '../../src/state/controller';
import { DISK_LOW_CHIP_TEXT, diskNotice, driveLabel, formatFreeMB, sameServerStatus } from '../../src/state/helpers';
import { createMockServices } from '../../src/state/mock';
import { initialState, reducer } from '../../src/state/reducer';
import type { Services } from '../../src/state/services';
import type { ServerState } from '../../src/state/types';
import { file, store } from './fixtures';

const SERVER_MESSAGE =
  'C: ドライブの空き容量が少なくなっています（残り 1.2 GB）。Excel本体での出力は一時ファイルを作るため、空きが足りないと失敗することがあります。不要なファイルを削除するか、「内蔵エンジン」で出力してください。';

const okStatus: ServerStatus = { app: 'excel-merger', version: '1.0.0', excel: { installed: true, version: '16.0' }, outputDir: 'D:\\出力', diskLow: false, diskFreeMB: 20000, outputDiskLow: false };
const lowStatus: ServerStatus = { ...okStatus, diskLow: true, diskFreeMB: 1200, diskMessage: SERVER_MESSAGE };
const outLowStatus: ServerStatus = { ...okStatus, outputDiskLow: true };
const server = (status: ServerStatus | null, patch: Partial<ServerState> = {}): ServerState => ({ checked: true, status, connected: true, ...patch });
const noop = () => undefined;

describe('diskNotice', () => {
  it('is null without a server, with a server that stopped answering, or when both drives are fine', () => {
    expect(diskNotice(server(null))).toBeNull();
    expect(diskNotice({ status: null, connected: true })).toBeNull();
    expect(diskNotice(server(lowStatus, { connected: false }))).toBeNull();
    expect(diskNotice(server(okStatus))).toBeNull();
    // a server that does not know the fields at all (older build)
    const { diskLow: _a, diskFreeMB: _b, outputDiskLow: _c, ...old } = okStatus;
    void _a;
    void _b;
    void _c;
    expect(diskNotice(server(old))).toBeNull();
  });

  it('diskLow: the server\'s own message, plus the advice to use the built-in engine', () => {
    const n = diskNotice(server(lowStatus))!;
    expect(n.tempLow).toBe(true);
    expect(n.outputLow).toBe(false);
    expect(n.tempMessage).toBe(SERVER_MESSAGE);
    expect(n.summary).toBe(SERVER_MESSAGE);
    expect(n.recommendation).toContain('Excel本体エンジンは');
    expect(n.recommendation).toContain('使えないことがあります');
    expect(n.recommendation).toContain('内蔵エンジン');
    expect(n.recommendation).toContain('おすすめ');
    expect(n.outputMessage).toBeNull();
  });

  it('diskLow without a server message (older server): the same kind of sentence is made here, with the free space', () => {
    const n = diskNotice(server({ ...lowStatus, diskMessage: undefined, diskFreeMB: 800 }))!;
    expect(n.tempMessage).toContain('空き容量が少なくなっています（残り 800 MB）');
    expect(n.tempMessage).toContain('内蔵エンジン');
    expect(n.summary).toBe(n.tempMessage);
    const blank = diskNotice(server({ ...lowStatus, diskMessage: '   ', diskFreeMB: undefined }))!;
    expect(blank.tempMessage).toContain('空き容量が少なくなっています。');
    expect(blank.tempMessage).not.toContain('残り');
  });

  it('outputDiskLow only: says the output folder drive is nearly full, no advice about the engine', () => {
    const n = diskNotice(server(outLowStatus))!;
    expect(n.tempLow).toBe(false);
    expect(n.outputLow).toBe(true);
    expect(n.tempMessage).toBeNull();
    expect(n.recommendation).toBeNull();
    expect(n.outputMessage).toContain('出力フォルダのドライブ（D: ドライブ）');
    expect(n.outputMessage).toContain('空き容量が少なくなっています');
    expect(n.summary).toBe(n.outputMessage);
    // an output path without a drive letter (UNC): no drive in the sentence
    const unc = diskNotice(server({ ...outLowStatus, outputDir: '\\\\nas\\share\\out' }))!;
    expect(unc.outputMessage).toMatch(/^出力フォルダのドライブの空き容量が少なくなっています/);
  });

  it('both: the chip text carries both sentences', () => {
    const n = diskNotice(server({ ...lowStatus, outputDiskLow: true }))!;
    expect(n.tempLow && n.outputLow).toBe(true);
    expect(n.summary).toContain(SERVER_MESSAGE);
    expect(n.summary).toContain('出力フォルダのドライブ');
    expect(n.recommendation).not.toBeNull();
    expect(n.outputMessage).not.toBeNull();
  });

  it('formatFreeMB / driveLabel', () => {
    expect(formatFreeMB(800)).toBe('800 MB');
    expect(formatFreeMB(1024)).toBe('1.0 GB');
    expect(formatFreeMB(6144)).toBe('6.0 GB');
    expect(formatFreeMB(-1)).toBe('');
    expect(formatFreeMB(Number.NaN)).toBe('');
    expect(driveLabel('c:\\Users\\x')).toBe('C: ドライブ');
    expect(driveLabel('D:/出力')).toBe('D: ドライブ');
    expect(driveLabel('\\\\nas\\share')).toBeNull();
    expect(driveLabel(undefined)).toBeNull();
  });

  it('sameServerStatus compares every field the UI reads', () => {
    expect(sameServerStatus(null, null)).toBe(true);
    expect(sameServerStatus(okStatus, okStatus)).toBe(true);
    expect(sameServerStatus(okStatus, { ...okStatus })).toBe(true);
    expect(sameServerStatus(okStatus, null)).toBe(false);
    expect(sameServerStatus(null, okStatus)).toBe(false);
    expect(sameServerStatus(okStatus, lowStatus)).toBe(false);
    expect(sameServerStatus(okStatus, { ...okStatus, outputDiskLow: true })).toBe(false);
    expect(sameServerStatus(okStatus, { ...okStatus, diskFreeMB: 19999 })).toBe(false);
    expect(sameServerStatus(okStatus, { ...okStatus, excel: { installed: false } })).toBe(false);
    expect(sameServerStatus(okStatus, { ...okStatus, outputDir: 'E:\\x' })).toBe(false);
  });
});

describe('reducer: server/refresh', () => {
  const withServer = (status: ServerStatus | null) => reducer(initialState, { type: 'server/status', status });

  it('replaces the status of a known server and keeps checked / connected', () => {
    const a = withServer(okStatus);
    const b = reducer(a, { type: 'server/refresh', status: lowStatus });
    expect(b.server).toEqual({ checked: true, status: lowStatus, connected: true });
    const lost = reducer(b, { type: 'server/connection', connected: false });
    const c = reducer(lost, { type: 'server/refresh', status: okStatus });
    expect(c.server.status).toEqual(okStatus);
    expect(c.server.connected).toBe(false);
  });

  it('an unchanged answer keeps the very same state object (no re-render)', () => {
    const a = withServer(lowStatus);
    expect(reducer(a, { type: 'server/refresh', status: { ...lowStatus } })).toBe(a);
  });

  it('standalone stays standalone (a server that shows up later has no heartbeat)', () => {
    const standalone = withServer(null);
    expect(reducer(standalone, { type: 'server/refresh', status: lowStatus })).toBe(standalone);
    expect(reducer(initialState, { type: 'server/refresh', status: lowStatus })).toBe(initialState);
  });
});

describe('controller: the export dialog refreshes the server status', () => {
  function setup(statuses: Array<ServerStatus | null | 'throw'>) {
    const st = store();
    let calls = 0;
    const base = createMockServices('?mock=1&delay=0.05');
    const services: Services = {
      ...base,
      async getServerStatus() {
        const i = calls++;
        // the first call is init(): a fine server; later calls follow the script
        if (i === 0) return okStatus;
        const next = statuses[Math.min(i - 1, statuses.length - 1)];
        if (next === 'throw') throw new Error('boom');
        return next;
      },
    };
    const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
    return { st, ctrl, calls: () => calls };
  }

  async function merged(env: ReturnType<typeof setup>) {
    env.ctrl.init();
    await env.ctrl.addFiles([file('学生回答.xlsx'), file('職場回答.xlsx'), file('人事回答.xlsx')]);
    await env.ctrl.idle();
    await env.ctrl.ensureAnalysis();
    await env.ctrl.merge();
    expect(env.st.state.phase).toBe('merged');
  }

  it('opening the dialog asks /api/status again: the notice data appears and disappears without a reload', async () => {
    const env = setup([lowStatus, okStatus, outLowStatus]);
    await merged(env);
    expect(env.st.state.server.status).toEqual(okStatus);
    expect(diskNotice(env.st.state.server)).toBeNull();

    env.ctrl.openExport();
    expect(env.st.state.modal?.kind).toBe('export');
    await env.ctrl.idle();
    expect(env.calls()).toBe(2);
    expect(env.st.state.server.status?.diskLow).toBe(true);
    expect(diskNotice(env.st.state.server)?.tempMessage).toBe(SERVER_MESSAGE);

    env.st.dispatch({ type: 'modal/close' });
    env.ctrl.openExport();
    await env.ctrl.idle();
    expect(env.st.state.server.status?.diskLow).toBe(false);
    expect(diskNotice(env.st.state.server)).toBeNull();

    env.st.dispatch({ type: 'modal/close' });
    env.ctrl.openExport();
    await env.ctrl.idle();
    expect(diskNotice(env.st.state.server)).toMatchObject({ tempLow: false, outputLow: true });
    expect(env.calls()).toBe(4);
  });

  it('a server that does not answer (null / error) keeps the last known status', async () => {
    const env = setup([null, 'throw']);
    await merged(env);
    const toastBefore = env.st.state.toast; // the merge's own toast
    env.ctrl.openExport();
    await env.ctrl.idle();
    expect(env.st.state.server.status).toEqual(okStatus);
    env.st.dispatch({ type: 'modal/close' });
    env.ctrl.openExport();
    await env.ctrl.idle();
    expect(env.st.state.server.status).toEqual(okStatus);
    expect(env.st.state.server.checked).toBe(true);
    expect(env.st.state.toast).toBe(toastBefore); // no error toast about the failed refresh
  });

  it('standalone: no request at all', async () => {
    const st = store();
    let calls = 0;
    const base = createMockServices('?mock=standalone&delay=0.05');
    const services: Services = {
      ...base,
      async getServerStatus() {
        calls++;
        return null;
      },
    };
    const ctrl = createController({ services, dispatch: st.dispatch, getState: st.getState });
    ctrl.init();
    await ctrl.idle();
    expect(calls).toBe(1);
    await ctrl.refreshServerStatus();
    expect(calls).toBe(1);
    expect(st.state.server.status).toBeNull();
  });

  it('a dialog that is still running an export is reopened without asking again', async () => {
    const env = setup([lowStatus]);
    await merged(env);
    const p = env.ctrl.runExport({ fileName: 'a', engine: 'js' });
    expect(env.st.state.exportState.phase).toBe('running');
    const before = env.calls();
    env.ctrl.openExport();
    expect(env.calls()).toBe(before);
    await p;
  });
});

// ------------------------------------------------------------------------------------------ markup

function dialog(srv: ServerState, patch: Partial<ExportDialogProps> = {}) {
  return renderToStaticMarkup(
    <ExportDialog
      defaultName="学生回答_統合_20261002.xlsx"
      format="xlsx"
      server={srv}
      exportState={{ phase: 'idle', engine: null, message: '', ratio: null }}
      editCount={2}
      onExport={noop}
      onCancelExport={noop}
      onOpenFile={noop}
      onOpenFolder={noop}
      onDownloadAgain={noop}
      onClose={noop}
      {...patch}
    />,
  );
}

/** the <input ...> tag of a radio, by its data-testid */
const radio = (html: string, testId: string) => new RegExp(`<input[^>]*data-testid="${testId}"[^>]*>`).exec(html)?.[0] ?? '';
const isChecked = (tag: string) => /\schecked(=|\s|>)/.test(tag);
const isDisabled = (tag: string) => /\sdisabled(=|\s|>)/.test(tag);

describe('header chip', () => {
  const header = (srv: ServerState) => renderToStaticMarkup(<AppHeader server={srv} step={0} />);

  it('amber chip 「空き容量が少ない」 with the server message as tooltip and aria-label, only when a drive is low', () => {
    const html = header(server(lowStatus));
    expect(html).toContain('data-testid="disk-low-chip"');
    expect(html).toContain(`>${DISK_LOW_CHIP_TEXT}<`);
    expect(DISK_LOW_CHIP_TEXT).toBe('空き容量が少ない');
    expect(html).toContain('tag-warn');
    expect(html).toContain(`title="${SERVER_MESSAGE}"`);
    expect(html).toContain(`aria-label="${SERVER_MESSAGE}"`);
    // right next to the Excel本体 tag, inside the status line
    expect(html.indexOf('Excel本体: 利用可能')).toBeLessThan(html.indexOf('disk-low-chip'));
    expect(html.indexOf('disk-low-chip')).toBeLessThan(html.indexOf('data-testid="step-indicator"'));
  });

  it('output drive low alone also shows it (with the output-drive sentence)', () => {
    const html = header(server(outLowStatus));
    expect(html).toContain('data-testid="disk-low-chip"');
    expect(html).toContain('aria-label="出力フォルダのドライブ（D: ドライブ）の空き容量が少なくなっています。');
  });

  it('nothing when the drives are fine, in standalone mode, or when the server stopped answering', () => {
    expect(header(server(okStatus))).not.toContain('disk-low-chip');
    expect(header(server(null))).not.toContain('disk-low-chip');
    expect(header(server(lowStatus, { connected: false }))).not.toContain('disk-low-chip');
    expect(header({ checked: false, status: null, connected: true })).not.toContain('disk-low-chip');
    // the existing parts are untouched
    expect(header(server(okStatus))).toContain('Excel本体: 利用可能');
  });
});

describe('export dialog notice', () => {
  it('no notice and Excel本体 as the default when the disk is fine', () => {
    const html = dialog(server(okStatus));
    expect(html).not.toContain('export-disk-notice');
    expect(isChecked(radio(html, 'export-engine-excel'))).toBe(true);
    expect(isChecked(radio(html, 'export-engine-js'))).toBe(false);
  });

  it('diskLow: amber notice with the server message above the engine choice, 内蔵エンジン is the default, Excel本体 stays selectable', () => {
    const html = dialog(server(lowStatus));
    expect(html).toContain('data-testid="export-disk-notice"');
    expect(html).toContain('data-temp-low="true"');
    expect(html).toContain('data-output-low="false"');
    expect(html).toContain('role="status"');
    expect(html).toContain(SERVER_MESSAGE);
    expect(html).toContain('Excel本体エンジンは空き容量が足りないと使えないことがあります。内蔵エンジンでの出力をおすすめします。');
    expect(html).not.toContain('export-disk-output');
    // above the engine choice (and below the file name)
    expect(html.indexOf('data-testid="export-filename"')).toBeLessThan(html.indexOf('export-disk-notice'));
    expect(html.indexOf('export-disk-notice')).toBeLessThan(html.indexOf('出力方法'));
    expect(html.indexOf('export-disk-notice')).toBeLessThan(html.indexOf('data-testid="export-engine-excel"'));
    // default engine
    const excel = radio(html, 'export-engine-excel');
    const js = radio(html, 'export-engine-js');
    expect(isChecked(js)).toBe(true);
    expect(isChecked(excel)).toBe(false);
    expect(isDisabled(excel)).toBe(false);
    expect(isDisabled(js)).toBe(false);
  });

  it('outputDiskLow only: says the output folder drive is nearly full; the default engine stays Excel本体', () => {
    const html = dialog(server(outLowStatus));
    expect(html).toContain('data-testid="export-disk-notice"');
    expect(html).toContain('data-temp-low="false"');
    expect(html).toContain('data-output-low="true"');
    expect(html).toContain('出力フォルダのドライブ（D: ドライブ）の空き容量が少なくなっています');
    expect(html).not.toContain('export-disk-temp');
    expect(html).not.toContain('export-disk-advice');
    expect(isChecked(radio(html, 'export-engine-excel'))).toBe(true);
  });

  it('both low: both sentences, built-in engine default', () => {
    const html = dialog(server({ ...lowStatus, outputDiskLow: true }));
    expect(html).toContain('export-disk-temp');
    expect(html).toContain('export-disk-advice');
    expect(html).toContain('export-disk-output');
    expect(isChecked(radio(html, 'export-engine-js'))).toBe(true);
  });

  it('no Excel installed: built-in engine as before, and the notice still shows', () => {
    const html = dialog(server({ ...lowStatus, excel: { installed: false } }));
    expect(html).toContain('export-disk-notice');
    expect(isChecked(radio(html, 'export-engine-js'))).toBe(true);
    expect(isDisabled(radio(html, 'export-engine-excel'))).toBe(true);
  });

  it('standalone / server stopped answering: no notice (the numbers are stale or unknown)', () => {
    expect(dialog(server(null))).not.toContain('export-disk-notice');
    expect(dialog(server(lowStatus, { connected: false }))).not.toContain('export-disk-notice');
  });
});
