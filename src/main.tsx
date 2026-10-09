/// <reference types="vite/client" />
/**
 * Entry point. `?mock=1` (DEV builds only) swaps the module facades for realistic fake data (src/state/mock.ts,
 * loaded with a dynamic import that production builds drop entirely).
 */
import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App, { type DemoOptions } from './App';
import { ErrorBoundary } from './components/ErrorBoundary';
import { realServices, wantsMock, type Services } from './state/services';
import './styles.css';

// The mock is available in dev builds, and in production builds only when built with VITE_ALLOW_MOCK=1 (the UI
// performance check, tests/ui/perf.mjs); otherwise the condition is statically false and the mock is dropped.
const MOCK_ALLOWED = import.meta.env.DEV || import.meta.env.VITE_ALLOW_MOCK === '1';

async function pickServices(): Promise<{ services: Services; demo: DemoOptions | null }> {
  if (MOCK_ALLOWED && wantsMock(window.location.search)) {
    const mock = await import('./state/mock');
    const stage = mock.demoStage(window.location.search);
    return { services: mock.createMockServices(window.location.search), demo: stage ? { files: () => mock.demoFiles(window.location.search), stage } : null };
  }
  return { services: realServices, demo: null };
}

function showFatal(root: HTMLElement, message: string) {
  root.innerHTML = '';
  const box = document.createElement('div');
  box.className = 'error-fallback';
  const h = document.createElement('h2');
  h.textContent = 'アプリを起動できませんでした';
  const p = document.createElement('p');
  p.textContent = message;
  box.append(h, p);
  root.append(box);
}

(window as unknown as { __excelMergerStarted?: boolean }).__excelMergerStarted = true;

const rootElement = document.getElementById('root');
if (rootElement) {
  pickServices()
    .then(({ services, demo }) => {
      createRoot(rootElement).render(
        <StrictMode>
          <ErrorBoundary area="アプリ">
            <App services={services} demo={demo} />
          </ErrorBoundary>
        </StrictMode>,
      );
    })
    .catch((e: unknown) => showFatal(rootElement, e instanceof Error ? e.message : String(e)));
}
