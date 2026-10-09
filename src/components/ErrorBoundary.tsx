/**
 * Never a blank screen: a render error shows a Japanese message with ways out instead of unmounting the app.
 */
import { Component, type ErrorInfo, type ReactNode } from 'react';

interface Props {
  children: ReactNode;
  /** short name of the area for the message ('画面' / '表') */
  area?: string;
  /** optional recovery action shown next to 「もう一度表示」 (e.g. 結合をやり直す) */
  onReset?: () => void;
  resetLabel?: string;
  /** changes of this key reset the boundary (e.g. the doc version) */
  resetKey?: unknown;
}

interface State {
  error: Error | null;
  key: unknown;
}

export class ErrorBoundary extends Component<Props, State> {
  state: State = { error: null, key: undefined };

  static getDerivedStateFromError(error: Error): Partial<State> {
    return { error };
  }

  static getDerivedStateFromProps(props: Props, state: State): Partial<State> | null {
    if (props.resetKey !== state.key) return { key: props.resetKey, error: null };
    return null;
  }

  componentDidCatch(error: Error, info: ErrorInfo): void {
    console.error('[ErrorBoundary]', error, info.componentStack);
  }

  render() {
    const { error } = this.state;
    if (!error) return this.props.children;
    const area = this.props.area ?? '画面';
    return (
      <div className="error-fallback" role="alert" data-testid="error-fallback">
        <h2>{area}の表示中に問題が発生しました</h2>
        <p>データは失われていません。「もう一度表示」を押すか、元に戻す（Ctrl+Z）を試してください。</p>
        <details className="error-tech">
          <summary>技術的な詳細（問い合わせ用）</summary>
          <p className="error-detail">{error.message || String(error)}</p>
        </details>
        <div className="btn-row-inline">
          <button type="button" className="btn btn-primary" onClick={() => this.setState({ error: null })}>
            もう一度表示
          </button>
          {this.props.onReset && (
            <button
              type="button"
              className="btn"
              onClick={() => {
                this.setState({ error: null });
                this.props.onReset?.();
              }}
            >
              {this.props.resetLabel ?? '元に戻す'}
            </button>
          )}
          <button type="button" className="btn" onClick={() => window.location.reload()}>
            画面を再読み込み
          </button>
        </div>
      </div>
    );
  }
}
