/**
 * Yes / no confirmation (結合をやり直す, ファイル追加で結合結果を破棄, タイトルの結合を解除して移動 …).
 * The OK (primary) button has the initial focus: Enter confirms, Esc cancels.
 */
import { Modal } from './Modal';

export interface ConfirmDialogProps {
  title: string;
  message: string;
  okLabel: string;
  cancelLabel?: string;
  /** red OK button (the action discards work) */
  danger?: boolean;
  testId?: string;
  onOk: () => void;
  onCancel: () => void;
}

export function ConfirmDialog(props: ConfirmDialogProps) {
  return (
    <Modal
      title={props.title}
      testId={props.testId ?? 'confirm-dialog'}
      size="sm"
      onClose={props.onCancel}
      initialFocus="[data-testid=confirm-ok]"
      footer={
        <>
          <button type="button" className="btn" onClick={props.onCancel} data-testid="confirm-cancel">
            {props.cancelLabel ?? 'キャンセル'}
          </button>
          <button type="button" className={`btn ${props.danger ? 'btn-danger' : 'btn-primary'}`} onClick={props.onOk} data-testid="confirm-ok">
            {props.okLabel}
          </button>
        </>
      }
    >
      <p className="modal-text">{props.message}</p>
    </Modal>
  );
}
