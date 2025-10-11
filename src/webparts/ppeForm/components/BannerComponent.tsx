import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react';

export type BannerKind = 'info' | 'success' | 'warning' | 'error';

export interface BannerComponentProps {
  // Text to display. When undefined/empty, nothing is rendered.
  text?: string;

  // Visual style
  kind?: BannerKind; // default: 'info'

  // Show the close (X) button
  dismissible?: boolean; // default: true

  // Auto-hide config (applies only to this instance)
  autoHideMs?: number; // e.g., 40000 for 40s. Omit for no auto-dismiss.

  // Fade configuration
  fade?: boolean;       // default: false (enable fade on this instance only)
  fadeMs?: number;      // default: 600ms

  // Multiline bar (prevents truncation for long text)
  multiline?: boolean; // default: true

  // Optional extra content (buttons/links)
  actions?: React.ReactNode;

  // Callbacks and styling hooks
  onDismiss?: () => void;                // called when dismissed (manual or auto)
  className?: string;
  style?: React.CSSProperties;
}

const kindToType = (k?: BannerKind): MessageBarType => {
  switch (k) {
    case 'success': return MessageBarType.success;
    case 'warning': return MessageBarType.warning;
    case 'error': return MessageBarType.error;
    case 'info':
    default: return MessageBarType.info;
  }
};

const BannerComponent: React.FC<BannerComponentProps> = ({
  text,
  kind = 'info',
  dismissible = true,
  autoHideMs,
  fade = false,
  fadeMs = 600,
  multiline = true,
  actions,
  onDismiss,
  className,
  style,
}) => {
  const [isFading, setIsFading] = React.useState(false);
  const hideTimerRef = React.useRef<number | null>(null);
  const fadeTimerRef = React.useRef<number | null>(null);

  // Clear timers on unmount or text change
  React.useEffect(() => {
    return () => {
      if (hideTimerRef.current) { window.clearTimeout(hideTimerRef.current); hideTimerRef.current = null; }
      if (fadeTimerRef.current) { window.clearTimeout(fadeTimerRef.current); fadeTimerRef.current = null; }
    };
  }, []);

  React.useEffect(() => {
    // Reset fade on new text
    setIsFading(false);

    // Clear any previous timers
    if (hideTimerRef.current) { window.clearTimeout(hideTimerRef.current); hideTimerRef.current = null; }
    if (fadeTimerRef.current) { window.clearTimeout(fadeTimerRef.current); fadeTimerRef.current = null; }

    if (!text) return;

    if (autoHideMs && autoHideMs > 0) {
      // Schedule fade slightly before hide
      if (fade) {
        const startFadeAt = Math.max(0, autoHideMs - fadeMs);
        fadeTimerRef.current = window.setTimeout(() => setIsFading(true), startFadeAt) as unknown as number;
      }

      hideTimerRef.current = window.setTimeout(() => {
        // If fading is enabled and not already started, do a quick fade
        if (fade && !isFading) setIsFading(true);

        // Give the fade time to complete, then notify parent to clear banner
        window.setTimeout(() => {
          onDismiss?.();
          // Reset internal state
          setIsFading(false);
        }, fade ? fadeMs : 0);
      }, autoHideMs) as unknown as number;
    }
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [text, autoHideMs, fade, fadeMs]);

  const handleManualDismiss = React.useCallback(() => {
    if (fade) {
      setIsFading(true);
      window.setTimeout(() => {
        onDismiss?.();
        setIsFading(false);
      }, fadeMs);
    } else {
      onDismiss?.();
    }
  }, [fade, fadeMs, onDismiss]);

  if (!text) return null;

  return (
    <div
      className={className}
      style={{
        transition: `opacity ${fadeMs}ms ease`,
        opacity: isFading ? 0 : 1,
        // You can style padding/background around the bar if desired:
        // padding: 4,
        ...style,
      }}
    >
      <MessageBar
        messageBarType={kindToType(kind)}
        isMultiline={multiline}
        onDismiss={dismissible ? handleManualDismiss : undefined}
        dismissButtonAriaLabel="Close"
      >
        <span>{text}</span>
        {actions}
      </MessageBar>
    </div>
  );
};

export default BannerComponent;