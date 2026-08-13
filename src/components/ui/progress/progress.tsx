import * as React from 'react'
import { cn } from '@/lib/utils'

export interface ProgressProps
  extends React.HTMLAttributes<HTMLDivElement> {
  value?: number
  showValue?: boolean
  label?: string
  indeterminate?: boolean
}

const Progress = React.forwardRef<
  HTMLDivElement,
  ProgressProps
>(({ className, value = 0, showValue = true, label, indeterminate = false, ...props }, ref) => {
  const pct     = Math.min(Math.max(value ?? 0, 0), 100)
  const filled  = Math.round((pct / 100) * 20)
  const empty   = 20 - filled
  const barText = '='.repeat(filled) + ' '.repeat(empty)

  return (
    <div className="sci-progress" style={{ display: 'flex', flexDirection: 'column', gap: '0.25rem', width: '100%' }}>
      {label && (
        <span
          style={{
            fontSize:      '0.65rem',
            color:         'var(--text-muted)',
            letterSpacing: '0.1em',
            textTransform: 'uppercase',
          }}
        >
          {label}
        </span>
      )}

      <div
        ref={ref}
        className={cn('sci-progress-root relative w-full overflow-hidden rounded-none', className)}
        role="progressbar"
        aria-valuemin={0}
        aria-valuemax={100}
        aria-valuenow={indeterminate ? undefined : pct}
        {...props}
      >
        {indeterminate ? (
          <div className="sci-progress-line" aria-hidden="true">
            <span className="sci-progress-bracket">[</span>
            <span className="sci-progress-track">
              <span className="sci-progress-sweep" />
            </span>
            <span className="sci-progress-bracket">]</span>
          </div>
        ) : (
          <div
            style={{
              display:    'grid',
              gridTemplateColumns: 'auto minmax(0, 1fr) auto auto',
              alignItems: 'center',
              gap:        '0.25rem',
              fontFamily: 'var(--font-mono)',
              fontSize:   '0.75rem',
              color:      'var(--color-green)',
              textShadow: 'var(--text-glow-green)',
            }}
          >
            <span style={{ color: 'var(--text-muted)' }}>[</span>
            <span style={{ letterSpacing: '-0.05em', whiteSpace: 'pre', overflow: 'hidden' }}>{barText}</span>
            <span style={{ color: 'var(--text-muted)' }}>]</span>
            {showValue && (
              <span style={{ color: 'var(--text-muted)', fontSize: '0.65rem', marginLeft: '0.25rem' }}>
                {Math.round(pct)}%
              </span>
            )}
          </div>
        )}

      </div>
    </div>
  )
})
Progress.displayName = 'Progress'

export { Progress }
