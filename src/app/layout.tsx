import '@/styles/globals.css'
import { OfficeEnvironmentProvider } from '@/components/OfficeEnvironmentProvider'
import { ErrorBoundary } from '@/components/ErrorBoundary'
import { OfficePolyfillGuard } from '@/components/OfficePolyfillGuard'
import Script from 'next/script'

export default function RootLayout({
  children,
}: {
  children: React.ReactNode
}) {
  return (
    <html lang="zh-CN" suppressHydrationWarning>
      <head>
        {/* 
          CRITICAL: Inline polyfill to ensure it loads FIRST, before any Next.js scripts.
          This patches window.history API which is restricted in Office iframe.
        */}
        <script
          dangerouslySetInnerHTML={{
            __html: `
(function() {
  'use strict';
  console.log('[Office Polyfill] Initializing inline...');
  
  try {
    // Create or patch window.history
    if (!window.history || typeof window.history !== 'object') {
      window.history = { length: 1, scrollRestoration: 'auto', state: null };
    }
    
    // Patch replaceState - critical for Next.js
    if (!window.history.replaceState || window.history.replaceState === null || typeof window.history.replaceState !== 'function') {
      console.log('[Office Polyfill] Patching history.replaceState');
      window.history.replaceState = function(state, title, url) {
        console.debug('[Office Polyfill] replaceState called:', url);
        try {
          Object.defineProperty(window.history, 'state', {
            value: state,
            writable: true,
            configurable: true
          });
        } catch (e) {}
      };
    }
    
    // Patch pushState
    if (!window.history.pushState || window.history.pushState === null || typeof window.history.pushState !== 'function') {
      console.log('[Office Polyfill] Patching history.pushState');
      window.history.pushState = function(state, title, url) {
        console.debug('[Office Polyfill] pushState called:', url);
        try {
          Object.defineProperty(window.history, 'state', {
            value: state,
            writable: true,
            configurable: true
          });
        } catch (e) {}
      };
    }
    
    // Patch other methods
    ['back', 'forward', 'go'].forEach(function(method) {
      if (!window.history[method] || window.history[method] === null || typeof window.history[method] !== 'function') {
        window.history[method] = function() {
          console.debug('[Office Polyfill] history.' + method + ' called (no-op)');
        };
      }
    });
    
    // Ensure state property
    if (!('state' in window.history)) {
      Object.defineProperty(window.history, 'state', {
        get: function() { return null; },
        set: function(v) {},
        configurable: true,
        enumerable: true
      });
    }
    
    // Ensure length property
    if (!('length' in window.history)) {
      Object.defineProperty(window.history, 'length', {
        get: function() { return 1; },
        configurable: true,
        enumerable: true
      });
    }
    
    console.log('[Office Polyfill] ✓ Initialized successfully');
  } catch (error) {
    console.error('[Office Polyfill] Fatal error:', error);
  }
})();
            `,
          }}
        />
        
        <title>DocuPilot - Office AI Assistant</title>
        <meta name="description" content="Intelligent Office Add-in powered by Claude" />
      </head>
      <body className="antialiased" suppressHydrationWarning>
        {/* Load external polyfill as backup */}
        <Script src="/office-polyfill.js" strategy="beforeInteractive" />
        
        <OfficePolyfillGuard />
        <ErrorBoundary>
          <OfficeEnvironmentProvider>
            {children}
          </OfficeEnvironmentProvider>
        </ErrorBoundary>
      </body>
    </html>
  )
}
