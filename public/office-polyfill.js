/**
 * Office Add-in Environment Polyfills
 * 
 * This script patches missing or restricted APIs in Office iframe environment.
 * It must be loaded as early as possible, before Next.js hydration.
 */

(function() {
  'use strict';
  
  console.log('[Office Polyfill] Initializing...');
  
  try {
    // ============================================
    // 1. Patch window.history API
    // ============================================
    
    // Create history object if completely missing
    if (!window.history || typeof window.history !== 'object') {
      console.warn('[Office Polyfill] window.history is missing, creating stub');
      window.history = {
        length: 1,
        scrollRestoration: 'auto',
        state: null
      };
    }
    
    // Patch replaceState
    if (!window.history.replaceState || 
        window.history.replaceState === null || 
        typeof window.history.replaceState !== 'function') {
      console.log('[Office Polyfill] Patching history.replaceState');
      
      window.history.replaceState = function(state, title, url) {
        console.debug('[Office Polyfill] history.replaceState called (no-op):', {
          state: state,
          title: title,
          url: url
        });
        // Update state property
        try {
          Object.defineProperty(window.history, 'state', {
            value: state,
            writable: true,
            configurable: true
          });
        } catch (e) {
          // Ignore if we can't define property
        }
      };
    }
    
    // Patch pushState
    if (!window.history.pushState || 
        window.history.pushState === null || 
        typeof window.history.pushState !== 'function') {
      console.log('[Office Polyfill] Patching history.pushState');
      
      window.history.pushState = function(state, title, url) {
        console.debug('[Office Polyfill] history.pushState called (no-op):', {
          state: state,
          title: title,
          url: url
        });
        // Update state property
        try {
          Object.defineProperty(window.history, 'state', {
            value: state,
            writable: true,
            configurable: true
          });
        } catch (e) {
          // Ignore if we can't define property
        }
      };
    }
    
    // Patch other history methods
    ['back', 'forward', 'go'].forEach(function(method) {
      if (!window.history[method] || 
          window.history[method] === null || 
          typeof window.history[method] !== 'function') {
        console.log('[Office Polyfill] Patching history.' + method);
        window.history[method] = function() {
          console.debug('[Office Polyfill] history.' + method + ' called (no-op)');
        };
      }
    });
    
    // Ensure state property exists and is readable
    try {
      if (!('state' in window.history) || window.history.state === undefined) {
        Object.defineProperty(window.history, 'state', {
          get: function() { return null; },
          set: function(v) { /* no-op */ },
          configurable: true,
          enumerable: true
        });
      }
    } catch (e) {
      console.warn('[Office Polyfill] Could not define history.state:', e);
    }
    
    // Ensure length property exists
    try {
      if (!('length' in window.history)) {
        Object.defineProperty(window.history, 'length', {
          get: function() { return 1; },
          configurable: true,
          enumerable: true
        });
      }
    } catch (e) {
      console.warn('[Office Polyfill] Could not define history.length:', e);
    }
    
    // ============================================
    // 2. Freeze history to prevent Next.js from overwriting
    // ============================================
    try {
      // Store references
      var replaceState = window.history.replaceState;
      var pushState = window.history.pushState;
      
      // Re-apply after a short delay (in case something overwrites them)
      setTimeout(function() {
        if (!window.history.replaceState || typeof window.history.replaceState !== 'function') {
          console.log('[Office Polyfill] Re-applying replaceState after delay');
          window.history.replaceState = replaceState;
        }
        if (!window.history.pushState || typeof window.history.pushState !== 'function') {
          console.log('[Office Polyfill] Re-applying pushState after delay');
          window.history.pushState = pushState;
        }
      }, 100);
    } catch (e) {
      console.warn('[Office Polyfill] Could not preserve history methods:', e);
    }
    
    console.log('[Office Polyfill] ✓ History API patched successfully');
    
    // ============================================
    // 3. Log final state for debugging
    // ============================================
    console.log('[Office Polyfill] Final state:', {
      hasHistory: !!window.history,
      hasReplaceState: !!(window.history && window.history.replaceState),
      hasPushState: !!(window.history && window.history.pushState),
      replaceStateType: window.history && typeof window.history.replaceState,
      pushStateType: window.history && typeof window.history.pushState
    });
    
  } catch (error) {
    console.error('[Office Polyfill] Fatal error:', error);
  }
})();
