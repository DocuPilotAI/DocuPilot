"use client";

import { useEffect, useRef } from "react";

/**
 * OfficePolyfillGuard - Ensures history API polyfills remain active
 * 
 * This component runs after React hydration to verify and re-apply
 * polyfills if they were somehow overwritten by Next.js or other code.
 * 
 * It also sets up an interval to continuously monitor for changes.
 */
export function OfficePolyfillGuard() {
  const intervalRef = useRef<NodeJS.Timeout | null>(null);

  useEffect(() => {
    // Polyfill implementation
    const applyPolyfill = () => {
      // Check if window.history exists and has proper methods
      if (typeof window === "undefined") return false;

      let applied = false;

      // Ensure history object exists
      if (!window.history || typeof window.history !== "object") {
        console.warn("[OfficePolyfillGuard] window.history missing, creating");
        (window as any).history = {
          length: 1,
          scrollRestoration: 'auto',
          state: null
        };
        applied = true;
      }

      // Check replaceState
      if (
        !window.history.replaceState ||
        window.history.replaceState === null ||
        typeof window.history.replaceState !== "function"
      ) {
        console.warn("[OfficePolyfillGuard] history.replaceState missing, patching");
        applied = true;

        window.history.replaceState = function (state: any, title: string, url?: string | URL | null) {
          console.debug("[OfficePolyfillGuard] replaceState called (no-op):", url);
          try {
            Object.defineProperty(window.history, "state", {
              value: state,
              writable: true,
              configurable: true,
            });
          } catch (e) {
            // Ignore
          }
        };
      }

      // Check pushState
      if (
        !window.history.pushState ||
        window.history.pushState === null ||
        typeof window.history.pushState !== "function"
      ) {
        console.warn("[OfficePolyfillGuard] history.pushState missing, patching");
        applied = true;

        window.history.pushState = function (state: any, title: string, url?: string | URL | null) {
          console.debug("[OfficePolyfillGuard] pushState called (no-op):", url);
          try {
            Object.defineProperty(window.history, "state", {
              value: state,
              writable: true,
              configurable: true,
            });
          } catch (e) {
            // Ignore
          }
        };
      }

      // Check other methods
      ["back", "forward", "go"].forEach((method) => {
        if (
          !(window.history as any)[method] ||
          (window.history as any)[method] === null ||
          typeof (window.history as any)[method] !== "function"
        ) {
          applied = true;
          (window.history as any)[method] = function () {
            console.debug(`[OfficePolyfillGuard] history.${method} called (no-op)`);
          };
        }
      });

      // Ensure state property
      try {
        if (!('state' in window.history)) {
          Object.defineProperty(window.history, 'state', {
            get: function() { return null; },
            set: function(v) { /* no-op */ },
            configurable: true,
            enumerable: true
          });
        }
      } catch (e) {
        // Ignore
      }

      // Ensure length property
      try {
        if (!('length' in window.history)) {
          Object.defineProperty(window.history, 'length', {
            get: function() { return 1; },
            configurable: true,
            enumerable: true
          });
        }
      } catch (e) {
        // Ignore
      }

      return applied;
    };

    // Apply immediately
    const initialApply = applyPolyfill();
    if (initialApply) {
      console.log("[OfficePolyfillGuard] ✓ Initial polyfill applied");
    } else {
      console.log("[OfficePolyfillGuard] ✓ Polyfills already present");
    }

    // Check again after short delays (for hydration)
    const timer1 = setTimeout(() => {
      if (applyPolyfill()) {
        console.log("[OfficePolyfillGuard] ✓ Polyfill re-applied at 100ms");
      }
    }, 100);
    
    const timer2 = setTimeout(() => {
      if (applyPolyfill()) {
        console.log("[OfficePolyfillGuard] ✓ Polyfill re-applied at 500ms");
      }
    }, 500);

    // Set up continuous monitoring (every 2 seconds) for hot reload scenarios
    intervalRef.current = setInterval(() => {
      if (applyPolyfill()) {
        console.warn("[OfficePolyfillGuard] ⚠️ Polyfill was lost and re-applied (likely hot reload)");
      }
    }, 2000);

    return () => {
      clearTimeout(timer1);
      clearTimeout(timer2);
      if (intervalRef.current) {
        clearInterval(intervalRef.current);
      }
    };
  }, []);

  return null; // This component doesn't render anything
}
