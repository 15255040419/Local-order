// Browser polyfills for Node.js modules
import { Buffer } from 'buffer';

// Make Buffer globally available
if (typeof window !== 'undefined') {
  (window as any).Buffer = Buffer;
  (window as any).global = window;
  (window as any).process = { env: {} };
}
