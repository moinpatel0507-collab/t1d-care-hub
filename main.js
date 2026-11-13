import './backend.js';   // sign-in, family link and storage (must load first)
import './app.js';       // the care hub itself
if ('serviceWorker' in navigator) addEventListener('load', () => navigator.serviceWorker.register('/sw.js').catch(() => {}));
