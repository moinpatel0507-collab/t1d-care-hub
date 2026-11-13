import './backend.js';   // sign-in, family link and storage (must load first)
import './app.js';       // the care hub itself
import { registerSW } from 'virtual:pwa-register';
registerSW({ immediate: true });
