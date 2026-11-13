// Runs on Vercel at deploy time. No npm packages needed.
// Copies the app into "dist" and writes config.js from the two Supabase keys in Vercel's Environment Variables.
const fs = require('fs');
const files = ['index.html', 'main.js', 'backend.js', 'app.js', 'styles.css', 'sw.js', 'manifest.webmanifest',
  'icon-192.png', 'icon-512.png', 'icon-512-maskable.png', 'apple-touch-icon.png', 'favicon-32.png', 'icon.svg'];
const missing = files.filter(f => !fs.existsSync(f));
if (missing.length) { console.error('\n*** MISSING FILES in the GitHub repository: ' + missing.join(', ') + ' ***\nUpload them to the top level of the repository and try again.\n'); process.exit(1); }
fs.rmSync('dist', { recursive: true, force: true }); fs.mkdirSync('dist');
for (const f of files) fs.copyFileSync(f, 'dist/' + f);
const url = (process.env.VITE_SUPABASE_URL || '').trim().replace(/^["']|["']$/g, '');
const key = (process.env.VITE_SUPABASE_ANON_KEY || '').trim().replace(/^["']|["']$/g, '');
if (!url || !key) console.warn('\n*** WARNING: VITE_SUPABASE_URL or VITE_SUPABASE_ANON_KEY is not set in Vercel. The app will say "Almost there". ***\n');
else if (!/^https:\/\/[a-z0-9-]+\.supabase\.co\/?$/.test(url)) console.warn('\n*** WARNING: VITE_SUPABASE_URL does not look like https://xxxx.supabase.co ***\n');
fs.writeFileSync('dist/config.js', 'window.T1D_CONFIG = ' + JSON.stringify({ url: url.replace(/\/$/, ''), key }) + ';\n');
fs.writeFileSync('dist/sw.js', fs.readFileSync('sw.js', 'utf8').replace('__VERSION__', Date.now().toString(36)));
console.log('T1D Care Hub built OK: ' + files.length + ' files + config.js' + (url ? ' (database: ' + url + ')' : ''));
