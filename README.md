## Dev Setup

1. Install taskpane dependencies
   ```bash
   npm ci
   ```
2. Install launch event dependencies
   ```bash
   cd launchEvent
   npm ci
   ```
3. Run add-in and launch event
   ```bash
   npm run dev
   ```
4. Open add-in url at https://localhost:3000
5. Export localhost certificate and import to trusted roots
6. Sideload manifest.xml
