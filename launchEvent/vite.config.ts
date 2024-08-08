import { defineConfig } from 'vite'
import { resolve } from 'path'

// https://vitejs.dev/config/
export default defineConfig({
    esbuild: {
        // we have to keep names for Outlook Message Handlers (eg. onMessageComposeHandler)
        minifyIdentifiers: false,
        keepNames: true
    },
    build: {
        target: 'es2016',
        outDir: '../public',
        lib: {
            entry: resolve(__dirname, 'launchEvent.ts'),
            name: 'OutlookLaunchEvent',
            fileName: 'launchEvent'
        },
        rollupOptions: {
            output: {}
        }
    }
})
