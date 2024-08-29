import { defineConfig } from 'vite'
import { resolve } from 'path'

// https://vitejs.dev/config/
export default defineConfig({
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
