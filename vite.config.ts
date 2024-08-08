import { fileURLToPath, URL } from 'node:url'
import { defineConfig } from 'vite'
import vue from '@vitejs/plugin-vue'
import basicSsl from '@vitejs/plugin-basic-ssl'

export default () => {
    return defineConfig({
        build: {
            sourcemap: true
        },
        plugins: [
            vue({
                template: {
                    compilerOptions: {}
                }
            }),
            basicSsl()
        ],
        resolve: {
            alias: {
                '@': fileURLToPath(new URL('./src', import.meta.url))
            }
        },
        preview: {
            port: 3000
        },
        server: {
            port: 3000
        }
    })
}
