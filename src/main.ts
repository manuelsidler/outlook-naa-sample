/// <reference path="../node_modules/@types/office-js/index.d.ts" />

import { createApp } from 'vue'
import App from './App.vue'

// workaround for router bug https://github.com/OfficeDev/office-js/pull/2808
function boot(): Promise<void> {
    const replaceState = window.history.replaceState
    const pushState = window.history.pushState

    return new Promise((resolve, reject) => {
        const script = document.createElement('script')
        script.src = 'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js'

        script.onload = () => {
            Office.onReady(({ host }) => {
                window.history.replaceState = replaceState
                window.history.pushState = pushState

                if (host) {
                    resolve()
                } else {
                    reject(new Error('The application has to run from within Office'))
                }
            })
        }

        document.body.appendChild(script)
    })
}

boot().then(() => {
    const app = createApp(App)
    app.mount('#app')
})
