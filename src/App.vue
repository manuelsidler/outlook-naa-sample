<script setup lang="ts">
import {
    createNestablePublicClientApplication,
    LogLevel,
    type IPublicClientApplication
} from '@azure/msal-browser'
import { onMounted } from 'vue'

let clientApp: IPublicClientApplication

const loginHint = Office.context.mailbox.userProfile.emailAddress
const scopes = ['Mail.ReadWrite']

onMounted(async () => {
    const msalConfig = {
        auth: {
            clientId: '45820bbb-d8e8-4c60-af7d-e8805407193a',
            authority: 'https://login.microsoftonline.com/common'
        },
        system: {
                loggerOptions: {
                    logLevel: LogLevel.Verbose,
                    loggerCallback: (level: LogLevel, message: string) => {
                        switch (level) {
                            case LogLevel.Error:
                                console.error(message)
                                return
                            case LogLevel.Info:
                                console.info(message)
                                return
                            case LogLevel.Verbose:
                                console.debug(message)
                                return
                            case LogLevel.Warning:
                                console.warn(message)
                                return
                        }
                    },
                    piiLoggingEnabled: true
                }
            }
    }

    clientApp = await createNestablePublicClientApplication(msalConfig)
})

async function getGraphAccessToken() {
    const authRequest = {
        scopes,
        loginHint
    }

    try {
        console.log('acquire token silently')
        const account = clientApp.getAccountByUsername(loginHint)
        const authResult = account
            ? await clientApp.acquireTokenSilent(authRequest)
            : await clientApp.ssoSilent(authRequest)
        console.log(authResult.accessToken)
    } catch (error) {
        console.error('Unable to acquire token silently', error)

        try {
            console.log('acquire token by popup')
            const authResult = await clientApp.acquireTokenPopup(authRequest)
            console.log(authResult.accessToken)
        } catch (error) {
            console.error('Unable to acquire token interactively', error)
        }
    }
}
</script>

<template>
    <button @click="getGraphAccessToken">Get Graph access token</button>
</template>
