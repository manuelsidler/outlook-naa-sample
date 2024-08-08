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
const redirectUri = 'https://localhost:3000'

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
    try {
        const authResult = await acquireTokenSilently()
        console.log(authResult.accessToken)
    } catch (error) {
        console.error('Unable to acquire token silently', error)

        try {
            const authResult = await clientApp.acquireTokenPopup({
                scopes,
                redirectUri
            })
            console.log(authResult.accessToken)
        } catch (error) {
            console.error('Unable to acquire token interactively', error)
        }
    }
}

async function acquireTokenSilently() {
    const account = clientApp.getAccountByUsername(loginHint)
    return account
        ? await clientApp.acquireTokenSilent({
              scopes,
              account,
              redirectUri
          })
        : await clientApp.ssoSilent({
              loginHint,
              scopes,
              redirectUri
          })
}
</script>

<template>
    <button @click="getGraphAccessToken">Get Graph access token</button>
</template>
