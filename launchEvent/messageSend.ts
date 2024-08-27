import { createNestablePublicClientApplication, IPublicClientApplication, LogLevel } from '@azure/msal-browser'

let clientApp: IPublicClientApplication | undefined = undefined
let clientAppInitialized = false

async function initializeClientApp() {
        if (clientAppInitialized) return

        clientApp = await createNestablePublicClientApplication({
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
        })

        clientAppInitialized = true
}

export async function onMessageSend(event: Office.MailboxEvent) {
    try {
        console.log('onMessageSend')

        console.log('initializeClientApp')
        await initializeClientApp()

        console.log('getGraphAccessToken')
        const accessToken = await getGraphAccessToken()
        console.log('accessToken: ' + accessToken)

        event.completed({ allowEvent: true })
    } catch (error) {
        console.error('Unable to get Graph access token: ' + error)
    }
}

async function getGraphAccessToken() {
    if (!clientApp) {
        console.error('clientApp not initialized')
        return ''
    }

    console.log('acquireTokenSilent')
    const tokenRequest = { scopes: ['Mail.ReadWrite'] }
    const account = await clientApp.acquireTokenSilent(tokenRequest)
    return account.accessToken
}
