import { createNestablePublicClientApplication, LogLevel } from '@azure/msal-browser'

export function onMessageSend(event: Office.MailboxEvent) {
    getGraphAccessToken()
        .then((accessToken) => console.log(accessToken))
        .then(() => event.completed({ allowEvent: true }))
        .catch((error) => {
            console.log(error)
            event.completed({ allowEvent: true })
        })
}

function getGraphAccessToken(): Promise<string> {
    return new Promise<string>((resolve, reject) => {
        createNestablePublicClientApplication({
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
            .then((clientApp) => {
                const scopes = ['Mail.ReadWrite']
                const redirectUri = 'https://localhost:3000'
                const loginHint = Office.context.mailbox.userProfile.emailAddress

                const account = clientApp.getAccountByUsername(loginHint)
                const acquireTokenSilent = account
                    ? clientApp.acquireTokenSilent({
                          scopes,
                          account,
                          redirectUri
                      })
                    : clientApp.ssoSilent({
                          loginHint,
                          scopes,
                          redirectUri
                      })

                acquireTokenSilent
                    .then((result) => resolve(result.accessToken))
                    .catch((error) => {
                        console.log('acquire token silently failed. Get by popup...')
                        console.error(error)

                        clientApp
                            .acquireTokenPopup({
                                scopes,
                                redirectUri
                            })
                            .then((popupResult) => resolve(popupResult.accessToken))
                            .catch((error) => {
                                console.log('acquire token by popup failed.')
                                console.error(error)
                                reject()
                            })
                    })
            })
            .catch((error) => {
                console.error(error)
                reject()
            })
    })
}
