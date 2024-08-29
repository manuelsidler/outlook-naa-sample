import { onMessageSend } from './messageSend'

declare global {
    let onMessageSendHandler: (event: Office.MailboxEvent) => void
}

// necessary for Mac
Office.onReady(() => {})

// events must be defined globally to prevent esbuild from renaming them
onMessageSendHandler = onMessageSend

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (
    Office.context &&
    (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null)
) {
    Office.actions.associate('onMessageSendHandler', onMessageSendHandler)
}
