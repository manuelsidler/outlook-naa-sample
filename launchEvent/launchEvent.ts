import { onMessageSend } from './messageSend'

// IMPORTANT: The JavaScript code of event-based add-ins that run in Outlook on Windows only supports ECMAScript 2016 and earlier specifications
// -> avoid using async/await
// -> avoid using conditional (ternary) operator

// necessary for Mac
Office.onReady(() => {})

// *** Message Send Event ***
function onMessageSendHandler(event: Office.MailboxEvent) {
    onMessageSend(event)
}

// IMPORTANT: To ensure your add-in is supported in the Outlook client on Windows, remember to map the event handler name specified in the manifest to its JavaScript counterpart.
if (
    Office.context &&
    (Office.context.platform === Office.PlatformType.PC || Office.context.platform == null)
) {
    Office.actions.associate('onMessageSendHandler', onMessageSendHandler)
}
