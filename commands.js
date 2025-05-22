// Global variabel för den interna domänen
const INTERNAL_DOMAIN = "riksbank.se";

// Unik nyckel för notifikationsmeddelandet
const NOTIFICATION_KEY = "externalRecipientWarning";

// Strängar för lokalisering
const localizedStrings = {
    "en-US": { externalRecipientsMessage: "One or more recipients are external." },
    "sv-SE": { externalRecipientsMessage: "En eller flera mottagare är externa." },
    "default": { externalRecipientsMessage: "One or more recipients are external." }
};

Office.onReady(info => {
    console.log("External Warning: Office.onReady fired. Host: " + info.host + ", Platform: " + info.platform);
    // Inget specifikt behövs här för detta event-baserade tillägg
    // om du inte har särskild initieringslogik.
});

// Händelsehanterare för när mottagare ändras
async function onRecipientsChangedHandler(event) {
    console.log("External Warning: onRecipientsChangedHandler event triggered.");

    if (!Office.context || !Office.context.mailbox || !Office.context.mailbox.item) {
        console.error("External Warning: Office context or mailbox item not available in onRecipientsChangedHandler.");
        return;
    }

    // Kontrollera att vi är i skrivläge (compose mode)
    if (Office.context.mailbox.item.itemType !== Office.MailboxEnums.ItemType.Message ||
        Office.context.mailbox.item.displayMode !== Office.MailboxEnums.DisplayMode.Edit) {
        console.log("External Warning: Not in message compose mode. Exiting.");
        return;
    }
    
    try {
        const item = Office.context.mailbox.item;
        let hasExternalRecipient = false;

        const toRecipients = item.to ? await getRecipientsAsync(item.to) : [];
        const ccRecipients = item.cc ? await getRecipientsAsync(item.cc) : [];
        const bccRecipients = item.bcc ? await getRecipientsAsync(item.bcc) : [];

        const allRecipients = [...toRecipients, ...ccRecipients, ...bccRecipients];
        console.log("External Warning: All recipients:", JSON.stringify(allRecipients));

        if (allRecipients.length > 0) {
            for (const recipient of allRecipients) {
                if (recipient && recipient.emailAddress) {
                    if (isExternalEmail(recipient.emailAddress)) {
                        hasExternalRecipient = true;
                        break; 
                    }
                }
            }
        }
        
        console.log("External Warning: Has external recipient:", hasExternalRecipient);

        if (hasExternalRecipient) {
            const userLanguage = Office.context.displayLanguage;
            let message = localizedStrings.default.externalRecipientsMessage; 
            if (localizedStrings[userLanguage]) {
                message = localizedStrings[userLanguage].externalRecipientsMessage;
            } else if (userLanguage.startsWith("en") && localizedStrings["en-US"]) {
                message = localizedStrings["en-US"].externalRecipientsMessage;
            } else if (userLanguage.startsWith("sv") && localizedStrings["sv-SE"]) {
                message = localizedStrings["sv-SE"].externalRecipientsMessage;
            }

            item.notificationMessages.addAsync(NOTIFICATION_KEY, {
                type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
                message: message,
                icon: "Icon.Warning", 
                persistent: false 
            }, handleCallback);
        } else {
            item.notificationMessages.removeAsync(NOTIFICATION_KEY, handleCallback);
        }

    } catch (error) {
        console.error("External Warning: Error in onRecipientsChangedHandler: ", error);
    }
}

function getRecipientsAsync(recipientField) {
    return new Promise((resolve, reject) => {
        if (!recipientField) {
            console.warn("External Warning: recipientField is null in getRecipientsAsync.");
            resolve([]);
            return;
        }
        recipientField.getAsync(result => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                console.error("External Warning: Failed to get recipients: ", JSON.stringify(result.error));
                resolve([]); 
            }
        });
    });
}

function isExternalEmail(emailAddress) {
    if (!emailAddress || typeof emailAddress !== 'string') {
        return false;
    }
    const lowerEmail = emailAddress.toLowerCase();
    const domainPart = lowerEmail.substring(lowerEmail.lastIndexOf("@") + 1);
    return domainPart !== INTERNAL_DOMAIN.toLowerCase();
}

function handleCallback(asyncResult) {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        console.error("External Warning: Notification operation failed: " + JSON.stringify(asyncResult.error));
    } else {
        console.log("External Warning: Notification operation successful.");
    }
}

// Registrera funktionen globalt så att den kan anropas från manifestet.
// Detta är nödvändigt för event-based add-ins.
// MER DETALJERAD LOGGNING HÄR:
if (typeof Office !== 'undefined') {
    console.log("External Warning: Office object IS defined.");
    if (Office.actions) {
        console.log("External Warning: Office.actions IS defined.");
        if (Office.actions.associate) {
            console.log("External Warning: Office.actions.associate IS defined. Attempting to associate 'onRecipientsChangedHandler'...");
            try {
                // Kontrollera att funktionen onRecipientsChangedHandler faktiskt är definierad
                if (typeof onRecipientsChangedHandler === 'function') {
                    Office.actions.associate("onRecipientsChangedHandler", onRecipientsChangedHandler);
                    console.log("External Warning: Successfully called Office.actions.associate for onRecipientsChangedHandler.");
                } else {
                    console.error("External Warning: ERROR - onRecipientsChangedHandler is NOT a function when associate was called!");
                }
            } catch (e) {
                console.error("External Warning: ERROR during Office.actions.associate call:", e);
            }
        } else {
            console.error("External Warning: Office.actions.associate is UNDEFINED.");
        }
    } else {
        console.error("External Warning: Office.actions is UNDEFINED. (This is normal if not in an event-based context for some platforms).");
    }
} else {
    console.error("External Warning: Office object is UNDEFINED. office.js did not load/initialize correctly or script is running too early.");
}