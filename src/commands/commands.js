/* global Office */

const FLOW_URL = "https://default783a2c3aadb945ef8d986601d1686f.35.environment.api.powerplatform.com/powerautomate/automations/direct/workflows/69ef74a9b09a4859a90ed5fdd79f09fa/triggers/manual/paths/invoke?api-version=1&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=sXz7nrvkzJQnVMnLNKusvLyiJvvOyC0oHDYM2R264Ec";

let isSending = false;

Office.onReady(() => {
  Office.actions.associate("monitorarReuniao", monitorarReuniao);
});

async function monitorarReuniao(event) {
  if (isSending) {
    event.completed();
    return;
  }

  isSending = true;
  const startedAt = Date.now();
  const correlationId = createCorrelationId();

  try {
    showInfo("monitorar-status", "Enviando reunião para monitoramento...");

    const item = Office.context.mailbox.item;

    if (!item) {
      throw new Error("Item do Outlook não encontrado.");
    }

    const bodyText = await getBodyText(item);
    const subject = getDirectValue(item.subject);
    const start = await getAsyncProperty(item.start);
    const end = await getAsyncProperty(item.end);
    const location = await getAsyncProperty(item.location);
    const requiredAttendees = await getRecipients(item.requiredAttendees);
    const optionalAttendees = await getRecipients(item.optionalAttendees);
    const meetingId = extractTeamsMeetingId(bodyText);
    const teamsJoinUrl = extractTeamsJoinUrl(bodyText);

    if (!subject) {
      throw new Error("A reunião precisa ter um assunto antes de ser monitorada.");
    }

    const payload = {
      correlationId,
      source: "Outlook Add-in",
      action: "Monitorar",
      userEmail: Office.context.mailbox.userProfile.emailAddress,
      userDisplayName: Office.context.mailbox.userProfile.displayName,
      itemId: item.itemId || null,
      subject,
      start,
      end,
      location,
      requiredAttendees,
      optionalAttendees,
      meetingId,
      teamsJoinUrl,
      bodyPreview: bodyText ? bodyText.substring(0, 1500) : "",
      durationMs: Date.now() - startedAt,
      sentAt: new Date().toISOString()
    };

    await sendToPowerAutomate(payload);

    removeMessage("monitorar-status");
    showInfo("monitorar-success", "Reunião enviada para monitoramento.");
  } catch (error) {
    removeMessage("monitorar-status");
    console.error("Erro ao monitorar reunião", error);
    showError("monitorar-error", getFriendlyErrorMessage(error));
  } finally {
    isSending = false;
    event.completed();
  }
}

async function sendToPowerAutomate(payload) {
  const response = await fetch(FLOW_URL, {
    method: "POST",
    headers: {
      "Content-Type": "application/json"
    },
    body: JSON.stringify(payload)
  });

  if (!response.ok && response.type !== "opaque") {
    throw new Error(`Falha ao chamar o fluxo. Status: ${response.status}`);
  }
}

function getDirectValue(value) {
  if (!value) return "";
  if (typeof value === "string") return value;
  return "";
}

function getAsyncProperty(asyncProperty) {
  return new Promise((resolve) => {
    if (!asyncProperty || typeof asyncProperty.getAsync !== "function") {
      resolve("");
      return;
    }

    asyncProperty.getAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        resolve("");
      }
    });
  });
}

function getBodyText(item) {
  return new Promise((resolve) => {
    if (!item.body || typeof item.body.getAsync !== "function") {
      resolve("");
      return;
    }

    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value || "");
      } else {
        resolve("");
      }
    });
  });
}

function getRecipients(recipientsProperty) {
  return new Promise((resolve) => {
    if (!recipientsProperty || typeof recipientsProperty.getAsync !== "function") {
      resolve([]);
      return;
    }

    recipientsProperty.getAsync((result) => {
      if (result.status !== Office.AsyncResultStatus.Succeeded || !Array.isArray(result.value)) {
        resolve([]);
        return;
      }

      resolve(
        result.value.map((recipient) => ({
          displayName: recipient.displayName || "",
          emailAddress: recipient.emailAddress || ""
        }))
      );
    });
  });
}

function extractTeamsMeetingId(text) {
  if (!text) return "";

  const patterns = [
    /ID da reunião:\s*([\d\s]+)/i,
    /Meeting ID:\s*([\d\s]+)/i,
    /ID de la reunión:\s*([\d\s]+)/i
  ];

  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match && match[1]) {
      return match[1].replace(/\s/g, "");
    }
  }

  return "";
}

function extractTeamsJoinUrl(text) {
  if (!text) return "";

  const match = text.match(/https:\/\/teams\.microsoft\.com\/l\/meetup-join\/[^\s<>\"]+/i);
  return match ? match[0] : "";
}

function showInfo(key, message) {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return;

  item.notificationMessages.replaceAsync(key, {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message,
    icon: "icon16",
    persistent: false
  });
}

function showError(key, message) {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return;

  item.notificationMessages.replaceAsync(key, {
    type: Office.MailboxEnums.ItemNotificationMessageType.ErrorMessage,
    message
  });
}

function removeMessage(key) {
  const item = Office.context.mailbox.item;
  if (!item || !item.notificationMessages) return;
  item.notificationMessages.removeAsync(key);
}

function getFriendlyErrorMessage(error) {
  const message = error && error.message ? error.message : "Erro inesperado.";

  if (message.includes("assunto")) {
    return message;
  }

  return "Não foi possível enviar a reunião para monitoramento. Tente novamente.";
}

function createCorrelationId() {
  if (crypto && typeof crypto.randomUUID === "function") {
    return crypto.randomUUID();
  }

  return `corr-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}
