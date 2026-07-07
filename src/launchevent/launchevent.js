import { addDiagnostic } from "../shared/diagnostics";

/*
 * Outlook Agenda Enforcer — OnAppointmentSend handler v10
 *
 * IMPORTANT: For event-based add-ins, Office.actions.associate() MUST be called
 * at the top level of the script — NOT inside Office.onReady() or any callback.
 * Outlook resolves the handler name synchronously when the send event fires.
 * Ref: https://learn.microsoft.com/office/dev/add-ins/outlook/autolaunch
 */
addDiagnostic("runtime-script-loaded", {
  path: typeof window !== "undefined" ? window.location.pathname : "unknown"
});

Office.onReady()
  .then(function (info) {
    addDiagnostic("office-ready", {
      host: info.host,
      platform: info.platform
    });
  })
  .catch(function (error) {
    addDiagnostic("office-ready-failed", String(error));
  });

try {
  Office.actions.associate("onAppointmentSendHandler", onAppointmentSendHandler);
  addDiagnostic("handler-associated", "onAppointmentSendHandler");
} catch (error) {
  addDiagnostic("handler-associate-failed", String(error));
}

function completeEvent(event, payload, stage) {
  addDiagnostic(stage, payload);
  event.completed(payload);
}

function onAppointmentSendHandler(event) {
  addDiagnostic("send-handler-entered", {
    itemType: Office.context && Office.context.mailbox && Office.context.mailbox.item
      ? Office.context.mailbox.item.itemType
      : "unknown"
  });

  Office.context.mailbox.item.body.getTypeAsync(function (typeResult) {
    if (typeResult.status === Office.AsyncResultStatus.Failed) {
      completeEvent(event, {
        allowEvent: false,
        errorMessage: "[v10] System error: could not determine body type. Please try again."
      }, "send-blocked-type-read-error");
      return;
    }

    var isHtml = typeResult.value === Office.CoercionType.Html;
    addDiagnostic("body-type-detected", { bodyType: typeResult.value, isHtml: isHtml });

    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Text,
      function (asyncResult) {
        addDiagnostic("body-read-completed", {
          status: asyncResult.status,
          error: asyncResult.error ? asyncResult.error.message : ""
        });

        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          completeEvent(event, {
            allowEvent: false,
            errorMessage: "[v10] System error: could not read the meeting body. Please try again."
          }, "send-blocked-body-read-error");
          return;
        }

        var bodyText = asyncResult.value ? asyncResult.value.trim() : "";
        addDiagnostic("body-read-succeeded", {
          length: bodyText.length,
          preview: bodyText.slice(0, 80)
        });

        if (bodyText.length < 10) {
          // Body is blank — use setAsync to write the template as the full body.
          // setAsync is more reliable than prependAsync on a blank body.
          var htmlTemplate = [
            "<p><strong>Summary:</strong> [Brief description of the meeting purpose]</p>",
            "<p><strong>Agenda:</strong></p>",
            "<ul>",
            "<li>[Topic 1 &ndash; Owner &ndash; Est. time]</li>",
            "<li>[Topic 2 &ndash; Owner &ndash; Est. time]</li>",
            "<li>[Topic 3 &ndash; Owner &ndash; Est. time]</li>",
            "</ul>",
            "<p><em>Please complete the agenda above before sending this invite.</em></p>"
          ].join("");

          var coercion = isHtml ? Office.CoercionType.Html : Office.CoercionType.Text;
          var template = isHtml ? htmlTemplate : htmlTemplate.replace(/<[^>]+>/g, "").replace(/&ndash;/g, "-");

          addDiagnostic("agenda-template-set-started", { coercionType: isHtml ? "Html" : "Text", templateLength: template.length });

          Office.context.mailbox.item.body.setAsync(
            template,
            { coercionType: coercion },
            function (setResult) {
              addDiagnostic("agenda-template-set-completed", {
                status: setResult.status,
                error: setResult.error ? setResult.error.message : ""
              });
              if (setResult.status === Office.AsyncResultStatus.Failed) {
                completeEvent(event, {
                  allowEvent: false,
                  errorMessage: "[v10] Agenda required, but template could not be inserted. Please add an agenda before sending."
                }, "send-blocked-set-error");
                return;
              }
              completeEvent(event, {
                allowEvent: false,
                errorMessage: "[v10] An agenda template has been added — please complete it before sending."
              }, "send-blocked-template-added");
            }
          );
        } else {
          completeEvent(event, { allowEvent: true }, "send-allowed");
        }
      }
    );
  });
}

