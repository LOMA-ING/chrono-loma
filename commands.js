/*
 * Chrono LOMA - gestionnaire d'envoi (OnMessageSend).
 * Aucune dépendance externe : fonctionne dans tous les Outlook (classique, nouveau, web, Mac).
 * Il se contente de vérifier si une décision « chrono » a été prise dans le volet :
 *   - "fait:<numéro>" ou "non"  -> l'envoi est autorisé
 *   - rien                      -> Outlook pose la question « chrono ou pas ? »
 */
/* global Office */

Office.onReady(function () {});

var QUESTION =
  "Ce mail doit-il recevoir un numéro de chrono LOMA ?\n\n" +
  "OUI : cliquez sur « Chrono LOMA », choisissez l'opération, le rédacteur et l'expéditeur, " +
  "cliquez sur « Attribuer un numéro », puis cliquez à nouveau sur Envoyer.\n\n" +
  "NON : cliquez sur le bouton qui envoie le mail sans vérification.";

function onMessageSendHandler(event) {
  var item = Office.context.mailbox.item;
  try {
    item.sessionData.getAllAsync(function (res) {
      var etat = "";
      if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
        etat = res.value.chrono || "";
      }
      if (etat === "non" || etat.indexOf("fait:") === 0) {
        event.completed({ allowEvent: true });
        return;
      }
      event.completed({
        allowEvent: false,
        errorMessage: QUESTION,
        cancelLabel: "Chrono LOMA",
        commandId: "btnChronoLoma",
        sendModeOverride: "promptUser"
      });
    });
  } catch (e) {
    // En cas de souci technique, on ne bloque jamais l'envoi.
    event.completed({ allowEvent: true });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
