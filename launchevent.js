/*
 * Alerte hors horaires de travail – Atelier M3
 * Complément Outlook (Smart Alerts, événement OnMessageSend, mode PromptUser)
 *
 * Au clic sur « Envoyer », si l'heure est hors des horaires de travail,
 * Outlook affiche le message ci-dessous avec les boutons
 * « Envoyer quand même » / « Ne pas envoyer ».
 *
 * Remarque : pas d'import/export dans ce fichier (non supporté par
 * le moteur JavaScript d'Outlook classique sous Windows).
 */

// ---------------------------------------------------------------
// Réglages (modifiables)
// ---------------------------------------------------------------
// Horaires de travail par jour : [heure de début, heure de fin]
// 0 = dimanche, 1 = lundi … 6 = samedi. Jour absent = pas travaillé.
var HORAIRES = {
  1: [9, 18], // lundi    9h-18h
  2: [9, 18], // mardi    9h-18h
  3: [9, 12], // mercredi 9h-18h
  4: [9, 18], // jeudi    9h-18h
  5: [9, 17]  // vendredi 9h-17h
};
var FUSEAU = "Europe/Paris";

var MESSAGE =
  "Envoi en dehors des horaires de travail. " +
  "Vous pouvez programmer cet envoi. " +
  "Confirmer l'envoi maintenant ?";

// ---------------------------------------------------------------
// Heure et jour à Paris (avec repli sur l'heure de l'ordinateur)
// ---------------------------------------------------------------
function heureEtJour(date) {
  try {
    var parties = new Intl.DateTimeFormat("en-US", {
      timeZone: FUSEAU,
      hour: "numeric",
      hourCycle: "h23",
      weekday: "short"
    }).formatToParts(date);

    var heure = null;
    var jour = null;
    var correspondanceJours = { Sun: 0, Mon: 1, Tue: 2, Wed: 3, Thu: 4, Fri: 5, Sat: 6 };

    for (var i = 0; i < parties.length; i++) {
      if (parties[i].type === "hour") heure = parseInt(parties[i].value, 10) % 24;
      if (parties[i].type === "weekday") jour = correspondanceJours[parties[i].value];
    }
    if (heure !== null && jour !== undefined && jour !== null) {
      return { heure: heure, jour: jour };
    }
  } catch (e) {
    // Fuseau non pris en charge : on utilise l'heure locale de l'ordinateur.
  }
  return { heure: date.getHours(), jour: date.getDay() };
}

function estHorsHoraires(date) {
  var t = heureEtJour(date);
  var plage = HORAIRES[t.jour];
  if (!plage) return true; // week-end ou jour non travaillé
  return !(t.heure >= plage[0] && t.heure < plage[1]);
}

// ---------------------------------------------------------------
// Gestionnaire de l'événement d'envoi
// ---------------------------------------------------------------
function onMessageSendHandler(event) {
  try {
    if (estHorsHoraires(new Date())) {
      // allowEvent: false + mode PromptUser = fenêtre d'avertissement
      // avec « Envoyer quand même » / « Ne pas envoyer ».
      event.completed({ allowEvent: false, errorMessage: MESSAGE });
    } else {
      event.completed({ allowEvent: true });
    }
  } catch (e) {
    // En cas d'erreur, on ne bloque jamais l'envoi.
    event.completed({ allowEvent: true });
  }
}

// Association du gestionnaire à l'événement d'envoi.
// Outlook Web / nouvel Outlook / Mac exigent que la bibliothèque Office
// soit initialisée (Office.onReady) avant l'association ; Outlook classique
// Windows accepte l'association immédiate. On fait les deux, sans erreur.
function associerGestionnaire() {
  try {
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
  } catch (e) {
    // Office pas encore prêt : l'association se fera dans Office.onReady.
  }
}

if (typeof Office !== "undefined") {
  if (typeof Office.onReady === "function") {
    Office.onReady(function () {
      associerGestionnaire();
    });
  }
  associerGestionnaire();
}

// Pour les tests hors Outlook uniquement (ignoré par Outlook)
if (typeof module !== "undefined" && module.exports) {
  module.exports = { estHorsHoraires: estHorsHoraires };
}
