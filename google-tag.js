(function () {
  "use strict";

  var TAG_ID = "AW-18442609599";
  var STORAGE_KEY = "visapilot_google_consent";
  var savedChoice = null;

  try {
    savedChoice = window.localStorage.getItem(STORAGE_KEY);
  } catch (_) {
    savedChoice = null;
  }

  window.dataLayer = window.dataLayer || [];
  window.gtag = window.gtag || function () {
    window.dataLayer.push(arguments);
  };

  var granted = savedChoice === "accepted";
  window.gtag("consent", "default", {
    ad_storage: granted ? "granted" : "denied",
    analytics_storage: granted ? "granted" : "denied",
    ad_user_data: granted ? "granted" : "denied",
    ad_personalization: granted ? "granted" : "denied",
    wait_for_update: 500
  });
  window.gtag("js", new Date());
  window.gtag("config", TAG_ID);

  window.visaPilotTrackRegistration = function (registrationId) {
    if (!registrationId) return false;
    var conversionKey = "visapilot_registration_conversion_" + String(registrationId);
    try {
      if (window.localStorage.getItem(conversionKey) === "sent") return false;
    } catch (_) {
      // Continue without persistent deduplication when storage is unavailable.
    }

    window.gtag("event", "conversion", {
      send_to: "AW-18442609599/1DlVCMrtovQcEL_Hj9pE",
      value: 1.0,
      currency: "GBP",
      transaction_id: String(registrationId)
    });

    try {
      window.localStorage.setItem(conversionKey, "sent");
    } catch (_) {
      // The conversion has still been queued for this page.
    }
    return true;
  };

  var googleScript = document.createElement("script");
  googleScript.async = true;
  googleScript.src = "https://www.googletagmanager.com/gtag/js?id=" + encodeURIComponent(TAG_ID);
  document.head.appendChild(googleScript);

  function saveChoice(choice) {
    try {
      window.localStorage.setItem(STORAGE_KEY, choice);
    } catch (_) {
      // Consent still applies for this page when browser storage is unavailable.
    }
  }

  function updateConsent(accepted) {
    window.gtag("consent", "update", {
      ad_storage: accepted ? "granted" : "denied",
      analytics_storage: accepted ? "granted" : "denied",
      ad_user_data: accepted ? "granted" : "denied",
      ad_personalization: accepted ? "granted" : "denied"
    });
    saveChoice(accepted ? "accepted" : "rejected");
  }

  function initialiseConsentControls() {
    var banner = document.createElement("section");
    banner.className = "vp-consent";
    banner.setAttribute("role", "dialog");
    banner.setAttribute("aria-label", "Cookie choices");
    banner.innerHTML =
      '<div class="vp-consent__copy"><strong>Your privacy, your choice</strong>' +
      '<span>We use optional Google advertising measurement to understand whether our campaigns are helpful. You can accept or reject it. <a href="privacy.html">Privacy notice</a></span></div>' +
      '<div class="vp-consent__actions"><button type="button" data-consent="reject">Reject optional</button><button type="button" class="vp-consent__accept" data-consent="accept">Accept optional</button></div>';

    var settings = document.createElement("button");
    settings.type = "button";
    settings.className = "vp-consent-settings";
    settings.textContent = "Cookie settings";
    settings.setAttribute("aria-label", "Change cookie settings");

    function showBanner() {
      banner.hidden = false;
      settings.hidden = true;
    }

    function hideBanner() {
      banner.hidden = true;
      settings.hidden = false;
    }

    banner.addEventListener("click", function (event) {
      var button = event.target.closest("[data-consent]");
      if (!button) return;
      updateConsent(button.getAttribute("data-consent") === "accept");
      hideBanner();
    });
    settings.addEventListener("click", showBanner);

    document.body.appendChild(banner);
    document.body.appendChild(settings);

    if (savedChoice === "accepted" || savedChoice === "rejected") {
      hideBanner();
    } else {
      showBanner();
    }
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", initialiseConsentControls);
  } else {
    initialiseConsentControls();
  }
})();
