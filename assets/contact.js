/* Builds the support email link in the browser, so the address never appears
   in the page text or source for address-harvesting bots. Links marked
   data-contact keep a fallback href (the support page) until this runs. */
(function () {
  "use strict";
  var user = ["youhave", "reached", "omdave"].join("");
  var host = ["gmail", "com"].join(".");
  var address = user + String.fromCharCode(64) + host;
  var links = document.querySelectorAll("a[data-contact]");
  for (var i = 0; i < links.length; i++) {
    var subject = links[i].getAttribute("data-subject");
    links[i].setAttribute("href", "mailto:" + address + (subject ? "?subject=" + encodeURIComponent(subject) : ""));
  }
})();
