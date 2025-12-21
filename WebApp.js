/*************************************************
 * Web App Entry
 *************************************************/

function doGet() {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle("HCIS Sabilul Qur'an")
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
