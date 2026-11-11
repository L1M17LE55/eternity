/* APP_METADATA
{
  "nexusVersion": "1.0.0-alpha.1",
  "bundleVersion": "1.0.0-alpha.1"
}
*/
var App = (() => {
  var __defProp = Object.defineProperty;
  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
  var __getOwnPropNames = Object.getOwnPropertyNames;
  var __hasOwnProp = Object.prototype.hasOwnProperty;
  var __export = (target, all) => {
    for (var name in all)
      __defProp(target, name, { get: all[name], enumerable: true });
  };
  var __copyProps = (to, from, except, desc) => {
    if (from && typeof from === "object" || typeof from === "function") {
      for (let key of __getOwnPropNames(from))
        if (!__hasOwnProp.call(to, key) && key !== except)
          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
    }
    return to;
  };
  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

  // src/app.js
  var app_exports = {};
  __export(app_exports, {
    engine: () => engine_exports,
    viewport: () => viewport_exports
  });

  // src/engine.js
  var engine_exports = {};
  __export(engine_exports, {
    helloFromEngine: () => helloFromEngine
  });
  function helloFromEngine() {
    return "Engine says hello";
  }

  // src/viewport.js
  var viewport_exports = {};
  __export(viewport_exports, {
    showModal: () => showModal,
    showSidebar: () => showSidebar
  });

  // src/html/sidebar.html
  var sidebar_default = `<!DOCTYPE html>
<html>

<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">
</head>

<body>
  <button onclick="sayHello()">Hello</button>
  <script>
    function sayHello() {
      google.script.run
        .withSuccessHandler((result) => {
          console.log('SUCCESS:', result)
        })
        .withFailureHandler((error) => {
          console.error('FAILURE:', error)
        })
        .invoke('showModal', 'helloFromEngine')
    }
  </script>
</body>

</html>
`;

  // src/html/modal.html
  var modal_default = `<!DOCTYPE html>
<html>

<head>
  <base target="_top">
  <meta charset="UTF-8">
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">
</head>

<body>
  <div id="method-data" data-payload="<?= payload ?>"></div>

  <p id="result">Loading...</p>
  
  <script>
    const { method, args } = JSON.parse(document.getElementById('method-data').dataset.payload)

    google.script.run
      .withSuccessHandler((result) => {
        document.getElementById('result').textContent = result
      })
      .withFailureHandler((error) => {
        document.getElementById('result').textContent = error.message
      })
      .invoke(method, ...args)
  </script>
</body>

</html>
`;

  // src/viewport.js
  function showSidebar() {
    const html = HtmlService.createHtmlOutput(sidebar_default).setTitle("Eternity");
    SpreadsheetApp.getUi().showSidebar(html);
  }
  function showModal(method, ...args) {
    const template = HtmlService.createTemplate(modal_default);
    template.payload = JSON.stringify({ method, args });
    const html = template.evaluate();
    SpreadsheetApp.getUi().showModalDialog(html, "Eternity");
  }
  return __toCommonJS(app_exports);
})();
return App;
