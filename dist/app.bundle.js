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
  var sidebar_default = `<!DOCTYPE html>\r
<html>\r
\r
<head>\r
  <base target="_top">\r
  <meta charset="UTF-8">\r
  <meta name="viewport" content="width=device-width, initial-scale=1">\r
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">\r
</head>\r
\r
<body>\r
  <button onclick="sayHello()">Hello</button>\r
  <script>\r
    function sayHello() {\r
      google.script.run\r
        .withSuccessHandler((result) => {\r
          console.log('SUCCESS:', result)\r
        })\r
        .withFailureHandler((error) => {\r
          console.error('FAILURE:', error)\r
        })\r
        .invoke('showModal', 'helloFromEngine')\r
    }\r
  </script>\r
</body>\r
\r
</html>\r
`;

  // src/html/modal.html
  var modal_default = `<!DOCTYPE html>\r
<html>\r
\r
<head>\r
  <base target="_top">\r
  <meta charset="UTF-8">\r
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">\r
</head>\r
\r
<body>\r
  <div id="method-data" data-payload="<?= payload ?>"></div>\r
\r
  <p id="result">Loading...</p>\r
  \r
  <script>\r
    const { method, args } = JSON.parse(document.getElementById('method-data').dataset.payload)\r
\r
    google.script.run\r
      .withSuccessHandler((result) => {\r
        document.getElementById('result').textContent = result\r
      })\r
      .withFailureHandler((error) => {\r
        document.getElementById('result').textContent = error.message\r
      })\r
      .invoke(method, ...args)\r
  </script>\r
</body>\r
\r
</html>\r
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
