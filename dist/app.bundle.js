/* APP_METADATA
{
  "nexusVersion": "1.0.0-alpha.1",
  "bundleVersion": "1.0.0-alpha.1",
  "gameVersion": "7.1"
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
    ui: () => ui_exports
  });

  // src/engine.js
  var engine_exports = {};
  __export(engine_exports, {
    importTracker: () => importTracker,
    updateTracker: () => updateTracker
  });
  var TEMPLATE_TRACKER_ID = "18fYzPwB-_81qwOX1-mYxnEEMM5oo-vCeqhC8LIwxXWw";
  var HUB_NAME = "Home";
  function updateTracker() {
    const userTracker = SpreadsheetApp.getActiveSpreadsheet();
    if (isTemplateTracker_(userTracker)) {
      throw new Error("Nope!");
    }
    const templateTracker = SpreadsheetApp.openById(TEMPLATE_TRACKER_ID);
    const userTabMap = createTabMap_(userTracker);
    const templateTabMap = createTabMap_(templateTracker);
    const { missingTabs, outdatedTabs } = diffTabs_(userTabMap, templateTabMap);
    const tabsToUpdate = [...missingTabs, ...outdatedTabs];
    if (tabsToUpdate.length === 0) {
      return { status: "tracker-up-to-date" };
    }
    const copiedTabMap = Object.fromEntries(
      tabsToUpdate.map((tabName) => {
        const templateTab = templateTabMap[tabName];
        return [
          tabName,
          templateTab.copyTo(userTracker)
        ];
      })
    );
    const { tabsWithProgress, tabsWithoutProgress } = splitTabsByProgress_(outdatedTabs, userTabMap);
    const trashBin = [...tabsWithoutProgress];
    const skippedAchievements = [];
    if (tabsWithProgress.length > 0) {
      backUpTabs_(tabsWithProgress, userTabMap);
      skippedAchievements.push(
        ...transferAchievementProgress_(tabsWithProgress, userTabMap, copiedTabMap)
      );
      trashBin.push(
        ...getBackupsToDelete_(tabsWithProgress, skippedAchievements)
      );
    }
    trashBin.forEach((tabName) => {
      userTracker.deleteSheet(userTabMap[tabName]);
    });
    Object.entries(copiedTabMap).forEach(([tabName, tab]) => {
      tab.setName(tabName);
    });
    syncUserTabOrder_(userTracker, templateTabMap);
    updateUserHub_(userTracker, userTabMap, templateTabMap);
    if (skippedAchievements.length > 0) {
      return {
        status: "tracker-updated-with-warnings",
        skippedAchievements
      };
    }
    return { status: "tracker-updated" };
  }
  function importTracker() {
  }
  function isTemplateTracker_(userTracker) {
    return userTracker.getId() === TEMPLATE_TRACKER_ID;
  }
  function createTabMap_(tracker) {
    return Object.fromEntries(
      tracker.getSheets().map((tab) => [
        tab.getName(),
        tab
      ])
    );
  }
  function diffTabs_(userTabMap, templateTabMap) {
    const missingTabs = [];
    const outdatedTabs = [];
    for (const tabName of Object.keys(templateTabMap)) {
      if (tabName === HUB_NAME) {
        continue;
      }
      const tab = userTabMap[tabName];
      if (!tab) {
        missingTabs.push(tabName);
        continue;
      }
      const templateTab = templateTabMap[tabName];
      const tabVersion = tab.getRange("A1").getValue();
      const templateTabVersion = templateTab.getRange("A1").getValue();
      if (tabVersion !== templateTabVersion) {
        outdatedTabs.push(tabName);
      }
    }
    return {
      missingTabs,
      outdatedTabs
    };
  }
  function splitTabsByProgress_(tabNames, tabMap) {
    const tabsWithProgress = [];
    const tabsWithoutProgress = [];
    tabNames.forEach((tabName) => {
      const tab = tabMap[tabName];
      const lastRow = tab.getLastRow();
      const hasProgress = lastRow > 0 && tab.getRange(1, 1, lastRow, 1).getValues().some(([value]) => value === true);
      if (hasProgress) {
        tabsWithProgress.push(tabName);
      } else {
        tabsWithoutProgress.push(tabName);
      }
    });
    return {
      tabsWithProgress,
      tabsWithoutProgress
    };
  }
  function backUpTabs_(tabNames, tabMap) {
    const timestamp = Math.floor(Date.now() / 1e3);
    tabNames.forEach((tabName) => {
      tabMap[tabName].setName(`${tabName} Backup ${timestamp}`);
    });
  }
  function transferAchievementProgress_(tabNames, srcTabMap, destTabMap) {
    const skippedAchievements = [];
    tabNames.forEach((tabName) => {
      const srcTab = srcTabMap[tabName];
      const destTab = destTabMap[tabName];
      const srcAchievements = getSrcAchievements_(srcTab);
      const destAchievements = getDestAchievements_(destTab);
      const { matched, notMatched } = findMatches_(tabName, srcAchievements, destAchievements);
      if (matched.length > 0) {
        destTab.getRangeList(matched).setValue(true);
      }
      skippedAchievements.push(...notMatched);
    });
    return skippedAchievements;
  }
  function getSrcAchievements_(tab) {
    const columnsAB = tab.getRange(1, 1, tab.getLastRow(), 2).getValues();
    const srcAchievements = [];
    columnsAB.forEach(([isChecked, achievement], index) => {
      if (isChecked === true) {
        srcAchievements.push({
          achievement,
          srcRow: index + 1
        });
      }
    });
    return srcAchievements;
  }
  function getDestAchievements_(tab) {
    const lastRow = tab.getLastRow();
    const columnB = lastRow > 0 ? tab.getRange(1, 2, lastRow, 1).getValues() : [];
    const destAchievements = /* @__PURE__ */ new Map();
    columnB.forEach(([achievement], index) => {
      if (achievement !== "") {
        destAchievements.set(achievement, index + 1);
      }
    });
    return destAchievements;
  }
  function findMatches_(tabName, srcAchievements, destAchievements) {
    const matched = [];
    const notMatched = [];
    srcAchievements.forEach(({ achievement, srcRow }) => {
      const destRow = destAchievements.get(achievement);
      if (destRow !== void 0) {
        matched.push(`A${destRow}`);
      } else {
        notMatched.push({
          tabName,
          achievement,
          srcRow
        });
      }
    });
    return { matched, notMatched };
  }
  function getBackupsToDelete_(tabNames, skippedAchievements) {
    const skippedTabs = new Set(
      skippedAchievements.map(({ tabName }) => tabName)
    );
    return tabNames.filter(
      (tabName) => !skippedTabs.has(tabName)
    );
  }
  function syncUserTabOrder_(userTracker, templateTabMap) {
    const userTabMap = createTabMap_(userTracker);
    Object.keys(templateTabMap).forEach((tabName, index) => {
      if (tabName !== HUB_NAME) {
        userTracker.setActiveSheet(userTabMap[tabName]);
        userTracker.moveActiveSheet(index + 1);
      }
    });
  }
  function updateUserHub_(userTracker, userTabMap, templateTabMap) {
    if (userTabMap[HUB_NAME]) {
      userTracker.deleteSheet(userTabMap[HUB_NAME]);
    }
    const newHub = templateTabMap[HUB_NAME].copyTo(userTracker);
    newHub.setName(HUB_NAME);
    userTracker.setActiveSheet(newHub);
    userTracker.moveActiveSheet(1);
  }

  // src/ui.js
  var ui_exports = {};
  __export(ui_exports, {
    promptForImportUrl: () => promptForImportUrl,
    showModal: () => showModal,
    showSidebar: () => showSidebar
  });

  // src/metadata.json
  var metadata_default = {
    nexusVersion: "1.0.0-alpha.1",
    bundleVersion: "1.0.0-alpha.1",
    gameVersion: "7.1"
  };

  // src/html/sidebar.html
  var sidebar_default = `<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">
  <title>Eternity Sidebar</title>

  <style>
    html,
    body {
      height: 100%;
    }

    body {
      display: flex;
      flex-direction: column;
    }

    main {
      flex: 1;
    }

    .sidebar-footer {
      margin-top: auto;
    }
  </style>

</head>

<body>
  <main>
    <button id="update-tracker">Update tracker</button>
    <button id="import-progress">Import progress</button>

    <script>
      document
        .getElementById('update-tracker')
        .addEventListener('click', updateTracker)

      function updateTracker() {
        google.script.run
          .withSuccessHandler((result) => {
            console.log('SUCCESS:', result)
          })
          .withFailureHandler((error) => {
            console.error('FAILURE:', error)
          })
          .invoke('showModal', 'updateTracker')
      }

      document
        .getElementById('import-progress')
        .addEventListener('click', promptForImportUrl)

      function promptForImportUrl() {
        google.script.run
          .withSuccessHandler((result) => {
            console.log('SUCCESS:', result)
          })
          .withFailureHandler((error) => {
            console.error('FAILURE:', error)
          })
          .invoke('promptForImportUrl')
      }
    </script>

  </main>

  <footer class="sidebar-footer">
    <span>Version
      <?= version ?>
    </span>
  </footer>

</body>

</html>
`;

  // src/html/modal.html
  var modal_default = `<!DOCTYPE html>
<html lang="en">

<head>
  <meta charset="UTF-8">
  <link rel="stylesheet" href="https://ssl.gstatic.com/docs/script/css/add-ons.css">
  <title>Eternity Modal</title>
</head>

<body>
  <div id="method-data" data-payload="<?= payload ?>"></div>
  <p id="result">Loading...</p>

  <script>
    const { method, args } = JSON.parse(document.getElementById('method-data').dataset.payload)
    google.script.run
      .withSuccessHandler((result) => {
        document.getElementById('result').textContent
          = JSON.stringify(result, null, 2)
      })
      .withFailureHandler((error) => {
        document.getElementById('result').textContent = error.message
      })
      .invoke(method, ...args)
  </script>

</body>

</html>
`;

  // src/ui.js
  function showSidebar() {
    const template = HtmlService.createTemplate(sidebar_default);
    template.version = metadata_default.bundleVersion;
    const html = HtmlService.createHtmlOutput(template.evaluate()).setTitle("Eternity Alpha");
    SpreadsheetApp.getUi().showSidebar(html);
  }
  function showModal(method, ...args) {
    const template = HtmlService.createTemplate(modal_default);
    template.payload = JSON.stringify({ method, args });
    const html = template.evaluate();
    SpreadsheetApp.getUi().showModalDialog(html, "Eternity");
  }
  function promptForImportUrl() {
    const ui = SpreadsheetApp.getUi();
    while (true) {
      const response = ui.prompt(
        "Import progress",
        "Paste the URL of your old tracker:",
        ui.ButtonSet.OK_CANCEL
      );
      if (response.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      const id = response.getResponseText().match(/\/d\/([^/]+)\//);
      if (!id) {
        ui.alert("Please provide a valid URL.");
        continue;
      }
      SpreadsheetApp.openById(id[1]);
      return showModal("importProgress", id[1]);
    }
  }
  return __toCommonJS(app_exports);
})();
return App;
