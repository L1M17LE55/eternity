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
    importProgress: () => importProgress,
    updateTracker: () => updateTracker
  });
  var TEMPLATE_TRACKER_ID = "18fYzPwB-_81qwOX1-mYxnEEMM5oo-vCeqhC8LIwxXWw";
  var HUB_NAME = "Home";
  function updateTracker() {
    const currentTracker = SpreadsheetApp.getActiveSpreadsheet();
    if (currentTracker.getId() === TEMPLATE_TRACKER_ID) {
      return {
        op: "update",
        status: "current-id-matches-template"
      };
    }
    const templateTracker = SpreadsheetApp.openById(TEMPLATE_TRACKER_ID);
    const currentTabMap = createTabMap_(currentTracker);
    const templateTabMap = createTabMap_(templateTracker);
    const { missingTabs, outdatedTabs } = diffTabs_(currentTabMap, templateTabMap);
    const tabsToUpdate = [...missingTabs, ...outdatedTabs];
    if (tabsToUpdate.length === 0) {
      return {
        op: "update",
        status: "no-updates-available"
      };
    }
    const newTabMap = Object.fromEntries(
      tabsToUpdate.map((tabName) => {
        const templateTab = templateTabMap[tabName];
        return [
          tabName,
          templateTab.copyTo(currentTracker)
        ];
      })
    );
    const { tabsWithProgress, tabsWithoutProgress } = splitTabsByProgress_(outdatedTabs, currentTabMap);
    const trashBin = [...tabsWithoutProgress];
    const untransferredAchievements = [];
    if (tabsWithProgress.length > 0) {
      backUpTabs_(tabsWithProgress, currentTabMap);
      untransferredAchievements.push(
        ...transferAchievements_(tabsWithProgress, currentTabMap, newTabMap)
      );
      trashBin.push(
        ...getBackupsToDelete_(tabsWithProgress, untransferredAchievements)
      );
    }
    trashBin.forEach((tabName) => {
      currentTracker.deleteSheet(currentTabMap[tabName]);
    });
    Object.entries(newTabMap).forEach(([tabName, tab]) => {
      tab.setName(tabName);
    });
    syncTabPosition_(currentTracker, templateTabMap);
    updateHub_(currentTracker, currentTabMap, templateTabMap);
    if (untransferredAchievements.length > 0) {
      return {
        op: "update",
        status: "successful-with-warnings",
        warnings: untransferredAchievements
      };
    }
    return {
      op: "update",
      status: "successful"
    };
  }
  function importProgress(sourceTrackerId) {
    if (sourceTrackerId === TEMPLATE_TRACKER_ID) {
      return {
        op: "import",
        status: "source-id-matches-template"
      };
    }
    const currentTracker = SpreadsheetApp.getActiveSpreadsheet();
    if (sourceTrackerId === currentTracker.getId()) {
      return {
        op: "import",
        status: "source-id-matches-current"
      };
    }
    const versionTabs = getVersionTabsFromTemplate_();
    const sourceTracker = SpreadsheetApp.openById(sourceTrackerId);
    const sourceTabMap = createTabMap_(sourceTracker, versionTabs);
    const sourceTabs = Object.keys(sourceTabMap);
    if (sourceTabs.length === 0) {
      return {
        op: "import",
        status: "no-valid-tabs-in-source"
      };
    }
    const currentTabMap = createTabMap_(currentTracker, sourceTabs);
    const currentTabs = Object.keys(currentTabMap);
    if (currentTabs.length === 0) {
      return {
        op: "import",
        status: "no-valid-tabs-in-current"
      };
    }
    const { tabsWithProgress } = splitTabsByProgress_(currentTabs, sourceTabMap);
    if (tabsWithProgress.length === 0) {
      return {
        op: "import",
        status: "no-import-data-available"
      };
    }
    const untransferredAchievements = transferAchievements_(tabsWithProgress, sourceTabMap, currentTabMap);
    if (untransferredAchievements.length > 0) {
      return {
        op: "import",
        status: "successful-with-warnings",
        warnings: untransferredAchievements
      };
    }
    return {
      op: "import",
      status: "successful"
    };
  }
  function createTabMap_(tracker, allowlist) {
    let tabEntries = tracker.getSheets().map((tab) => [tab.getName(), tab]);
    if (allowlist) {
      tabEntries = tabEntries.filter(([tabName]) => allowlist.includes(tabName));
    }
    return Object.fromEntries(tabEntries);
  }
  function getVersionTabsFromTemplate_() {
    const versionTabs = SpreadsheetApp.openById(TEMPLATE_TRACKER_ID).getSheets().map((tab) => tab.getName()).filter((tabName) => tabName !== HUB_NAME);
    return versionTabs;
  }
  function diffTabs_(currentTabMap, templateTabMap) {
    const missingTabs = [];
    const outdatedTabs = [];
    for (const tabName of Object.keys(templateTabMap)) {
      if (tabName === HUB_NAME) {
        continue;
      }
      const tab = currentTabMap[tabName];
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
    const timestamp = Utilities.formatDate(
      /* @__PURE__ */ new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH-mm-ss"
    );
    tabNames.forEach((tabName) => {
      tabMap[tabName].setName(`${tabName} Backup ${timestamp}`);
    });
  }
  function transferAchievements_(tabNames, srcTabMap, destTabMap) {
    const untransferredAchievements = [];
    tabNames.forEach((tabName) => {
      const srcTab = srcTabMap[tabName];
      const destTab = destTabMap[tabName];
      const srcAchievements = getSrcAchievements_(srcTab);
      const destAchievements = getDestAchievements_(destTab);
      const { matched, unmatched } = findMatches_(tabName, srcAchievements, destAchievements);
      if (matched.length > 0) {
        destTab.getRangeList(matched).setValue(true);
      }
      untransferredAchievements.push(...unmatched);
    });
    return untransferredAchievements;
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
    const unmatched = [];
    srcAchievements.forEach(({ achievement, srcRow }) => {
      const destRow = destAchievements.get(achievement);
      if (destRow !== void 0) {
        matched.push(`A${destRow}`);
      } else {
        unmatched.push({
          tabName,
          achievement,
          srcRow
        });
      }
    });
    return { matched, unmatched };
  }
  function getBackupsToDelete_(backups, untransferredAchievements) {
    const backupsToKeep = new Set(
      untransferredAchievements.map(({ tabName }) => tabName)
    );
    return backups.filter(
      (tabName) => !backupsToKeep.has(tabName)
    );
  }
  function syncTabPosition_(currentTracker, templateTabMap) {
    const currentTabMap = createTabMap_(currentTracker);
    let position = 1;
    Object.keys(templateTabMap).forEach((tabName) => {
      if (tabName !== HUB_NAME) {
        currentTracker.setActiveSheet(currentTabMap[tabName]);
        currentTracker.moveActiveSheet(position++);
      }
    });
  }
  function updateHub_(currentTracker, currentTabMap, templateTabMap) {
    if (currentTabMap[HUB_NAME]) {
      currentTracker.deleteSheet(currentTabMap[HUB_NAME]);
    }
    const newHub = templateTabMap[HUB_NAME].copyTo(currentTracker);
    newHub.setName(HUB_NAME);
    currentTracker.setActiveSheet(newHub);
    currentTracker.moveActiveSheet(1);
  }

  // src/ui.js
  var ui_exports = {};
  __export(ui_exports, {
    showImportPrompt: () => showImportPrompt,
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
  <title>Eternity Sidebar</title>

  <style>
    html,
    body {
      height: 100%;
    }

    body {
      margin: 0;
      display: flex;
      flex-direction: column;
      background-image:
        radial-gradient(circle at 75% 25%, rgba(255, 180, 230, 0.35), transparent 55%),
        radial-gradient(circle at 35% 60%, rgba(255, 220, 150, 0.25), transparent 50%),
        url('https://raw.githubusercontent.com/L1M17LE55/eternity/refs/heads/dev/assets/sidebar-background.webp');
      background-blend-mode: screen, screen, normal;
      background-size: cover;
      background-position: center bottom;
      background-repeat: no-repeat;
      image-rendering: -webkit-optimize-contrast;
      padding: 16px;
      box-sizing: border-box;
    }

    main {
      flex: 1;
    }

    .sidebar-footer {
      margin-top: auto;
      padding-bottom: 8px;
    }
  </style>

</head>

<body>
  <main>
    <div class="sidebar-buttons">
      <div class="sidebar-action">
        <button id="update-tracker">Update tracker</button>
        <p>Genshin Impact <span>v<?= gameVersion ?></span> achievements are available!</p>
      </div>

      <div class="sidebar-action">
        <button id="import-progress">Import progress</button>
        <p>Transfer your achievement progress from your old Wonders of the World checklist.</p>
      </div>
    </div>

    <script>
      function disableButtons(disabled) {
        document.querySelectorAll('button').forEach((button) => {
          button.disabled = disabled
        })
      }

      const updateButton = document.getElementById('update-tracker')
      updateButton.addEventListener('click', () => {
        disableButtons(true)

        google.script.run
          .withSuccessHandler(() => {
            disableButtons(false)
          })
          .withFailureHandler((error) => {
            disableButtons(false)
            console.error(error)
          })
          .invoke('updateTracker')
      })

      const importButton = document.getElementById('import-progress')
      importButton.addEventListener('click', () => {
        disableButtons(true)

        google.script.run
          .withSuccessHandler(() => {
            disableButtons(false)
          })
          .withFailureHandler((error) => {
            disableButtons(false)
            console.error(error)
          })
          .invoke('importProgress')
      })
    </script>

  </main>

  <footer class="sidebar-footer">
    <span>Version <?= bundleVersion ?></span>
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

<!-- if (oldTrackerId === TEMPLATE_TRACKER_ID) {
  throw new Error('Oops! :( You pasted a link to the template. '
    + 'Please try again and paste a link to your old checklist.')
}

const currentTracker = SpreadsheetApp.getActiveSpreadsheet()

if (oldTrackerId === currentTracker.getId()) {
  throw new Error('Oops! :( You pasted a link to your current tracker. '
    + 'Please try again and paste a link to your old checklist.')
} -->
`;

  // src/ui.js
  function showSidebar() {
    const template = HtmlService.createTemplate(sidebar_default);
    template.bundleVersion = metadata_default.bundleVersion;
    template.gameVersion = metadata_default.gameVersion;
    const html = template.evaluate().setTitle("\u200B");
    SpreadsheetApp.getUi().showSidebar(html);
  }
  function showModal(method, ...args) {
    const template = HtmlService.createTemplate(modal_default);
    template.payload = JSON.stringify({ method, args });
    const html = template.evaluate();
    SpreadsheetApp.getUi().showModalDialog(html, "Eternity");
  }
  function showImportPrompt() {
    const ui = SpreadsheetApp.getUi();
    while (true) {
      const response = ui.prompt(
        "Transfer your progress",
        "Paste a link to your old Wonders of the World checklist:",
        ui.ButtonSet.OK_CANCEL
      );
      if (response.getSelectedButton() !== ui.Button.OK) {
        return;
      }
      const trackerId = response.getResponseText().match(/\/d\/([^/]+)\//)?.[1];
      if (!trackerId) {
        ui.alert(
          `Couldn't recognize link :(`,
          "Please paste a link to Google Sheets.\n\nExample:\nhttps://docs.google.com/spreadsheets/d/[ID]/edit?usp=sharing",
          ui.ButtonSet.OK
        );
        continue;
      }
      try {
        SpreadsheetApp.openById(trackerId);
      } catch (error) {
        ui.alert(error.message);
        continue;
      }
      return showModal("executeImportProgress", trackerId);
    }
  }
  return __toCommonJS(app_exports);
})();
return App;
