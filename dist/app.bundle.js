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
      width: 100%;
      margin: 0;
      padding: 0;
      overflow: hidden;
    }

    body {
      margin: 0;
      display: flex;
      flex-direction: column;
      background-image:
        url('https://raw.githubusercontent.com/L1M17LE55/eternity/refs/heads/dev/assets/sidebar-background.webp');
      background-size: cover;
      background-position: center bottom;
      background-repeat: no-repeat;
      padding: 16px;
      box-sizing: border-box;
    }

    main {
      flex: 1;
      z-index: 2;
    }

    .sidebar-footer {
      margin-top: auto;
      z-index: 2;
    }
  </style>

  <style>
    #particles {
      position: fixed;
      inset: 0;
      pointer-events: none;
      overflow: hidden;
      z-index: 1;
      perspective: 700px;
    }

    .petal {
      position: absolute;
      top: -20px;
      width: 11px;
      height: 7px;
      background: linear-gradient(135deg, #fff2f9 0%, #ffc2e6 65%, #f495ce 100%);
      border-radius: 80% 15% 75% 25% / 75% 25% 80% 15%;
      box-shadow: 0 0 6px 1px rgba(255, 175, 230, 0.4);
      transform-style: preserve-3d;

      animation:
        fallDown linear infinite,
        horizontalDrift ease-in-out infinite alternate,
        tumble3D ease-in-out infinite alternate;
      will-change: transform, opacity;
    }

    @keyframes fallDown {
      0% {
        transform: translate3d(0, -30px, 0);
        opacity: 0;
      }

      15% {
        opacity: var(--max-opacity, 0.8);
      }

      85% {
        opacity: var(--max-opacity, 0.75);
      }

      100% {
        transform: translate3d(var(--wind-drift, 80px), 108vh, 0);
        opacity: 0;
      }
    }

    @keyframes horizontalDrift {
      0% {
        margin-left: -12px;
      }

      100% {
        margin-left: 12px;
      }
    }

    @keyframes tumble3D {
      0% {
        rotate: 0.8 0.5 0.2 -28deg;
      }

      50% {
        rotate: 0.3 0.9 -0.2 35deg;
      }

      100% {
        rotate: -0.6 0.4 0.5 75deg;
      }
    }
  </style>

</head>

<body>
  <div id="particles"></div>

  <header class="sidebar-header">
    <h1>Eternity</h1>
  </header>

  <main>

    <div class="sidebar-buttons">

      <div class="sidebar-action">
        <button id="update-tracker">Update tracker</button>
        <p>Genshin Impact v
          <?= gameVersion ?> achievements are available!
        </p>
      </div>

      <div class="sidebar-action">
        <button id="import-progress">Import progress</button>
        <p>Transfer your achievement progress from your old Wonders of the World checklist.</p>
      </div>

    </div>

  </main>

  <footer class="sidebar-footer">
    Version
    <?= bundleVersion ?>
  </footer>

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

  <script>
    const container = document.getElementById('particles')
    const petalCount = 9

    for (let i = 0; i < petalCount; i++) {
      const p = document.createElement('span')
      p.className = 'petal'

      p.style.left = (-15 + Math.random() * 110) + '%'

      const depth = Math.random()

      const scale = (0.55 + depth * 0.6).toFixed(2)
      p.style.scale = scale
      p.style.setProperty('--max-opacity', (0.35 + depth * 0.45).toFixed(2))

      if (depth < 0.4) {
        p.style.filter = 'blur(1px)'
      }

      const fallSpeed = (24 - depth * 6 + Math.random() * 4).toFixed(1) + 's'
      const driftSpeed = (4.5 + Math.random() * 3).toFixed(1) + 's'
      const tumbleSpeed = (4.0 + Math.random() * 3.5).toFixed(1) + 's'

      p.style.animationDuration = \`\${fallSpeed}, \${driftSpeed}, \${tumbleSpeed}\`

      const fallDelay = (-Math.random() * 24).toFixed(1) + 's'
      const driftDelay = (-Math.random() * 7).toFixed(1) + 's'
      const tumbleDelay = (-Math.random() * 7).toFixed(1) + 's'
      p.style.animationDelay = \`\${fallDelay}, \${driftDelay}, \${tumbleDelay}\`

      const windDrift = Math.round(45 + depth * 45 + Math.random() * 25) + 'px'
      p.style.setProperty('--wind-drift', windDrift)

      const r1 = Math.round(65 + Math.random() * 20)
      const r2 = Math.round(15 + Math.random() * 20)
      p.style.borderRadius = \`\${r1}% \${r2}% \${r1}% \${r2}%\`

      container.appendChild(p)
    }
  </script>

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
