/**
 * @typedef {GoogleAppsScript.Spreadsheet.Spreadsheet} Tracker
 * @typedef {GoogleAppsScript.Spreadsheet.Sheet} Tab
 * @typedef {string} TabName
 * @typedef {Record<TabName, Tab>} TabMap
 */

const TEMPLATE_TRACKER_ID = '18fYzPwB-_81qwOX1-mYxnEEMM5oo-vCeqhC8LIwxXWw'
const HUB_NAME = 'Home'

export function updateTracker() {
  const currentTracker = SpreadsheetApp.getActiveSpreadsheet()

  if (currentTracker.getId() === TEMPLATE_TRACKER_ID) {
    return {
      op: 'update',
      status: 'current-id-matches-template'
    }
  }

  const templateTracker = SpreadsheetApp.openById(TEMPLATE_TRACKER_ID)
  const currentTabMap = createTabMap_(currentTracker)
  const templateTabMap = createTabMap_(templateTracker)
  const { missingTabs, outdatedTabs }
    = diffTabs_(currentTabMap, templateTabMap)
  const tabsToUpdate = [...missingTabs, ...outdatedTabs]

  if (tabsToUpdate.length === 0) {
    return {
      op: 'update',
      status: 'no-updates-available'
    }
  }

  const newTabMap = Object.fromEntries(
    tabsToUpdate.map((tabName) => {
      const templateTab = templateTabMap[tabName]
      return [
        tabName,
        templateTab.copyTo(currentTracker)
      ]
    })
  )

  const { tabsWithProgress, tabsWithoutProgress }
    = splitTabsByProgress_(outdatedTabs, currentTabMap)
  const trashBin = [...tabsWithoutProgress]
  const untransferredAchievements = []

  if (tabsWithProgress.length > 0) {
    backUpTabs_(tabsWithProgress, currentTabMap)
    untransferredAchievements.push(
      ...transferAchievements_(tabsWithProgress, currentTabMap, newTabMap)
    )
    trashBin.push(
      ...getBackupsToDelete_(tabsWithProgress, untransferredAchievements)
    )
  }

  trashBin.forEach((tabName) => {
    currentTracker.deleteSheet(currentTabMap[tabName])
  })

  Object.entries(newTabMap).forEach(([tabName, tab]) => {
    tab.setName(tabName)
  })

  syncTabPosition_(currentTracker, templateTabMap)
  updateHub_(currentTracker, currentTabMap, templateTabMap)

  if (untransferredAchievements.length > 0) {
    return {
      op: 'update',
      status: 'successful-with-warnings',
      warnings: untransferredAchievements
    }
  }

  return {
    op: 'update',
    status: 'successful'
  }
}

export function importProgress(sourceTrackerId) {
  if (sourceTrackerId === TEMPLATE_TRACKER_ID) {
    return {
      op: 'import',
      status: 'source-id-matches-template'
    }
  }

  const currentTracker = SpreadsheetApp.getActiveSpreadsheet()

  if (sourceTrackerId === currentTracker.getId()) {
    return {
      op: 'import',
      status: 'source-id-matches-current'
    }
  }

  const versionTabs = getVersionTabsFromTemplate_()
  const sourceTracker = SpreadsheetApp.openById(sourceTrackerId)
  const sourceTabMap = createTabMap_(sourceTracker, versionTabs)
  const sourceTabs = Object.keys(sourceTabMap)

  if (sourceTabs.length === 0) {
    return {
      op: 'import',
      status: 'no-valid-tabs-in-source'
    }
  }

  const currentTabMap = createTabMap_(currentTracker, sourceTabs)
  const currentTabs = Object.keys(currentTabMap)

  if (currentTabs.length === 0) {
    return {
      op: 'import',
      status: 'no-valid-tabs-in-current'
    }
  }

  const { tabsWithProgress } = splitTabsByProgress_(currentTabs, sourceTabMap)

  if (tabsWithProgress.length === 0) {
    return {
      op: 'import',
      status: 'no-import-data-available'
    }
  }

  const untransferredAchievements = transferAchievements_(tabsWithProgress, sourceTabMap, currentTabMap)

  if (untransferredAchievements.length > 0) {
    return {
      op: 'import',
      status: 'successful-with-warnings',
      warnings: untransferredAchievements
    }
  }

  return {
    op: 'import',
    status: 'successful'
  }
}

/**
 * @param {Tracker} tracker
 * @param {TabName[]} [allowlist]
 * @returns {TabMap}
 */
function createTabMap_(tracker, allowlist) {
  let tabEntries = tracker.getSheets()
    .map(tab => [tab.getName(), tab])

  if (allowlist) {
    tabEntries = tabEntries.filter(([tabName]) => allowlist.includes(tabName))
  }

  return Object.fromEntries(tabEntries)
}

/**
 * @returns {TabName[]}
 */
function getVersionTabsFromTemplate_() {
  const versionTabs = SpreadsheetApp.openById(TEMPLATE_TRACKER_ID)
    .getSheets()
    .map(tab => tab.getName())
    .filter(tabName => tabName !== HUB_NAME)

  return versionTabs
}

/**
 * @param {TabMap} currentTabMap
 * @param {TabMap} templateTabMap
 * @returns {{
 *  missingTabs: TabName[],
 *  outdatedTabs: TabName[]
 * }}
 */
function diffTabs_(currentTabMap, templateTabMap) {
  const missingTabs = []
  const outdatedTabs = []

  for (const tabName of Object.keys(templateTabMap)) {
    if (tabName === HUB_NAME) {
      continue
    }

    const tab = currentTabMap[tabName]

    if (!tab) {
      missingTabs.push(tabName)
      continue
    }

    const templateTab = templateTabMap[tabName]
    const tabVersion = tab.getRange('A1').getValue()
    const templateTabVersion = templateTab.getRange('A1').getValue()

    if (tabVersion !== templateTabVersion) {
      outdatedTabs.push(tabName)
    }
  }

  return {
    missingTabs,
    outdatedTabs
  }
}

/**
 * @param {TabName[]} tabNames
 * @param {TabMap} tabMap
 * @returns {{
 *  tabsWithProgress: TabName[],
 *  tabsWithoutProgress: TabName[]
 * }}
 */
function splitTabsByProgress_(tabNames, tabMap) {
  const tabsWithProgress = []
  const tabsWithoutProgress = []

  tabNames.forEach((tabName) => {
    const tab = tabMap[tabName]
    const lastRow = tab.getLastRow()

    const hasProgress = lastRow > 0
      && tab.getRange(1, 1, lastRow, 1)
        .getValues()
        .some(([value]) => value === true)

    if (hasProgress) {
      tabsWithProgress.push(tabName)
    } else {
      tabsWithoutProgress.push(tabName)
    }
  })

  return {
    tabsWithProgress,
    tabsWithoutProgress
  }
}

/**
 * @param {TabName[]} tabNames
 * @param {TabMap} tabMap
 */
function backUpTabs_(tabNames, tabMap) {
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'yyyy-MM-dd HH-mm-ss'
  )

  tabNames.forEach((tabName) => {
    tabMap[tabName].setName(`${tabName} Backup ${timestamp}`)
  })
}

/**
 * @param {TabName[]} tabNames
 * @param {TabMap} srcTabMap
 * @param {TabMap} destTabMap
 * @returns {Array<{
 *  tabName: TabName,
 *  achievement: string,
 *  srcRow: number
 * }>}
 */
function transferAchievements_(tabNames, srcTabMap, destTabMap) {
  const untransferredAchievements = []

  tabNames.forEach((tabName) => {
    const srcTab = srcTabMap[tabName]
    const destTab = destTabMap[tabName]
    const srcAchievements = getSrcAchievements_(srcTab)
    const destAchievements = getDestAchievements_(destTab)
    const { matched, unmatched }
      = findMatches_(tabName, srcAchievements, destAchievements)

    if (matched.length > 0) {
      destTab.getRangeList(matched).setValue(true)
    }

    untransferredAchievements.push(...unmatched)
  })

  return untransferredAchievements
}

/**
 * @param {Tab} tab
 * @returns {Array<{ achievement: string, srcRow: number }>}
 */
function getSrcAchievements_(tab) {
  const columnsAB = tab
    .getRange(1, 1, tab.getLastRow(), 2)
    .getValues()
  const srcAchievements = []

  columnsAB.forEach(([isChecked, achievement], index) => {
    if (isChecked === true) {
      srcAchievements.push({
        achievement,
        srcRow: index + 1
      })
    }
  })

  return srcAchievements
}

/**
 * @param {Tab} tab
 * @returns {Map<string, number>}
 */
function getDestAchievements_(tab) {
  const lastRow = tab.getLastRow()
  const columnB = lastRow > 0
    ? tab.getRange(1, 2, lastRow, 1).getValues()
    : []
  const destAchievements = new Map()

  columnB.forEach(([achievement], index) => {
    if (achievement !== '') {
      destAchievements.set(achievement, index + 1)
    }
  })

  return destAchievements
}

/**
 * @param {TabName} tabName
 * @param {Array<{ achievement: string, srcRow: number }>} srcAchievements
 * @param {Map<string, number>} destAchievements
 * @returns {{
 *  matched: string[],
 *  unmatched: Array<{
 *    tabName: TabName,
 *    achievement: string,
 *    srcRow: number
 *  }>
 * }}
 */
function findMatches_(tabName, srcAchievements, destAchievements) {
  const matched = []
  const unmatched = []

  srcAchievements.forEach(({ achievement, srcRow }) => {
    const destRow = destAchievements.get(achievement)

    if (destRow !== undefined) {
      matched.push(`A${destRow}`)
    } else {
      unmatched.push({
        tabName,
        achievement,
        srcRow
      })
    }
  })

  return { matched, unmatched }
}

/**
 * @param {TabName[]} backups
 * @param {Array<{
 *   tabName: TabName,
 *   achievement: string,
 *   srcRow: number
 * }>} untransferredAchievements
 * @returns {TabName[]}
 */
function getBackupsToDelete_(backups, untransferredAchievements) {
  const backupsToKeep = new Set(
    untransferredAchievements.map(({ tabName }) => tabName)
  )

  return backups.filter(
    tabName => !backupsToKeep.has(tabName)
  )
}

/**
 * @param {Tracker} currentTracker
 * @param {TabMap} templateTabMap
 */
function syncTabPosition_(currentTracker, templateTabMap) {
  const currentTabMap = createTabMap_(currentTracker)
  let position = 1

  Object.keys(templateTabMap).forEach((tabName) => {
    if (tabName !== HUB_NAME) {
      currentTracker.setActiveSheet(currentTabMap[tabName])
      currentTracker.moveActiveSheet(position++)
    }
  })
}

/**
 * @param {Tracker} currentTracker
 * @param {TabMap} currentTabMap
 * @param {TabMap} templateTabMap
 */
function updateHub_(currentTracker, currentTabMap, templateTabMap) {
  if (currentTabMap[HUB_NAME]) {
    currentTracker.deleteSheet(currentTabMap[HUB_NAME])
  }

  const newHub = templateTabMap[HUB_NAME].copyTo(currentTracker)

  newHub.setName(HUB_NAME)
  currentTracker.setActiveSheet(newHub)
  currentTracker.moveActiveSheet(1)
}
