/**
 * @typedef {GoogleAppsScript.Spreadsheet.Spreadsheet} Tracker
 * @typedef {GoogleAppsScript.Spreadsheet.Sheet} Tab
 * @typedef {string} TabName
 * @typedef {Record<TabName, Tab>} TabMap
 */

const TEMPLATE_TRACKER_ID = '1p3p25O57MnzHDMGSyZsprkPWUb4EyeBAiNCql-li6Uo'
const HUB_NAME = 'Home'

export function updateTracker() {
  const userTracker = SpreadsheetApp.getActiveSpreadsheet()

  if (isTemplateTracker_(userTracker)) {
    throw new Error('Nope!')
  }

  const templateTracker = SpreadsheetApp.openById(TEMPLATE_TRACKER_ID)

  const userTabMap = createTabMap_(userTracker)
  const templateTabMap = createTabMap_(templateTracker)
  const { missingTabs, outdatedTabs } = diffTabs_(userTabMap, templateTabMap)
  const tabsToUpdate = [...missingTabs, ...outdatedTabs]

  if (tabsToUpdate.length === 0) {
    return { status: 'tracker-up-to-date' }
  }

  const copiedTabMap = Object.fromEntries(
    tabsToUpdate.map((tabName) => {
      const templateTab = templateTabMap[tabName]
      return [
        tabName,
        templateTab.copyTo(userTracker)
      ]
    })
  )

  const { tabsWithProgress, tabsWithoutProgress }
    = splitTabsByProgress_(outdatedTabs, userTabMap)
  const trashBin = [...tabsWithoutProgress]
  const skippedAchievements = []

  if (tabsWithProgress.length > 0) {
    backUpTabs_(tabsWithProgress, userTabMap)

    skippedAchievements.push(
      ...transferAchievementProgress_(tabsWithProgress, userTabMap, copiedTabMap)
    )

    trashBin.push(
      ...getBackupsToDelete_(tabsWithProgress, skippedAchievements)
    )
  }

  trashBin.forEach((tabName) => {
    userTracker.deleteSheet(userTabMap[tabName])
  })

  Object.entries(copiedTabMap).forEach(([tabName, tab]) => {
    tab.setName(tabName)
  })

  syncUserTabOrder_(userTracker, templateTabMap)
  updateUserHub_(userTracker, userTabMap, templateTabMap)

  if (skippedAchievements.length > 0) {
    return {
      status: 'tracker-updated-with-warnings',
      skippedAchievements
    }
  }

  return { status: 'tracker-updated' }
}

function isTemplateTracker_(userTracker) {
  return userTracker.getId() === TEMPLATE_TRACKER_ID
}

/**
 * @param {Tracker} tracker
 * @returns {TabMap}
 */
function createTabMap_(tracker) {
  return Object.fromEntries(
    tracker.getSheets().map(tab => [
      tab.getName(),
      tab
    ])
  )
}

/**
 * @param {TabMap} userTabMap
 * @param {TabMap} templateTabMap
 * @returns {{
 *  missingTabs: TabName[],
 *  outdatedTabs: TabName[]
 * }}
 */
function diffTabs_(userTabMap, templateTabMap) {
  const missingTabs = []
  const outdatedTabs = []

  for (const tabName of Object.keys(templateTabMap)) {
    if (tabName === HUB_NAME) {
      continue
    }

    const tab = userTabMap[tabName]
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
  const timestamp = Math.floor(Date.now() / 1000)
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
function transferAchievementProgress_(tabNames, srcTabMap, destTabMap) {
  const skippedAchievements = []

  tabNames.forEach((tabName) => {
    const srcTab = srcTabMap[tabName]
    const destTab = destTabMap[tabName]

    const srcAchievements = getSrcAchievements_(srcTab)
    const destAchievements = getDestAchievements_(destTab)
    const { matched, notMatched }
      = findMatches_(tabName, srcAchievements, destAchievements)

    if (matched.length > 0) {
      destTab.getRangeList(matched).setValue(true)
    }

    skippedAchievements.push(...notMatched)
  })

  return skippedAchievements
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
    && tab.getRange(1, 2, lastRow, 1)
      .getValues()

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
 *  notMatched: Array<{
 *    tabName: TabName,
 *    achievement: string,
 *    srcRow: number
 *  }>
 * }}
 */
function findMatches_(tabName, srcAchievements, destAchievements) {
  const matched = []
  const notMatched = []

  srcAchievements.forEach(({ achievement, srcRow }) => {
    const destRow = destAchievements.get(achievement)
    if (destRow !== undefined) {
      matched.push(`A${destRow}`)
    } else {
      notMatched.push({
        tabName,
        achievement,
        srcRow
      })
    }
  })

  return { matched, notMatched }
}

/**
 * @param {TabName[]} tabNames
 * @param {Array<{
 *   tabName: TabName,
 *   achievement: string,
 *   srcRow: number
 * }>} skippedAchievements
 * @returns {TabName[]}
 */
function getBackupsToDelete_(tabNames, skippedAchievements) {
  const skippedTabs = new Set(
    skippedAchievements.map(({ tabName }) => tabName)
  )

  return tabNames.filter(
    tabName => !skippedTabs.has(tabName)
  )
}

/**
 * @param {Tracker} userTracker
 * @param {TabMap} templateTabMap
 */
function syncUserTabOrder_(userTracker, templateTabMap) {
  const userTabMap = createTabMap_(userTracker)
  Object.keys(templateTabMap).forEach((tabName, index) => {
    if (tabName !== HUB_NAME) {
      userTracker.setActiveSheet(userTabMap[tabName])
      userTracker.moveActiveSheet(index + 1)
    }
  })
}

/**
 * @param {Tracker} userTracker
 * @param {TabMap} userTabMap
 * @param {TabMap} templateTabMap
 */
function updateUserHub_(userTracker, userTabMap, templateTabMap) {
  userTracker.deleteSheet(userTabMap[HUB_NAME])
  const newHub = templateTabMap[HUB_NAME].copyTo(userTracker)
  newHub.setName(HUB_NAME)
  userTracker.setActiveSheet(newHub)
  userTracker.moveActiveSheet(1)
}
