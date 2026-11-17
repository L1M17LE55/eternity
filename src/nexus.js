const NEXUS_VERSION = '1.0.0-alpha.1'

const CACHE_KEY = 'APP_BUNDLE'

const cache = () => CacheService.getScriptCache()
const props = () => PropertiesService.getScriptProperties()

let appInstance

const app = () => {
  if (!appInstance) {
    appInstance = new Function(getBundle_())()
  }
  return appInstance
}

function bootstrapApp() {
  clearAppCache_()
  app().ui.showSidebar()
}

function clearAppCache_() {
  if (props().getProperty('NEXUS_PURGED') === 'true') {
    cache().remove(CACHE_KEY)
    props().deleteProperty('NEXUS_PURGED')
  }
}

function getBundle_() {
  const cachedBundle = cache().get(CACHE_KEY)
  if (cachedBundle) {
    return cachedBundle
  }

  const fetchedBundle = fetchBundle_()
  cache().put(CACHE_KEY, fetchedBundle, 21600)
  compareNexusVersions_(fetchedBundle)
  return fetchedBundle
}

function fetchBundle_() {
  const response = UrlFetchApp.fetch(
    `https://api.github.com/repos/L1M17LE55/eternity/contents/dist/app.bundle.js?ref=${props().getProperty('BRANCH')}`,
    {
      headers: {
        'Authorization': `Bearer ${props().getProperty('TOKEN')}`,
        'Accept': 'application/vnd.github.raw+json',
        'X-GitHub-Api-Version': '2026-03-10'
      },
      muteHttpExceptions: true
    }
  )

  const statusCode = response.getResponseCode()

  if (statusCode === 404) {
    purgeNexusAndUpdate_()
    throw new Error('BUNDLE_NOT_FOUND')
  }

  if (statusCode !== 200) {
    const ui = SpreadsheetApp.getUi()
    ui.alert(
      'Something went wrong :(',
      'Please try again',
      ui.ButtonSet.OK
    )
    throw new Error(`HTTP_STATUS_${statusCode}`)
  }

  return response.getContentText()
}

function compareNexusVersions_(bundle) {
  const bundleNexusVersion = getMetadata_(bundle).nexusVersion
  if (NEXUS_VERSION !== bundleNexusVersion) {
    purgeNexusAndUpdate_()
    throw new Error('NEXUS_UPDATE_REQUIRED')
  }
}

function getMetadata_(bundle) {
  const metadata = bundle.match(/\/\* APP_METADATA\s*([\s\S]*?)\s*\*\//)
  if (!metadata) {
    purgeNexusAndUpdate_()
    throw new Error('METADATA_NOT_FOUND')
  }
  return JSON.parse(metadata[1])
}

function purgeNexusAndUpdate_() {
  props().deleteProperty('NEXUS')
  cache().remove('NEXUS')
  const ui = SpreadsheetApp.getUi()
  ui.alert(
    'Eternity app updated',
    'Please reopen the app',
    ui.ButtonSet.OK
  )
}

function updateTracker() {
  return app().ui.showModal('updateTracker')
}

function executeUpdateTracker() {
  return app().engine.updateTracker()
}

function showImportPrompt() {
  return app().ui.showImportPrompt()
}

function importProgress(trackerId) {
  return app().engine.importProgress(trackerId)
}

return {
  bootstrapApp,
  updateTracker,
  executeUpdateTracker,
  showImportPrompt,
  importProgress
}
