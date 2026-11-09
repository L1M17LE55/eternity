const CACHE_KEY = 'APP_BUNDLE'

const cache = () => CacheService.getScriptCache()
const props = () => PropertiesService.getScriptProperties()
const { engine, viewport } = new Function(getBundle_())()

function bootstrapApp() {
  clearStaleCache_()
  viewport.showSidebar()
}

function clearStaleCache_() {
  if (!props().getProperty('NEXUS')) {
    cache().remove(CACHE_KEY)
  }
}

function getBundle_() {
  const cachedBundle = cache().get(CACHE_KEY)
  if (cachedBundle) {
    return cachedBundle
  }

  const fetchedBundle = fetchBundle_()
  cache().put(CACHE_KEY, fetchedBundle, 21600)
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
      }
    }
  )
  return response.getContentText()
}

function showModal(method, ...args) {
  viewport.showModal(method, ...args)
}

function helloFromEngine() {
  return engine.helloFromEngine()
}

return {
  bootstrapApp,
  showModal,
  helloFromEngine
}
