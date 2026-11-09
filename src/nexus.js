const CACHE_KEY = 'APP_BUNDLE'

const cache = () => CacheService.getScriptCache()
const props = () => PropertiesService.getScriptProperties()

let engine
let viewport

function bootstrapApp() {
  clearStaleCache_()
  initApp_()
  viewport.showSidebar()
}

function clearStaleCache_() {
  if (props().getProperty('NEXUS_PURGED') === 'true') {
    cache().remove(CACHE_KEY)
    props().deleteProperty('NEXUS_PURGED')
    console.log('App cache cleared.')
  }
}

function initApp_() {
  const app = new Function(getBundle_())()
  engine = app.engine
  viewport = app.viewport
}

function getBundle_() {
  const cachedBundle = cache().get(CACHE_KEY)
  if (cachedBundle) {
    console.log('App loaded from cache.')
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
