const props = () => PropertiesService.getScriptProperties()
const { engine, viewport } = new Function(fetchBundle_())()

function bootstrapApp() {
  viewport.showSidebar()
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
