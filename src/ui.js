import sidebar from './html/sidebar.html'
import modal from './html/modal.html'

export function showSidebar() {
  const html = HtmlService.createHtmlOutput(sidebar).setTitle('Eternity')
  SpreadsheetApp.getUi().showSidebar(html)
}

export function showModal(method, ...args) {
  const template = HtmlService.createTemplate(modal)
  template.payload = JSON.stringify({ method, args })
  const html = template.evaluate()
  SpreadsheetApp.getUi().showModalDialog(html, 'Eternity')
}
