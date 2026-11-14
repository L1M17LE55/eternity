import metadata from './metadata.json'
import sidebar from './html/sidebar.html'
import modal from './html/modal.html'

export function showSidebar() {
  const template = HtmlService.createTemplate(sidebar)
  template.version = metadata.bundleVersion

  const html = HtmlService.createHtmlOutput(template.evaluate())
    .setTitle('Eternity Alpha')
  SpreadsheetApp.getUi().showSidebar(html)
}

export function showModal(method, ...args) {
  const template = HtmlService.createTemplate(modal)
  template.payload = JSON.stringify({ method, args })
  const html = template.evaluate()
  SpreadsheetApp.getUi().showModalDialog(html, 'Eternity')
}

export function promptForImportUrl() {
  const ui = SpreadsheetApp.getUi()

  while (true) {
    const response = ui.prompt(
      'Import progress',
      'Paste the URL of your old tracker:',
      ui.ButtonSet.OK_CANCEL
    )

    if (response.getSelectedButton() !== ui.Button.OK) {
      return
    }

    const id = response.getResponseText().match(/\/d\/([^/]+)\//)

    if (!id) {
      ui.alert('Please provide a valid URL.')
      continue
    }

    SpreadsheetApp.openById(id[1])

    return showModal('importProgress', id[1])
  }
}
