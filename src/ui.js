import metadata from './metadata.json'
import sidebar from './html/sidebar.html'
import modal from './html/modal.html'

export function showSidebar() {
  const template = HtmlService.createTemplate(sidebar)
  template.bundleVersion = metadata.bundleVersion
  template.gameVersion = metadata.gameVersion

  const html = template.evaluate().setTitle('Eternity Alpha')
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
      'Paste a link to your old Wonders of the World checklist:',
      ui.ButtonSet.OK_CANCEL
    )

    if (response.getSelectedButton() !== ui.Button.OK) {
      return
    }

    const trackerId = response.getResponseText().match(/\/d\/([^/]+)\//)?.[1]

    if (!trackerId) {
      ui.alert(
        'Please paste a valid link.\n\n'
        + 'Example:\n'
        + 'https://docs.google.com/spreadsheets/d/18fYzPwB-_81qwOX1-mYxnEEMM5oo-vCeqhC8LIwxXWw/edit?usp=sharing'
      )
      continue
    }

    try {
      SpreadsheetApp.openById(trackerId)
    } catch (error) {
      ui.alert(error.message)
      continue
    }

    return showModal('importProgress', trackerId)
  }
}
