import metadata from './metadata.json'
import sidebar from './html/sidebar.html'
import modal from './html/modal.html'

export function showSidebar() {
  const template = HtmlService.createTemplate(sidebar)
  template.bundleVersion = metadata.bundleVersion
  template.trackerVersion = metadata.trackerVersion

  const html = template.evaluate().setTitle('\u200B')
  SpreadsheetApp.getUi().showSidebar(html)
}

export function showModal(method, ...args) {
  const template = HtmlService.createTemplate(modal)
  template.payload = JSON.stringify({ method, args })
  const html = template.evaluate()
  SpreadsheetApp.getUi().showModalDialog(html, 'Eternity')
}

export function showImportPrompt() {
  const ui = SpreadsheetApp.getUi()

  while (true) {
    const response = ui.prompt(
      'Transfer your progress',
      'Paste a link to your old Wonders of the World checklist:',
      ui.ButtonSet.OK_CANCEL
    )

    if (response.getSelectedButton() !== ui.Button.OK) {
      return
    }

    const trackerId = response.getResponseText().match(/\/d\/([^/]+)\//)?.[1]

    if (!trackerId) {
      ui.alert(
        `Couldn't recognize link :(`,
        'Please paste a link to Google Sheets.\n\n'
        + 'Example:\n'
        + 'https://docs.google.com/spreadsheets/d/[ID]/edit?usp=sharing',
        ui.ButtonSet.OK
      )
      continue
    }

    try {
      SpreadsheetApp.openById(trackerId)
    } catch (error) {
      ui.alert(error.message)
      continue
    }

    return showModal('executeImportProgress', trackerId)
  }
}
