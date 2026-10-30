function bootstrapApp() {
  SpreadsheetApp.getUi().alert(
    'Hello, world',
    SpreadsheetApp.getUi().ButtonSet.OK
  )
  return bootstrapApp
}

return { bootstrapApp }
