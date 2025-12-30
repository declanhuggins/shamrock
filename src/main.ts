function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu("SHAMROCK")
    .addItem("Sync Public Views", "shamrockSyncPublicViews")
    .addToUi();
}

function shamrockSyncPublicViews(): void {
  Logger.log("SHAMROCK: sync triggered");
}