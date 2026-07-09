Office.onReady(() => {
  // Office.js is ready.
});

function openSophosQuarantine(event) {
  Office.ui.openBrowserWindow("https://cloud.sophos.com/manage/self-service");

  if (event && typeof event.completed === "function") {
    event.completed();
  }
}

Office.actions.associate("openSophosQuarantine", openSophosQuarantine);
