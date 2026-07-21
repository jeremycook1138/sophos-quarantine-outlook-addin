Office.onReady(() => {
  Office.actions.associate("openSophosQuarantine", openSophosQuarantine);
});

function openSophosQuarantine(event) {
  try {
    Office.context.ui.openBrowserWindow("https://cloud.sophos.com/manage/self-service");
  } finally {
    if (event && typeof event.completed === "function") {
      event.completed();
    }
  }
}
