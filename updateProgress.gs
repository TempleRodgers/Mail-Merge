function updateProgress(processed, total) {
  // Calls the function inside the HTML to update the progress
  DocumentApp.getUi().getActive().run(function() {
    google.script.run.updateProgress(processed, total);
  });
}

function closeDialog() {
  DocumentApp.getUi().alert('Mail merge completed successfully!');
  DocumentApp.getUi().close();
}

function getProgress() {
  console.log(`current progress = ${progress.processed} and ${progress.total}`);
  return progress;
}

function resetProgress() {
  progress.processed = 0;
  progress.total = 0;
}
