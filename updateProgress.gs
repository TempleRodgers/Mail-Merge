function updateProgress() {
  // Server-side function called to update progress during the merge
  return { processed: progress.processed, total: progress.total };
}

function closeDialog() {
  DocumentApp.getUi().alert('Mail merge completed successfully!');
  DocumentApp.getUi().close();
}

function getProgress() {
  console.log(`current progress = ${progress.processed} and ${progress.total}`);
  return { processed: progress.processed, total: progress.total }; // Return an object with processed and total
}

function resetProgress() {
  progress.processed = 0;
  progress.total = 0;
}
