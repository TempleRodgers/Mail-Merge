// Function to update the progress
function updateProgress() {
  CacheService.getScriptCache().put('progress', JSON.stringify(progress), 300); // Store progress in cache for 5 minutes
}

function closeDialog() {
  DocumentApp.getUi().alert('Mail merge completed successfully!');
  DocumentApp.getUi().close();
}

// Function to retrieve progress
function getProgress() {
  var cachedProgress = CacheService.getScriptCache().get('progress');
  return cachedProgress ? JSON.parse(cachedProgress) : progress;
}

function resetProgress() {
  progress.processed = 0;
  progress.total = 0;
  CacheService.getScriptCache().put('progress', JSON.stringify(progress), 300);
}

function getProgressHandler() {
  return getProgress();
}
