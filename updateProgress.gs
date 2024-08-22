function updateProgress(processed, total) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('processed', processed);
  userProperties.setProperty('total', total);
}

function getProgress() {
  var userProperties = PropertiesService.getUserProperties();
  return {
    processed: parseInt(userProperties.getProperty('processed'), 10) || 0,
    total: parseInt(userProperties.getProperty('total'), 10) || 0
  };
}
