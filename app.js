Office.onReady((info) => {
  // do stuff
  Office.addin.showAsTaskpane();
  console.log('tried to show taskpane');

  Office.context.document.settings.set("Office.AutoShowTaskpaneWithDocument", true);
  Office.context.document.settings.saveAsync();

})

function onLoad() {
  console.log('on load');
  // do stuff
}