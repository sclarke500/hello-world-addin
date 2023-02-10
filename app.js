Office.onReady((info) => {
  // do stuff
  Office.addin.showAsTaskpane();
  console.log('tried to show taskpane');

  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

})

function onLoad() {
  console.log('on load');
  // do stuff
}