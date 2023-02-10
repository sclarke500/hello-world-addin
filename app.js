
Office.onReady((info) => {

  // set load on doc open
  Office.addin.setStartupBehavior(Office.StartupBehavior.load);

  // show taskpane - only relevant if being run on doc open
  // -- else taskpane will alredy be visible
  Office.addin.showAsTaskpane();

})
