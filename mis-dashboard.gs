function wpfData() {
  const sheet = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1pulnY_UukPGPoc-vr6bTvhh4ppn3k-V0J9Jx8z42zWk/").getActiveSheet();
  let currentRow = sheet.getLastRow();
  let formData = sheet.getRange(currentRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const template = SlidesApp.openByUrl("https://docs.google.com/presentation/d/1nQkovd77j883JaWldwS5v8XcQHMH0ptUBsHI-wJ_ubo/");
  const final = SlidesApp.openByUrl("https://docs.google.com/presentation/d/1mHTPXZkI6JARTS24TKsO8mzxPcUb_vc_8i2084kTFOY/");

  let finalPages = final.getSlides();
  for (let i = finalPages.length - 1; i >= 0; i--) {
    finalPages[i].remove();            //    Clear the PPT from previous values
  }
  let templatePages = template.getSlides();
  for(let i=0; i<templatePages.length; i++) {
    let templatePage = templatePages[i];
    let newPage = final.appendSlide(SlidesApp.PredefinedLayout.BLANK);
    newPage.insertTextBox(formData[6], 100, 100, 400, 50);
  }
}
