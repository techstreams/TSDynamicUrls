/** @OnlyCurrentDoc */

function tsCreateUrlCheatSheet() {
  const ss = SpreadsheetApp.getActive(),
        sheet = ss.getActiveSheet()
        titles = [['DOCS','SHEETS','SLIDES','FORMS','KEEP','CALENDAR','SITES','SCRIPT','JAMBOARD']],
        data = [['doc.new','sheet.new','slide.new','form.new','keep.new','cal.new','site.new','script.new','jam.new'],
                ['docs.new','sheets.new','slides.new','forms.new','note.new','meeting.new','sites.new','',''],
                ['document.new','spreadsheet.new','deck.new','','notes.new','','website.new','',''],
                ['','','presentation.new','','','','','','']];
  let banding;

  sheet.getRange('A1:I1').setValues(titles).setHorizontalAlignment('center')
       .setFontWeight('body')
       .setFontSize(14)
       .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID)
       .activate();
  sheet.getRange('A2:I5').setValues(data).setFontSize(12).activate();
  sheet.getRange('A1:I5').applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
  sheet.getRange('A1:I5').getBandings()[0].setHeaderRowColor('#8bc34a').setFirstRowColor('#ffffff')
                         .setSecondRowColor('#eef7e3').setFooterRowColor(null);
  sheet.getRange('A5:I5')
       .setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  sheet.autoResizeColumns(1, 9);
  sheet.getRange('A7').setValue('For more macros by @techstreams see:').setFontSize(10).activate();
  sheet.getRange('C7').setValue('https://github.com/techstreams/google-sheets-macros')
       .setFontSize(10).activate();
