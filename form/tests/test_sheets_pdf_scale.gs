/*
  * Copyright Laura Taylor
  * (https://github.com/techstreams/TSDynamicUrls)
  *
  * Permission is hereby granted, free of charge, to any person obtaining a copy
  * of this software and associated documentation files (the "Software"), to deal
  * in the Software without restriction, including without limitation the rights
  * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
  * copies of the Software, and to permit persons to whom the Software is
  * furnished to do so, subject to the following conditions:
  *
  * The above copyright notice and this permission notice shall be included in all
  * copies or substantial portions of the Software.
  *
  * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
  * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
  * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
  * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
  * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
  * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
  * SOFTWARE.
  */

// Setup QUnit helpers with global object
QUnit.helpers(this);


/*
* Create QUnit Testing Dashboard for Google Apps Script
* See https://github.com/simula-innovation/qunit/tree/gas/gas for more information and setup instructions
*
* @param {Object} e - event object passed to doGet() method
* @return {string} html of testing output
*/
function doGet(e) {
  QUnit.urlParams(e.parameter);
  QUnit.config({ title: 'Unit tests for TSDynamicUrls' });
  QUnit.load(cases_);
  return QUnit.getHtml();
};

/*
* TSDynamicUrls Unit Tests Main Function
*
* To perform TSDynamicUrls unit testing, change the desired unit test configuration = "true" in the "testConfig" object below
*/
var cases_ = function() {
  
  // To perform TSDynamicUrls unit testing, change the desired unit configuration parameter to 'true' as indicated below
  var testConfig = {
    testscalenormalwidth: false,       // Change to 'true' to test Sheets PDF scale for normal & fit width
    testscaleheightpage: false,        // Change to 'true' to test Sheets PDF scale for fit height & fit page
    sheetspdfscale: {
      input: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/edit?usp=sharing',
      normal: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      fitwidth: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=2&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      fitheight: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=3&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      fitpage: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=4&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
    },
    pdfoptions: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers','Show Notes'],
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
 

 // Test Sheets - Orientation
  if (testConfig.testscalenormalwidth === true) {
    test('Sheets URLs PDF - Scale', function() {
      expect(2);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Scale = 'Normal'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfscale.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfscale.normal, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Scale = Normal) URL'); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Orientation = 'Landscape'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfscale.input, 0, ['Letter','Landscape','Fit to Width','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfscale.fitwidth, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Scale = Fit to Width) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - Orientation
  if (testConfig.testscaleheightpage === true) {
    test('Sheets URLs PDF - Scale', function() {
      expect(2);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Scale = 'Fit to Height'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfscale.input, 0, ['Letter','Landscape','Fit to Height','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfscale.fitheight, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Scale = Fit to Height) URL'); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Orientation = 'Fit to Page'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfscale.input, 0, ['Letter','Landscape','Fit to Page','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfscale.fitpage, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Scale = Fit to Page) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
    
  /*
  * Various TSDynamic Unit Test Utility Helper Functions
  */
  var TestUtil = {
    createSheetResponse: function(form, params) {
      var response = form.createResponse();
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Document Format?',0)).asMultipleChoiceItem().createResponse(params[0]));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Sheets Modification Type?',0)).asMultipleChoiceItem().createResponse(params[1])); 
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.LIST, 'Size?',0)).asListItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Orientation?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Scale?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Page Order?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));      
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Horizontal Alignment?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Vertical Alignment?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.CHECKBOX, 'Additional PDF Options?',0)).asCheckboxItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Custom Margins?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Include Sheets?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));      
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Google Document URL',params[3])).asTextItem().createResponse(params[2]));
      response.submit();
    },
    deleteTriggers: function(type, functionName) {
      ScriptApp.getProjectTriggers().forEach(function(trigger) {
        if ((trigger.getEventType() === type && trigger.getHandlerFunction() === functionName )) {
          ScriptApp.deleteTrigger(trigger);
        }
      });
    },
    getItemId: function(form, type, name, location) {
      var ids = form.getItems(type).filter(function(item) {
        return item.getTitle() === name;
      });
      return ids[location].getId();
    },
    getTriggers: function(type, functionName) {
      return ScriptApp.getProjectTriggers().filter(function(trigger) {
        if ((trigger.getEventType() === type && trigger.getHandlerFunction() === functionName )) {
          return trigger;
        }
      });
    }
  }
  
  
}
