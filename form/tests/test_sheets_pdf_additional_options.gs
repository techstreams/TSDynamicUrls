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
    testfrozenrows: false,       // Change to 'true' to test Sheets PDF core frozen rows
    testfrozencolumns: false,    // Change to 'true' to test Sheets PDF core frozen columns
    testgridlines: false,        // Change to 'true' to test Sheets PDF core gridlines
    testtitle: false,            // Change to 'true' to test Sheets PDF core ss title
    testsheetnames: false,       // Change to 'true' to test Sheets PDF core sheet names
    testpagenumbers: false,      // Change to 'true' to test Sheets PDF core page numbers
    testnotes: false,            // Change to 'true' to test Sheets PDF core notes
    testnooptions: false,        // Change to 'true' to test Sheets PDF core no additional pdf options
    sheetspdfoptions: {
      input: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/edit?usp=sharing',
      rows: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=false&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      columns: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=false&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      gridlines: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=false&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      title: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=false&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      sheetnames: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=false&pagenum=CENTER&printnotes=true&attachment=true',
      pagenumbers: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&printnotes=true&attachment=true',
      notes: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=false&attachment=true',
      nooptions: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=false&fzc=false&gridlines=false&printtitle=false&sheetnames=false&printnotes=false&attachment=true',
    },
    pdfoptions: {
       rows: ['Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers','Show Notes'],
       columns: ['Repeat Frozen Rows','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers','Show Notes'],
       gridlines: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers','Show Notes'],
       title: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Sheet Names','Show Page Numbers','Show Notes'],
       sheetnames: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Page Numbers','Show Notes'],
       pagenumbers: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Notes'],
       notes: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers'],
    },
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
 

 // Test Sheets - PDF Options
  if (testConfig.testfrozenrows === true) {
    test('Sheets URLs PDF - No Frozen Rows', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Frozen Rows
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.rows, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.rows, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Frozen Rows) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Options
  if (testConfig.testfrozencolumns === true) {
    test('Sheets URLs PDF - No Frozen Columns', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Frozen Columns
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.columns, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.columns, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Frozen Columns) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Options
  if (testConfig.testgridlines === true) {
    test('Sheets URLs PDF - No Gridlines', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Gridlines
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.gridlines, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.gridlines, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Gridlines) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Options
  if (testConfig.testtitle === true) {
    test('Sheets URLs PDF - No Spreadsheet Title', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Spreadsheet Title
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.title, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.title, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Spreadsheet Title) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Options
  if (testConfig.testsheetnames === true) {
    test('Sheets URLs PDF - No Sheet Names', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Sheet Names
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.sheetnames, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.sheetnames, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Sheet Names) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Options
  if (testConfig.testpagenumbers === true) {
    test('Sheets URLs PDF - No Page Numbers', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Page Numbers
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.pagenumbers, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.pagenumbers, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Page Numbers) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
 // Test Sheets - PDF Options
  if (testConfig.testnotes === true) {
    test('Sheets URLs PDF - No Notes', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Notes
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions.notes, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.notes, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Notes) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Options
  if (testConfig.testnooptions === true) {
    test('Sheets URLs PDF - No Additional Options', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - No Additional Options
      TestUtil.createSheetNoopResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfoptions.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top','No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfoptions.nooptions, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with No Additional Options) URL'); 

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
    createSheetNoopResponse: function(form, params) {
      var response = form.createResponse();
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Document Format?',0)).asMultipleChoiceItem().createResponse(params[0]));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Google Sheets Modification Type?',0)).asMultipleChoiceItem().createResponse(params[1])); 
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.LIST, 'Size?',0)).asListItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Orientation?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Scale?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Page Order?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));      
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Horizontal Alignment?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Vertical Alignment?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));
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
