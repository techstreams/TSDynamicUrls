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
    testsinglesheet: false,       // Change to 'true' to test Sheets PDF Single Sheet
    testnamedrange: false,        // Change to 'true' to test Sheets PDF Named Range
    testcustomrange: false,       // Change to 'true' to test Sheets PDF Custom Range
    sheetspdfsinglesheet: {
      input1: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/edit?usp=sharing',
      input2: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/edit#gid=929929678',
      sheet1: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true&gid=0',
      sheet2: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true&gid=929929678',
      namedrange: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true&range=CapitalSubset&gid=929929678',
      customrange: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true&gid=929929678&ir=false&ic=false&r1=0&c1=0&r2=10&c2=4',
    },
    pdfoptions: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers','Show Notes'],
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
 

 // Test Sheets - Single Sheet No Range
  if (testConfig.testsinglesheet === true) {
    test('Sheets URLs PDF - Test Single Sheet No Sheet ID', function() {
      expect(2);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - Single Sheet No Custom Margin No Sheet ID
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfsinglesheet.input1, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','Specific Sheet','Entire Sheet']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfsinglesheet.sheet1, 'TSDynamicUrls successfully creates Sheets PDF (Single Sheet No Sheet ID) URL'); 
      
      // Test Sheets PDF - Single Sheet No Custom Margin With Sheet ID
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfsinglesheet.input2, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','Specific Sheet','Entire Sheet']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfsinglesheet.sheet2, 'TSDynamicUrls successfully creates Sheets PDF (Single Sheet With Sheet ID) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - Single Sheet Named Range
  if (testConfig.testnamedrange === true) {
    test('Sheets URLs PDF - Test Single Sheet Named Range', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - Single Sheet No Custom Margin With Named Range
      TestUtil.createSheetNamedRangeResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfsinglesheet.input2, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','Specific Sheet','Specific Range','Named Range','CapitalSubset']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfsinglesheet.namedrange, 'TSDynamicUrls successfully creates Sheets PDF (Single Sheet with Named Range) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - Single Sheet No Range
  if (testConfig.testcustomrange === true) {
    test('Sheets URLs PDF - Test Single Sheet Custom Range', function() {
      expect(2);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - Single Sheet With Valid Custom Range
      TestUtil.createSheetCustomRangeResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfsinglesheet.input2, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','Specific Sheet','Specific Range','Custom Range','1','1','10','4']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfsinglesheet.customrange, 'TSDynamicUrls successfully creates Sheets PDF (Single Sheet with Valid Custom Range) URL'); 
      
      // Test Sheets PDF - Single Sheet With Invalid Custom Range
      TestUtil.createSheetCustomRangeResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfsinglesheet.input2, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','Specific Sheet','Specific Range','Custom Range','2','2','1','1']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfsinglesheet.sheet2, 'TSDynamicUrls successfully creates Sheets PDF (Single Sheet with Invalid Custom Range) URL'); 

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
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Include Range?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));   
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Google Document URL',params[3])).asTextItem().createResponse(params[2]));
      response.submit();
    },
    createSheetNamedRangeResponse: function(form, params) {
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
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Include Range?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));  
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Range Type?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));  
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Named Range?',0)).asTextItem().createResponse(params[4].shift())); 
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Google Document URL',params[3])).asTextItem().createResponse(params[2]));
      response.submit();
    },
    createSheetCustomRangeResponse: function(form, params) {
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
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Include Range?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));  
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.MULTIPLE_CHOICE, 'Range Type?',0)).asMultipleChoiceItem().createResponse(params[4].shift()));  
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Range Start Row?',0)).asTextItem().createResponse(params[4].shift())); 
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Range Start Column?',0)).asTextItem().createResponse(params[4].shift())); 
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Range End Row?',0)).asTextItem().createResponse(params[4].shift())); 
      response.withItemResponse(form.getItemById(TestUtil.getItemId(form, FormApp.ItemType.TEXT, 'Range End Column?',0)).asTextItem().createResponse(params[4].shift())); 
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
