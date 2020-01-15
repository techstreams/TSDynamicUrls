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
    testpagesizeletter: false,       // Change to 'true' to test Sheets PDF page size = Letter
    testpagesizetabloid: false,      // Change to 'true' to test Sheets PDF page size = Tabloid
    testpagesizelegal: false,        // Change to 'true' to test Sheets PDF page size = Legal
    testpagesizestatement: false,    // Change to 'true' to test Sheets PDF page size = Statement
    testpagesizeexecutive: false,    // Change to 'true' to test Sheets PDF page size = Executive
    testpagesizefolio: false,        // Change to 'true' to test Sheets PDF page size = Folio
    testpagesizea3: false,           // Change to 'true' to test Sheets PDF page size = A3
    testpagesizea4: false,           // Change to 'true' to test Sheets PDF page size = A4
    testpagesizea5: false,           // Change to 'true' to test Sheets PDF page size = A5
    testpagesizeb4: false,           // Change to 'true' to test Sheets PDF page size = B4
    testpagesizeb5: false,           // Change to 'true' to test Sheets PDF page size = B5
    sheetspdfpagesize: {
      input: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/edit?usp=sharing',
      letter: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=0&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      tabloid: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=1&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      legal: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=2&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      statement: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=3&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      executive: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=4&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      folio: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=5&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      a3: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=6&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      a4: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=7&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      a5: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=8&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      b4: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=9&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true',
      b5: 'https://docs.google.com/spreadsheets/d/1TxINLwjEkB8QnsaGgKpzvV_VjtFYHVmhPqdB5hFjC3E/export?format=pdf&size=10&portrait=false&scale=1&pageorder=1&horizontal_alignment=CENTER&vertical_alignment=TOP&fzr=true&fzc=true&gridlines=true&printtitle=true&sheetnames=true&pagenum=CENTER&printnotes=true&attachment=true'
  },
    pdfoptions: ['Repeat Frozen Rows','Repeat Frozen Columns','Show Gridlines','Show Spreadsheet Title','Show Sheet Names','Show Page Numbers','Show Notes'],
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
 

 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizeletter === true) {
    test('Sheets URLs PDF - Page Size Letter', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'Letter'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['Letter','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.letter, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = Letter) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizetabloid  === true) {
    test('Sheets URLs PDF - Page Size Tabloid', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'Tabloid'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['Tabloid','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.tabloid, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = Tabloid) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizelegal  === true) {
    test('Sheets URLs PDF - Page Size Legal', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'Legal'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['Legal','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.legal, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = Legal) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizestatement  === true) {
    test('Sheets URLs PDF - Page Size Statement', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'Legal'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['Statement','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.statement, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = Statement) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizeexecutive  === true) {
    test('Sheets URLs PDF - Page Size Executive', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'Executive'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['Executive','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.executive, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = Executive) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizefolio  === true) {
    test('Sheets URLs PDF - Page Size Folio', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'Executive'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['Folio','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.folio, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = Folio) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizea3 === true) {
    test('Sheets URLs PDF - Page Size A3', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'A3'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['A3','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.a3, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = A3) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  

 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizea4 === true) {
    test('Sheets URLs PDF - Page Size A4', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'A4'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['A4','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.a4, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = A4) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizea5 === true) {
    test('Sheets URLs PDF - Page Size A5', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'A5'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['A5','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.a5, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = A5) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizeb4 === true) {
    test('Sheets URLs PDF - Page Size B4', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'B4'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['B4','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.b4, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = B4) URL'); 

      // Re-enable form submit trigger and form email collection after testing
      tsdu.enableForm();

    });
  } 
  
  
 // Test Sheets - PDF Page Size
  if (testConfig.testpagesizeb5 === true) {
    test('Sheets URLs PDF - Page Size B5', function() {
      expect(1);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
          
      // Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction); 
      
      // Test Sheets PDF - All Sheets No Custom Margin - Paper Size = 'B5'
      TestUtil.createSheetResponse(form, ['Sheets', 'PDF (Portable Document Format)', testConfig.sheetspdfpagesize.input, 0, ['B5','Landscape','Normal (100%)','Down, then over','Center','Top',testConfig.pdfoptions, 'No','All Sheets']]);       
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), testConfig.sheetspdfpagesize.b5, 'TSDynamicUrls successfully creates Sheets PDF (All Sheets with Page Size = B5) URL'); 

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
