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
    testsheetsnopdf: false,      // Change to 'true' to test Sheets URLs (no pdf)
    formSubmitFunction: 'processResponses',
    email: Session.getActiveUser().getEmail()
  };
  
  
  module('TSDynamicUrls');
  
  // Test Sheets - No PDF
  if (testConfig.testsheetsnopdf === true) {
    test('Sheets URLs (No PDF)', function() {
      expect(11);
      
      var tsdu = null,
          form = FormApp.getActiveForm(),
          responses, response;
      var sheets = {
        input: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/edit?usp=sharing',
        idinput: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/edit#gid=929929678',
        preview: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/preview',
        copy: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/copy',
        copycomments: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/copy?copyComments=true',
        template: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/template/preview',
        html: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=zip',
        xlsx: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=xlsx',
        ods: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=ods',
        csvnoid: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=csv&gid=0',
        csvid: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=csv&gid=929929678',
        tsvnoid: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=tsv&gid=0',
        tsvid: 'https://docs.google.com/spreadsheets/d/1VwLdHIZHxE8kKBZ72e1LEwrlm2WtI9kgar6Z1iSPtW4/export?format=tsv&gid=929929678'
      };
      
      // Temporarily Disable form submit trigger for testing
      TestUtil.deleteTriggers(ScriptApp.EventType.ON_FORM_SUBMIT, testConfig.formSubmitFunction);
      
      // Test Sheets Preview
      TestUtil.createSheetResponse(form, ['Sheets', 'Preview', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.preview, 'TSDynamicUrls successfully creates Sheets Preview URL');
      
      // Test Sheets Copy
      TestUtil.createSheetResponse(form, ['Sheets', 'Copy', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.copy, 'TSDynamicUrls successfully creates Sheets Copy URL');
      
      // Test Sheets Copy With Comments
      TestUtil.createSheetResponse(form, ['Sheets', 'Copy with Comments', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.copycomments, 'TSDynamicUrls successfully creates Sheets Copy with Comments URL');
      
      // Test Sheets Template
      TestUtil.createSheetResponse(form, ['Sheets', 'Template', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.template, 'TSDynamicUrls successfully creates Sheets Template URL'); 
      
      // Test Sheets HTML
      TestUtil.createSheetResponse(form, ['Sheets', 'HTML (Web Page, Zipped)', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.html, 'TSDynamicUrls successfully creates Sheets HTML URL');
      
      // Test Sheets XLSX
      TestUtil.createSheetResponse(form, ['Sheets', 'XLSX (Microsoft Excel)', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.xlsx, 'TSDynamicUrls successfully creates Sheets XLSX URL');
      
      // Test Sheets ODS
      TestUtil.createSheetResponse(form, ['Sheets', 'ODS (OpenDocument Spreadsheet)', sheets.input, 0]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.ods, 'TSDynamicUrls successfully creates Sheets ODS URL');
      
      // Test Sheets CSV No Sheet Id - Default First Sheet
      TestUtil.createSheetResponse(form, ['Sheets', 'CSV (Comma-separated values)  - Individual Sheet Only', sheets.input, 1]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.csvnoid, 'TSDynamicUrls successfully creates Sheets CSV No Sheet Id URL');
      
      // Test Sheets CSV Sheet Id
      TestUtil.createSheetResponse(form, ['Sheets', 'CSV (Comma-separated values)  - Individual Sheet Only', sheets.idinput, 1]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.csvid, 'TSDynamicUrls successfully creates Sheets CSV Sheet Id URL');
      
      // Test Sheets TSV No Sheet Id - Default First Sheet
      TestUtil.createSheetResponse(form, ['Sheets', 'TSV (Tab-separated values)  - Individual Sheet Only', sheets.input, 1]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.tsvnoid, 'TSDynamicUrls successfully creates Sheets TSV No Sheet Id URL');
      
      // Test Sheets TSV Sheet Id
      TestUtil.createSheetResponse(form, ['Sheets', 'TSV (Tab-separated values)  - Individual Sheet Only', sheets.idinput, 1]);
      responses = form.getResponses();
      response = responses[responses.length-1];
      tsdu = new TSDynamicUrls(form, response).convertUrl();
      equal(tsdu.getUrl(), sheets.tsvid, 'TSDynamicUrls successfully creates Sheets TSV Sheet Id URL');
      
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
